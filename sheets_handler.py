"""
sheets_handler.py
Handles all Google Sheets operations — reading holdings, updating prices,
adding/removing stocks, updating P&L sheet.
Replaces excel_reader.py, excel_updater.py, pnl_updater.py Excel operations.
"""

import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import yfinance as yf
from config import GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE, NSE_SUFFIX

SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]


def get_client():
    """Authenticate and return gspread client"""
    creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def get_sheets():
    """Return Holdings and PnL worksheets"""
    client   = get_client()
    sheet    = client.open_by_key(GOOGLE_SHEET_ID)
    holdings = sheet.worksheet("Holdings")
    pnl      = sheet.worksheet("PnL")
    return holdings, pnl


# ── READ HOLDINGS ─────────────────────────────────────────────
def read_holdings():
    """Read all stocks from Holdings sheet"""
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()

        holdings = []
        for row in records:
            ticker = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            if not ticker:
                continue

            try:
                buying_price = float(str(row.get('Buying Price', 0)).replace('₹','').replace(',','').strip())
            except:
                continue

            try:
                qty = int(float(str(row.get('Qty', 0)).replace(',','').strip()))
            except:
                continue

            buying_date = str(row.get('Buying Date', '')).strip()
            if not buying_date:
                buying_date = datetime.now().strftime('%Y-%m-%d')

            industry = str(row.get('Industry', 'N/A')).strip()

            holdings.append({
                'ticker':       ticker,
                'ticker_yf':    ticker + NSE_SUFFIX,
                'stock_name':   ticker,
                'industry':     industry,
                'buying_price': buying_price,
                'buying_date':  buying_date,
                'qty':          qty,
                'current_price': None,
                'growth_pct':   None,
            })

        print(f"[sheets] Loaded {len(holdings)} holdings from Google Sheets.")
        return holdings

    except Exception as e:
        print(f"[sheets] Error reading holdings: {e}")
        return []


# ── UPDATE LIVE PRICES ────────────────────────────────────────
def get_live_price(ticker):
    try:
        info  = yf.Ticker(ticker + NSE_SUFFIX).info
        price = info.get('currentPrice') or info.get('regularMarketPrice')
        return round(float(price), 2) if price else None
    except:
        return None


def update_holdings_prices():
    """Update Current Share Price column in Holdings sheet"""
    try:
        holdings_sheet, _ = get_sheets()
        records   = holdings_sheet.get_all_records()
        headers   = holdings_sheet.row_values(1)

        # Find column index for Current Share Price
        try:
            price_col = headers.index('Current Share Price') + 1
        except ValueError:
            print("[sheets] 'Current Share Price' column not found in Holdings!")
            return False

        updated = 0
        for i, row in enumerate(records, start=2):
            ticker = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            if not ticker:
                continue
            price = get_live_price(ticker)
            if price:
                holdings_sheet.update_cell(i, price_col, price)
                print(f"[sheets] {ticker}: Rs.{price}")
                updated += 1

        print(f"[sheets] ✅ Updated {updated} prices in Holdings sheet")
        return True

    except Exception as e:
        print(f"[sheets] Error updating prices: {e}")
        return False


def update_pnl_prices():
    """Update Current Share Price column in PnL sheet for sold stocks"""
    try:
        _, pnl_sheet = get_sheets()
        records  = pnl_sheet.get_all_records()

        if not records:
            print("[sheets] PnL sheet has no data — skipping.")
            return True

        headers = pnl_sheet.row_values(1)
        try:
            price_col = headers.index('Current Share Price') + 1
        except ValueError:
            print("[sheets] 'Current Share Price' column not found in PnL!")
            return False

        updated = 0
        for i, row in enumerate(records, start=2):
            ticker = str(row.get('Ticker', '')).strip().upper()
            if not ticker:
                continue
            price = get_live_price(ticker)
            if price:
                pnl_sheet.update_cell(i, price_col, price)
                print(f"[sheets] PnL {ticker}: Rs.{price}")
                updated += 1

        if updated > 0:
            print(f"[sheets] ✅ Updated {updated} prices in PnL sheet")
        else:
            print("[sheets] No stocks to update in PnL sheet")
        return True

    except Exception as e:
        print(f"[sheets] Error updating PnL prices: {e}")
        return False


# ── ADD NEW STOCK ─────────────────────────────────────────────
def add_stock_to_holdings(ticker, stock_name, industry, buying_price, buying_date, qty):
    """Add a newly approved stock to Holdings sheet"""
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        sno     = len(records) + 1
        investment_amt = round(buying_price * qty, 2)

        new_row = [
            sno,
            industry,
            ticker,
            stock_name,
            buying_price,
            buying_date,
            qty,
            investment_amt,
            '',   # Current Share Price — updated by updater
            '',   # Profit per share — formula
            '',   # Total Profit — formula
            '',   # Networth — formula
            '',   # Growth — formula
            '',   # Investment Days — formula
        ]

        holdings_sheet.append_row(new_row)
        new_row_num = sno + 1
        print(f"[sheets] ✅ Added {ticker} to Holdings sheet at row {new_row_num}")
        # Add formulas to new row
        check_and_add_formulas_new_row('holdings', new_row_num)
        return sno

    except Exception as e:
        print(f"[sheets] Error adding stock: {e}")
        return None


# ── REMOVE STOCK (on sell) ────────────────────────────────────
def remove_stock_from_holdings(ticker, buying_price, buying_date):
    """Remove a sold stock from Holdings sheet and renumber S.no"""
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()

        row_to_delete = None
        for i, row in enumerate(records, start=2):
            t  = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            bp = float(str(row.get('Buying Price', 0)).replace('₹','').replace(',',''))
            bd = str(row.get('Buying Date', ''))
            if t == ticker.upper() and abs(bp - buying_price) < 0.01 and str(buying_date)[:7] in bd:
                row_to_delete = i
                break

        if row_to_delete:
            holdings_sheet.delete_rows(row_to_delete)
            print(f"[sheets] ✅ Removed {ticker} from Holdings sheet")

            # Renumber S.no
            records = holdings_sheet.get_all_records()
            for i, row in enumerate(records, start=2):
                holdings_sheet.update_cell(i, 1, i - 1)
            print(f"[sheets] ✅ S.no renumbered")
            return True
        else:
            print(f"[sheets] {ticker} not found in Holdings sheet")
            return False

    except Exception as e:
        print(f"[sheets] Error removing stock: {e}")
        return False


# ── ADD TO P&L ────────────────────────────────────────────────
def add_to_pnl(stock, selling_price, selling_date):
    """Add sold stock record to PnL sheet — checks for duplicates first"""
    try:
        _, pnl_sheet = get_sheets()
        records = pnl_sheet.get_all_records()

        # Check for duplicate — same ticker + same buying date
        for row in records:
            existing_ticker = str(row.get('Ticker', '')).strip().upper()
            existing_date   = str(row.get('Buying Date', '')).strip()[:10]
            if existing_ticker == stock['ticker'].upper() and existing_date == str(stock['buying_date'])[:10]:
                print(f"[sheets] ⚠️ {stock['ticker']} already in PnL — skipping duplicate")
                return len(records)

        sno = len(records) + 1

        buying_price    = float(stock['buying_price'])
        qty             = int(stock['qty'])
        buying_date     = str(stock['buying_date'])[:10]
        investment_amt  = round(buying_price * qty, 2)
        profit_per_share = round(selling_price - buying_price, 2)
        total_profit    = round(profit_per_share * qty, 2)
        return_pct      = round(((selling_price - buying_price) / buying_price) * 100, 2)

        try:
            bd = datetime.strptime(buying_date, '%Y-%m-%d')
            sd = datetime.strptime(selling_date, '%Y-%m-%d')
            investment_days = (sd - bd).days
            time_months     = round(investment_days / 30.44, 1)
        except:
            investment_days = 0
            time_months     = 0

        current_price = get_live_price(stock['ticker'])
        if current_price and selling_price:
            current_return = round(((current_price - selling_price) / selling_price) * 100, 2)
        else:
            current_price  = selling_price
            current_return = 0.0

        new_row = [
            sno,
            stock.get('industry', ''),
            stock.get('ticker', ''),
            stock.get('stock_name', ''),
            buying_date,
            selling_date,
            buying_price,
            float(selling_price),
            qty,
            investment_amt,
            profit_per_share,
            total_profit,
            return_pct,
            investment_days,
            current_price,
            current_return,
            time_months,
        ]

        pnl_sheet.append_row(new_row)
        new_row_num = sno + 1
        print(f"[sheets] ✅ Added {stock['ticker']} to PnL sheet at row {new_row_num}")
        # Add formulas to new row
        check_and_add_formulas_new_row('pnl', new_row_num)
        return sno

    except Exception as e:
        print(f"[sheets] Error adding to PnL: {e}")
        return None


def setup_holdings_formulas():
    """Add formulas to Holdings sheet columns H,J,K,M,N — run once"""
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        if not records:
            print("[sheets] No holdings data to add formulas to")
            return

        for i in range(2, len(records) + 2):  # Start from row 2
            # H = Investment amt = E*G (Buying Price * Qty)
            holdings_sheet.update_cell(i, 8,  f'=E{i}*G{i}')
            # J = Profit per share = I-E (CSP - Buying Price)
            holdings_sheet.update_cell(i, 10, f'=I{i}-E{i}')
            # K = Total Profit = J*G
            holdings_sheet.update_cell(i, 11, f'=J{i}*G{i}')
            # L = Growth = ((I-E)/E)*100 (Networth column removed, shifted left)
            holdings_sheet.update_cell(i, 12, f'=IF(E{i}=0,0,((I{i}-E{i})/E{i})*100)')
            # M = Investment Days = TODAY() - Buying Date
            holdings_sheet.update_cell(i, 13, f'=IF(F{i}="",0,DAYS(TODAY(),F{i}))')
            print(f"[sheets] ✅ Formulas added to Holdings row {i}")

        print("[sheets] ✅ Holdings formulas setup complete!")
    except Exception as e:
        print(f"[sheets] Error setting up Holdings formulas: {e}")


def setup_pnl_formulas():
    """Add formulas to PnL sheet — run once"""
    try:
        _, pnl_sheet = get_sheets()
        records = pnl_sheet.get_all_records()
        if not records:
            print("[sheets] No PnL data to add formulas to")
            return

        for i in range(2, len(records) + 2):
            # J = Investment amt = G*I (Buying Price * Qty)
            pnl_sheet.update_cell(i, 10, f'=G{i}*I{i}')
            # K = Profit per share = H-G (Selling Price - Buying Price)
            pnl_sheet.update_cell(i, 11, f'=H{i}-G{i}')
            # L = Total Profit = K*I
            pnl_sheet.update_cell(i, 12, f'=K{i}*I{i}')
            # M = Return % = ((H-G)/G)*100
            pnl_sheet.update_cell(i, 13, f'=IF(G{i}=0,0,((H{i}-G{i})/G{i})*100)')
            # N = Investment Days = DAYS(Selling Date, Buying Date)
            pnl_sheet.update_cell(i, 14, f'=IF(E{i}="",0,DAYS(F{i},E{i}))')
            # P = Current Return = ((O-H)/H)*100
            pnl_sheet.update_cell(i, 16, f'=IF(H{i}=0,0,((O{i}-H{i})/H{i})*100)')
            # Q = Time in months = DATEDIF(Selling Date, TODAY(), "m")
            pnl_sheet.update_cell(i, 17, f'=IFERROR(DATEDIF(F{i},TODAY(),"m"),0)')
            print(f"[sheets] ✅ Formulas added to PnL row {i}")

        print("[sheets] ✅ PnL formulas setup complete!")
    except Exception as e:
        print(f"[sheets] Error setting up PnL formulas: {e}")


def check_and_add_formulas_new_row(sheet_type, row_num):
    """Add formulas to a newly added row"""
    try:
        holdings_sheet, pnl_sheet = get_sheets()
        i = row_num
        if sheet_type == 'holdings':
            holdings_sheet.update_cell(i, 8,  f'=E{i}*G{i}')
            holdings_sheet.update_cell(i, 10, f'=I{i}-E{i}')
            holdings_sheet.update_cell(i, 11, f'=J{i}*G{i}')
            holdings_sheet.update_cell(i, 12, f'=IF(E{i}=0,0,((I{i}-E{i})/E{i})*100)')
            holdings_sheet.update_cell(i, 13, f'=IF(F{i}="",0,DAYS(TODAY(),F{i}))')
        elif sheet_type == 'pnl':
            pnl_sheet.update_cell(i, 10, f'=G{i}*I{i}')
            pnl_sheet.update_cell(i, 11, f'=H{i}-G{i}')
            pnl_sheet.update_cell(i, 12, f'=K{i}*I{i}')
            pnl_sheet.update_cell(i, 13, f'=IF(G{i}=0,0,((H{i}-G{i})/G{i})*100)')
            pnl_sheet.update_cell(i, 14, f'=IF(E{i}="",0,DAYS(F{i},E{i}))')
            pnl_sheet.update_cell(i, 16, f'=IF(H{i}=0,0,((O{i}-H{i})/H{i})*100)')
            pnl_sheet.update_cell(i, 17, f'=IFERROR(DATEDIF(F{i},TODAY(),"m"),0)')
    except Exception as e:
        print(f"[sheets] Error adding formulas to row {i}: {e}")


if __name__ == "__main__":
    import sys
    if '--setup-formulas' in sys.argv:
        print("Setting up Holdings formulas...")
        setup_holdings_formulas()
        print("Setting up PnL formulas...")
        setup_pnl_formulas()
    else:
        print("Testing Google Sheets connection...")
        holdings = read_holdings()
        for h in holdings:
            print(f"  {h['ticker']} | Buy: {h['buying_price']} | Qty: {h['qty']}")
