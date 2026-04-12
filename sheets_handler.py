"""
sheets_handler.py
All Google Sheets operations — extended with:
  - FLAW 1.4: check_and_fix_stock_splits()
  - FLAW 2.1: log_sentiment_history(), get_sentiment_history()
  - FLAW 2.4: log_recommendation(), get_recommendations_log()
  - FLAW 3.3: sector field in add_stock_to_holdings()
  - FLAW 3.4: flag_pending_sell(), get_pending_sells(), clear_pending_sell_flag()
  - FLAW 5.3: export_backup_to_json() for Sunday GitHub backup
  - Helpers: get_pnl_records(), get_sheet_client(), get_or_create_worksheet()
"""

import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import yfinance as yf
import json
from config import GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE, NSE_SUFFIX

SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]


def get_client():
    creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def get_sheet_client():
    return get_client()


def get_sheets():
    client   = get_client()
    sheet    = client.open_by_key(GOOGLE_SHEET_ID)
    holdings = sheet.worksheet("Holdings")
    pnl      = sheet.worksheet("PnL")
    return holdings, pnl


def get_or_create_worksheet(name):
    """Get worksheet by name, auto-create with headers if it does not exist"""
    try:
        client = get_client()
        sheet  = client.open_by_key(GOOGLE_SHEET_ID)
        try:
            return sheet.worksheet(name)
        except gspread.WorksheetNotFound:
            ws = sheet.add_worksheet(title=name, rows=1000, cols=20)
            print(f"[sheets] Created new worksheet: {name}")
            return ws
    except Exception as e:
        print(f"[sheets] Error getting/creating worksheet {name}: {e}")
        return None


# ── READ HOLDINGS ──────────────────────────────────────────────────────────────

def read_holdings():
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

            try:
                target_price = float(str(row.get('Target Price', '') or '').replace(',','').strip() or 0) or None
            except:
                target_price = None

            # Read cap_category if column exists
            cap_category = str(row.get('Cap Category', '') or '').strip() or None

            holdings.append({
                'ticker':        ticker,
                'ticker_yf':     ticker + NSE_SUFFIX,
                'stock_name':    ticker,
                'industry':      industry,
                'buying_price':  buying_price,
                'buying_date':   buying_date,
                'qty':           qty,
                'current_price': None,
                'growth_pct':    None,
                'target_price':  target_price,
                'cap_category':  cap_category,
            })

        print(f"[sheets] Loaded {len(holdings)} holdings.")
        return holdings

    except Exception as e:
        print(f"[sheets] Error reading holdings: {e}")
        return []


# ── LIVE PRICE ─────────────────────────────────────────────────────────────────

def get_live_price(ticker):
    try:
        info  = yf.Ticker(ticker + NSE_SUFFIX).info
        price = info.get('currentPrice') or info.get('regularMarketPrice')
        return round(float(price), 2) if price else None
    except:
        return None


def update_holdings_prices():
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        headers = holdings_sheet.row_values(1)

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

        print(f"[sheets] Updated {updated} prices in Holdings")
        return True

    except Exception as e:
        print(f"[sheets] Error updating prices: {e}")
        return False


def update_pnl_prices():
    try:
        _, pnl_sheet = get_sheets()
        records = pnl_sheet.get_all_records()
        if not records:
            return True

        headers = pnl_sheet.row_values(1)
        try:
            price_col = headers.index('Current Share Price') + 1
        except ValueError:
            print("[sheets] 'Current Share Price' not found in PnL!")
            return False

        updated = 0
        for i, row in enumerate(records, start=2):
            ticker = str(row.get('Ticker', '')).strip().upper()
            if not ticker:
                continue
            price = get_live_price(ticker)
            if price:
                pnl_sheet.update_cell(i, price_col, price)
                updated += 1

        if updated > 0:
            print(f"[sheets] Updated {updated} PnL prices")
        return True

    except Exception as e:
        print(f"[sheets] Error updating PnL prices: {e}")
        return False


# ── ADD NEW STOCK ──────────────────────────────────────────────────────────────

def add_stock_to_holdings(ticker, stock_name, industry, buying_price, buying_date,
                          qty, target_price=None, cap_category=None, sector=None):
    """
    Add a newly approved stock to Holdings sheet.
    FLAW 3.3: Now also accepts cap_category and sector parameters.
    """
    try:
        holdings_sheet, _ = get_sheets()
        records        = holdings_sheet.get_all_records()
        sno            = len(records) + 1
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
            '',   # Current Share Price
            '',   # Profit per share
            '',   # Total Profit
            '',   # Growth
            '',   # Investment Days
            target_price if target_price else '',
            cap_category if cap_category else '',   # Cap Category column
            sector if sector else industry,          # Sector column
        ]

        holdings_sheet.append_row(new_row)
        new_row_num = sno + 1
        print(f"[sheets] Added {ticker} to Holdings at row {new_row_num}")
        check_and_add_formulas_new_row('holdings', new_row_num)
        return sno

    except Exception as e:
        print(f"[sheets] Error adding stock: {e}")
        return None


# ── REMOVE STOCK ───────────────────────────────────────────────────────────────

def remove_stock_from_holdings(ticker, buying_price, buying_date):
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
            print(f"[sheets] Removed {ticker} from Holdings")
            records = holdings_sheet.get_all_records()
            for i, row in enumerate(records, start=2):
                holdings_sheet.update_cell(i, 1, i - 1)
            return True

        print(f"[sheets] {ticker} not found in Holdings")
        return False

    except Exception as e:
        print(f"[sheets] Error removing stock: {e}")
        return False


# ── ADD TO P&L ─────────────────────────────────────────────────────────────────

def add_to_pnl(stock, selling_price, selling_date):
    try:
        _, pnl_sheet = get_sheets()
        records = pnl_sheet.get_all_records()

        # Duplicate check
        for row in records:
            et = str(row.get('Ticker', '')).strip().upper()
            ed = str(row.get('Buying Date', '')).strip()[:10]
            if et == stock['ticker'].upper() and ed == str(stock['buying_date'])[:10]:
                print(f"[sheets] {stock['ticker']} already in PnL — skipping duplicate")
                return len(records)

        sno              = len(records) + 1
        buying_price     = float(stock['buying_price'])
        qty              = int(stock['qty'])
        buying_date      = str(stock['buying_date'])[:10]
        investment_amt   = round(buying_price * qty, 2)
        profit_per_share = round(float(selling_price) - buying_price, 2)
        total_profit     = round(profit_per_share * qty, 2)
        return_pct       = round(((float(selling_price) - buying_price) / buying_price) * 100, 2)

        try:
            bd = datetime.strptime(buying_date, '%Y-%m-%d')
            sd = datetime.strptime(str(selling_date)[:10], '%Y-%m-%d')
            investment_days = (sd - bd).days
            time_months     = round(investment_days / 30.44, 1)
        except:
            investment_days = 0
            time_months     = 0

        current_price = get_live_price(stock['ticker'])
        if current_price and selling_price:
            current_return = round(((current_price - float(selling_price)) / float(selling_price)) * 100, 2)
        else:
            current_price  = float(selling_price)
            current_return = 0.0

        new_row = [
            sno,
            stock.get('industry', ''),
            stock.get('ticker', ''),
            stock.get('stock_name', ''),
            buying_date,
            str(selling_date)[:10],
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
        print(f"[sheets] Added {stock['ticker']} to PnL at row {new_row_num}")
        check_and_add_formulas_new_row('pnl', new_row_num)
        return sno

    except Exception as e:
        print(f"[sheets] Error adding to PnL: {e}")
        return None


def get_pnl_records():
    """Return all PnL records as list of dicts (used by stock_scout for Flaw 3.5)"""
    try:
        _, pnl_sheet = get_sheets()
        return pnl_sheet.get_all_records()
    except Exception as e:
        print(f"[sheets] Error reading PnL: {e}")
        return []


# ── FLAW 1.4: STOCK SPLIT DETECTION & AUTO-CORRECTION ─────────────────────────

def check_and_fix_stock_splits():
    """
    For each holding, compare stored price vs yfinance current price.
    If difference > 40%, check yfinance splits history for today/yesterday.
    If split confirmed: auto-correct qty and buying price, send notification email.
    """
    from email_handler import send_report_email

    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        headers = holdings_sheet.row_values(1)

        try:
            price_col = headers.index('Current Share Price') + 1
        except ValueError:
            return

        splits_detected = []

        for i, row in enumerate(records, start=2):
            ticker = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            if not ticker:
                continue

            try:
                stored_price = float(str(row.get('Current Share Price', 0) or 0).replace('₹','').replace(',',''))
                if stored_price <= 0:
                    continue

                ticker_yf = ticker + NSE_SUFFIX
                stock_obj = yf.Ticker(ticker_yf)
                hist      = stock_obj.history(period="3d")

                if hist.empty:
                    continue

                current_yf_price = round(float(hist['Close'].iloc[-1]), 2)
                diff_pct = abs((current_yf_price - stored_price) / stored_price) * 100

                if diff_pct > 40:
                    print(f"[sheets] {ticker}: Price diff {diff_pct:.1f}% — checking splits...")

                    splits      = stock_obj.splits
                    today_str   = date.today().strftime('%Y-%m-%d')
                    # Check today and yesterday to handle yfinance 1-day lag
                    from datetime import timedelta
                    yest_str    = (date.today() - timedelta(days=1)).strftime('%Y-%m-%d')
                    split_ratio = None
                    split_date  = None

                    for split_dt, ratio in splits.items():
                        split_dt_str = str(split_dt)[:10]
                        if split_dt_str in [today_str, yest_str]:
                            split_ratio = float(ratio)
                            split_date  = split_dt_str
                            break

                    if split_ratio and split_ratio > 0:
                        old_qty       = int(float(str(row.get('Qty', 0)).replace(',','')))
                        old_buy_price = float(str(row.get('Buying Price', 0)).replace('₹','').replace(',',''))
                        new_qty       = int(old_qty * split_ratio)
                        new_buy_price = round(old_buy_price / split_ratio, 2)

                        try:
                            qty_col       = headers.index('Qty') + 1
                            buy_price_col = headers.index('Buying Price') + 1
                        except ValueError:
                            print(f"[sheets] Column not found for {ticker} split correction")
                            continue

                        holdings_sheet.update_cell(i, qty_col, new_qty)
                        holdings_sheet.update_cell(i, buy_price_col, new_buy_price)

                        # Flag as SPLIT ADJUSTED for the day (pause auto-sell)
                        try:
                            split_col = headers.index('Split Adjusted') + 1
                            holdings_sheet.update_cell(i, split_col, today_str)
                        except ValueError:
                            pass  # Column may not exist yet — that is fine

                        print(f"[sheets] {ticker} split corrected: {old_qty} -> {new_qty} shares, Rs.{old_buy_price} -> Rs.{new_buy_price}")

                        splits_detected.append({
                            'ticker':        ticker,
                            'split_ratio':   split_ratio,
                            'split_date':    split_date,
                            'old_qty':       old_qty,
                            'new_qty':       new_qty,
                            'old_buy_price': old_buy_price,
                            'new_buy_price': new_buy_price,
                            'investment_amt': round(old_buy_price * old_qty, 2),
                        })
                    else:
                        print(f"[sheets] {ticker}: Large price diff but no split found — flagging for manual review.")

            except Exception as e:
                print(f"[sheets] Split check error for {ticker}: {e}")

        if splits_detected:
            _send_split_notification_email(splits_detected)

    except Exception as e:
        print(f"[sheets] check_and_fix_stock_splits error: {e}")


def _send_split_notification_email(splits):
    from email_handler import send_report_email

    cards = ''
    for s in splits:
        cards += f"""
        <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;border-left:5px solid #1976D2;">
          <h2 style="margin:0 0 12px;">{s['ticker']} — Stock Split 1:{int(s['split_ratio'])}</h2>
          <table style="width:100%;border-collapse:collapse;">
            <tr style="background:#f8f9fa;"><td style="padding:8px 14px;font-weight:bold;">Split Date</td><td style="padding:8px 14px;">{s['split_date']}</td></tr>
            <tr><td style="padding:8px 14px;font-weight:bold;">Split Ratio</td><td style="padding:8px 14px;font-size:18px;font-weight:bold;color:#1976D2;">1:{int(s['split_ratio'])}</td></tr>
            <tr style="background:#f8f9fa;"><td style="padding:8px 14px;font-weight:bold;">Shares Before</td><td style="padding:8px 14px;">{s['old_qty']}</td></tr>
            <tr><td style="padding:8px 14px;font-weight:bold;">Shares After</td><td style="padding:8px 14px;font-size:18px;font-weight:bold;color:#2E7D32;">{s['new_qty']}</td></tr>
            <tr style="background:#f8f9fa;"><td style="padding:8px 14px;font-weight:bold;">Buy Price Before</td><td style="padding:8px 14px;">Rs.{s['old_buy_price']:,.2f}</td></tr>
            <tr><td style="padding:8px 14px;font-weight:bold;">Buy Price After</td><td style="padding:8px 14px;">Rs.{s['new_buy_price']:,.2f}</td></tr>
            <tr style="background:#E8F5E9;"><td style="padding:8px 14px;font-weight:bold;">Investment Amount</td><td style="padding:8px 14px;font-weight:bold;color:#2E7D32;">Rs.{s['investment_amt']:,.2f} (unchanged)</td></tr>
          </table>
          <div style="background:#E3F2FD;padding:10px;border-radius:6px;margin-top:12px;font-size:12px;">
            Google Sheets updated automatically. Auto-sell paused today for this stock.
          </div>
        </div>"""

    html    = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:700px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#0D47A1,#1976D2);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;">Stock Split Detected & Auto-Corrected</h1>
    <p style="margin:6px 0 0;opacity:0.8;">{len(splits)} stock(s) adjusted in Google Sheets</p>
  </div>
  {cards}
</body></html>"""
    subject = f"Stock Split Auto-Corrected: {', '.join(s['ticker'] for s in splits)}"
    send_report_email(subject, html)
    print(f"[sheets] Split notification email sent.")


def is_split_adjusted_today(ticker):
    """Check if a stock was split-adjusted today — used to pause auto-sell"""
    try:
        holdings_sheet, _ = get_sheets()
        records   = holdings_sheet.get_all_records()
        today_str = date.today().strftime('%Y-%m-%d')
        for row in records:
            t = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            if t == ticker.upper():
                flag = str(row.get('Split Adjusted', '') or '')
                return flag == today_str
        return False
    except:
        return False


# ── FLAW 3.4: PENDING SELL FLAG ────────────────────────────────────────────────

def flag_pending_sell(ticker, decision_price, target_hit=False):
    """
    Mark a stock as PENDING_SELL in Holdings sheet.
    Actual sell executes next morning at open price via --morning job.
    """
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        headers = holdings_sheet.row_values(1)

        for i, row in enumerate(records, start=2):
            t = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            if t == ticker.upper():
                try:
                    col = headers.index('Pending Sell') + 1
                    holdings_sheet.update_cell(i, col, 'YES')
                except ValueError:
                    pass
                try:
                    col = headers.index('Decision Price') + 1
                    holdings_sheet.update_cell(i, col, decision_price)
                except ValueError:
                    pass
                try:
                    col = headers.index('Target Hit') + 1
                    holdings_sheet.update_cell(i, col, 'YES' if target_hit else 'NO')
                except ValueError:
                    pass
                print(f"[sheets] {ticker} flagged as PENDING_SELL at Rs.{decision_price}")
                return True

        print(f"[sheets] {ticker} not found for PENDING_SELL flag")
        return False

    except Exception as e:
        print(f"[sheets] Error flagging pending sell: {e}")
        return False


def get_pending_sells():
    """Return list of stocks with PENDING_SELL flag set"""
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        pending = []

        for row in records:
            if str(row.get('Pending Sell', '') or '').upper() != 'YES':
                continue

            ticker = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            try:
                buying_price = float(str(row.get('Buying Price', 0)).replace('₹','').replace(',',''))
                qty          = int(float(str(row.get('Qty', 0)).replace(',','')))
            except:
                continue

            pending.append({
                'ticker':             ticker,
                'stock_name':         str(row.get('Stock Name', ticker)),
                'industry':           str(row.get('Industry', 'N/A')),
                'buying_price':       buying_price,
                'buying_date':        str(row.get('Buying Date', '')),
                'qty':                qty,
                'pending_sell_price': float(str(row.get('Decision Price', 0) or 0)),
                'target_hit':         str(row.get('Target Hit', 'NO')).upper() == 'YES',
            })

        if pending:
            print(f"[sheets] Found {len(pending)} pending sell(s): {[p['ticker'] for p in pending]}")
        return pending

    except Exception as e:
        print(f"[sheets] Error getting pending sells: {e}")
        return []


def clear_pending_sell_flag(ticker, buying_price, buying_date):
    """Clear PENDING_SELL flag after morning execution"""
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        headers = holdings_sheet.row_values(1)

        for i, row in enumerate(records, start=2):
            t  = str(row.get('Stock', '') or row.get('Ticker', '')).strip().upper()
            bp = float(str(row.get('Buying Price', 0)).replace('₹','').replace(',',''))
            if t == ticker.upper() and abs(bp - buying_price) < 0.01:
                try:
                    col = headers.index('Pending Sell') + 1
                    holdings_sheet.update_cell(i, col, '')
                except ValueError:
                    pass
                return True
        return False
    except Exception as e:
        print(f"[sheets] Error clearing pending sell: {e}")
        return False


# ── FLAW 2.1: SENTIMENT HISTORY ────────────────────────────────────────────────

def log_sentiment_history(ticker, verdict, run_date):
    """Store today's Groq verdict in SentimentHistory sheet"""
    try:
        ws = get_or_create_worksheet("SentimentHistory")
        if ws is None:
            return

        existing = ws.get_all_values()
        if not existing or existing[0] != ['Date', 'Ticker', 'Verdict']:
            ws.insert_row(['Date', 'Ticker', 'Verdict'], 1)

        ws.append_row([run_date, ticker.upper(), verdict.upper()])

    except Exception as e:
        print(f"[sheets] Error logging sentiment history: {e}")


def get_sentiment_history(ticker, days=5):
    """Get last N verdicts for a stock from SentimentHistory sheet"""
    try:
        ws = get_or_create_worksheet("SentimentHistory")
        if ws is None:
            return []

        records = ws.get_all_records()
        ticker_records = [
            r for r in records
            if str(r.get('Ticker', '')).upper() == ticker.upper()
        ]

        ticker_records.sort(key=lambda x: str(x.get('Date', '')), reverse=True)
        recent = ticker_records[:days]
        recent.reverse()  # chronological order for Groq prompt

        return [{'date': r['Date'], 'verdict': r['Verdict']} for r in recent]

    except Exception as e:
        print(f"[sheets] Error fetching sentiment history for {ticker}: {e}")
        return []


# ── FLAW 2.4: RECOMMENDATIONS LOG ──────────────────────────────────────────────

def log_recommendation(ticker, recommended_price, target_price, recommended_date):
    """Log each top-5 recommendation for accuracy tracking"""
    try:
        ws = get_or_create_worksheet("RecommendationsLog")
        if ws is None:
            return

        existing = ws.get_all_values()
        if not existing or (existing[0] and existing[0][0] != 'Date'):
            ws.insert_row(['Date', 'Ticker', 'Recommended Price', 'Target Price', 'Status'], 1)

        ws.append_row([
            recommended_date,
            ticker.upper(),
            recommended_price,
            target_price if target_price else '',
            'OPEN'
        ])
        print(f"[sheets] Logged recommendation: {ticker} @ Rs.{recommended_price}")

    except Exception as e:
        print(f"[sheets] Error logging recommendation: {e}")


def get_recommendations_log():
    """Get all recommendation records"""
    try:
        ws = get_or_create_worksheet("RecommendationsLog")
        if ws is None:
            return []
        return ws.get_all_records()
    except Exception as e:
        print(f"[sheets] Error reading recommendations log: {e}")
        return []


# ── FLAW 5.3: WEEKLY BACKUP ─────────────────────────────────────────────────────

def export_backup_to_json():
    """Export all sheets to JSON and save as backup file for GitHub repo"""
    try:
        client  = get_client()
        sheet   = client.open_by_key(GOOGLE_SHEET_ID)
        backup  = {}
        today   = date.today().strftime('%Y-%m-%d')

        for ws in sheet.worksheets():
            try:
                backup[ws.title] = ws.get_all_records()
            except Exception as e:
                print(f"[sheets] Could not export {ws.title}: {e}")
                backup[ws.title] = []

        filename = f"backup_{today}.json"
        with open(filename, 'w') as f:
            json.dump({'exported_on': today, 'sheets': backup}, f, default=str, indent=2)

        print(f"[sheets] Backup exported to {filename}")
        return filename

    except Exception as e:
        print(f"[sheets] Backup error: {e}")
        return None


# ── FORMULA HELPERS ────────────────────────────────────────────────────────────

def setup_holdings_formulas():
    try:
        holdings_sheet, _ = get_sheets()
        records = holdings_sheet.get_all_records()
        if not records:
            return
        for i in range(2, len(records) + 2):
            holdings_sheet.update_cell(i, 8,  f'=E{i}*G{i}')
            holdings_sheet.update_cell(i, 10, f'=I{i}-E{i}')
            holdings_sheet.update_cell(i, 11, f'=J{i}*G{i}')
            holdings_sheet.update_cell(i, 12, f'=IF(E{i}=0,0,((I{i}-E{i})/E{i})*100)')
            holdings_sheet.update_cell(i, 13, f'=IF(F{i}="",0,DAYS(TODAY(),F{i}))')
            print(f"[sheets] Formulas added to Holdings row {i}")
        print("[sheets] Holdings formulas setup complete!")
    except Exception as e:
        print(f"[sheets] Error setting up Holdings formulas: {e}")


def setup_pnl_formulas():
    try:
        _, pnl_sheet = get_sheets()
        records = pnl_sheet.get_all_records()
        if not records:
            return
        for i in range(2, len(records) + 2):
            pnl_sheet.update_cell(i, 10, f'=G{i}*I{i}')
            pnl_sheet.update_cell(i, 11, f'=H{i}-G{i}')
            pnl_sheet.update_cell(i, 12, f'=K{i}*I{i}')
            pnl_sheet.update_cell(i, 13, f'=IF(G{i}=0,0,((H{i}-G{i})/G{i})*100)')
            pnl_sheet.update_cell(i, 14, f'=IF(E{i}="",0,DAYS(F{i},E{i}))')
            pnl_sheet.update_cell(i, 16, f'=IF(H{i}=0,0,((O{i}-H{i})/H{i})*100)')
            pnl_sheet.update_cell(i, 17, f'=IFERROR(DATEDIF(F{i},TODAY(),"m"),0)')
            print(f"[sheets] Formulas added to PnL row {i}")
        print("[sheets] PnL formulas setup complete!")
    except Exception as e:
        print(f"[sheets] Error setting up PnL formulas: {e}")


def check_and_add_formulas_new_row(sheet_type, row_num):
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
        print(f"[sheets] Error adding formulas to row {row_num}: {e}")


if __name__ == "__main__":
    import sys
    if '--setup-formulas' in sys.argv:
        setup_holdings_formulas()
        setup_pnl_formulas()
    elif '--backup' in sys.argv:
        export_backup_to_json()
    elif '--split-check' in sys.argv:
        check_and_fix_stock_splits()
    else:
        print("Testing Google Sheets connection...")
        holdings = read_holdings()
        for h in holdings:
            print(f"  {h['ticker']} | Buy: {h['buying_price']} | Qty: {h['qty']}")