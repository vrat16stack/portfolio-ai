"""
approval_checker.py
Processes YES/NO email replies for stock additions.
  - FLAW 3.2: Position sizing warning if stock > 15% of portfolio value
  - FLAW 3.3: Sector concentration warning if sector would exceed 30%
"""

import imaplib
import email
import re
from datetime import datetime
import yfinance as yf
from config import EMAIL_SENDER, EMAIL_PASSWORD, NSE_SUFFIX
from email_handler import send_report_email

try:
    from config import EXCEL_FILE_PATH, HOLDINGS_SHEET
except ImportError:
    EXCEL_FILE_PATH = ""
    HOLDINGS_SHEET  = "Sheet1"

try:
    from add_stock_macro import add_stock_via_vba
except ImportError:
    def add_stock_via_vba(ticker, buying_price, buying_date, qty):
        return True, 1


def fetch_approval_replies():
    approvals_with_qty   = []
    approvals_need_qty   = []
    rejections           = []
    confirmed_overrides  = []  # "YES SYRMA 10 CONFIRM" — override concentration warning

    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(EMAIL_SENDER, EMAIL_PASSWORD)
        mail.select('inbox')
        _, messages = mail.search(None, 'UNSEEN')

        for msg_id in messages[0].split():
            _, msg_data = mail.fetch(msg_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])

            body = ''
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        break
            else:
                body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

            first_line = body.strip().split('\n')[0].strip().upper()
            print(f"[approval] Reading email: '{first_line}'")

            # FLAW 3.2: "YES SYRMA 10 CONFIRM" — override concentration warning
            confirm_match = re.search(r'YES\s+([A-Z&]+)\s+(\d+)\s+CONFIRM', first_line)
            yes_qty_match = re.search(r'YES\s+([A-Z&]+)\s+(\d+)', first_line)
            yes_match     = re.search(r'YES\s+([A-Z&]+)', first_line)
            no_match      = re.search(r'NO\s+([A-Z&]+)', first_line)

            if confirm_match:
                ticker = confirm_match.group(1).strip()
                qty    = int(confirm_match.group(2).strip())
                confirmed_overrides.append((ticker, qty))
                print(f"[approval] CONFIRMED OVERRIDE: {ticker} x {qty}")
                mail.store(msg_id, '+FLAGS', '\\Seen')
            elif yes_qty_match:
                ticker = yes_qty_match.group(1).strip()
                qty    = int(yes_qty_match.group(2).strip())
                approvals_with_qty.append((ticker, qty))
                print(f"[approval] APPROVED with QTY: {ticker} x {qty}")
                mail.store(msg_id, '+FLAGS', '\\Seen')
            elif yes_match:
                ticker = yes_match.group(1).strip()
                approvals_need_qty.append(ticker)
                print(f"[approval] APPROVED (asking qty): {ticker}")
                mail.store(msg_id, '+FLAGS', '\\Seen')
            elif no_match:
                ticker = no_match.group(1).strip()
                rejections.append(ticker)
                print(f"[approval] REJECTED: {ticker}")
                mail.store(msg_id, '+FLAGS', '\\Seen')

        mail.logout()

    except Exception as e:
        print(f"[approval] Error checking Gmail: {e}")

    return approvals_with_qty, approvals_need_qty, rejections, confirmed_overrides


def get_stock_details(ticker):
    try:
        ticker_yf = ticker + NSE_SUFFIX
        stock     = yf.Ticker(ticker_yf)
        info      = stock.info
        price     = info.get('currentPrice') or info.get('regularMarketPrice')
        if not price:
            return None

        market_cap = info.get('marketCap')
        cap_category = 'Mid Cap'
        if market_cap:
            if market_cap >= 20_000_00_00_000:
                cap_category = 'Large Cap'
            elif market_cap >= 5_000_00_00_000:
                cap_category = 'Mid Cap'
            else:
                cap_category = 'Small Cap'

        return {
            'ticker':        ticker,
            'ticker_yf':     ticker_yf,
            'name':          info.get('longName', ticker),
            'sector':        info.get('sector', 'N/A'),
            'cap_category':  cap_category,
            'current_price': round(float(price), 2),
        }
    except Exception as e:
        print(f"[approval] Error fetching details for {ticker}: {e}")
        return None


def get_portfolio_summary():
    """Get current portfolio total value and sector breakdown for concentration checks"""
    try:
        from config import USE_GOOGLE_SHEETS
        if USE_GOOGLE_SHEETS:
            from sheets_handler import read_holdings
            holdings = read_holdings()
        else:
            return 0, {}

        total_value     = 0
        sector_values   = {}

        for h in holdings:
            try:
                info  = yf.Ticker(h['ticker'] + NSE_SUFFIX).info
                price = info.get('currentPrice') or info.get('regularMarketPrice') or h.get('buying_price', 0)
                value = round(float(price) * h['qty'], 2)
                total_value += value

                sector = info.get('sector', h.get('industry', 'Unknown'))
                sector_values[sector] = sector_values.get(sector, 0) + value
            except:
                value = round(h.get('buying_price', 0) * h['qty'], 2)
                total_value += value

        return total_value, sector_values

    except Exception as e:
        print(f"[approval] Error fetching portfolio summary: {e}")
        return 0, {}


def check_concentration(ticker, stock_info, qty):
    """
    FLAW 3.2: Position sizing check — warn if single stock > 15% of portfolio.
    FLAW 3.3: Sector check — warn if sector would exceed 30%.
    Returns (ok_to_proceed, warning_type, warning_message)
    """
    total_value, sector_values = get_portfolio_summary()

    if total_value == 0:
        return True, None, None

    new_position_value = round(stock_info['current_price'] * qty, 2)
    new_total          = total_value + new_position_value
    position_pct       = round((new_position_value / new_total) * 100, 1)

    # FLAW 3.2: Position sizing
    if position_pct > 15:
        msg = (f"⚠️ CONCENTRATION WARNING: Adding {ticker} ({qty} shares @ Rs.{stock_info['current_price']:,.2f}) "
               f"would make it {position_pct}% of your portfolio (max recommended: 15%).\n\n"
               f"Reply 'YES {ticker} {qty} CONFIRM' to override this warning and proceed anyway.")
        return False, 'position', msg

    # FLAW 3.3: Sector concentration
    new_sector        = stock_info.get('sector', 'Unknown')
    current_sector_v  = sector_values.get(new_sector, 0)
    new_sector_total  = current_sector_v + new_position_value
    sector_pct        = round((new_sector_total / new_total) * 100, 1)

    if sector_pct > 30:
        msg = (f"⚠️ SECTOR OVERWEIGHT WARNING: Adding {ticker} would make '{new_sector}' "
               f"{sector_pct}% of your portfolio (max recommended: 30%).\n\n"
               f"Reply 'YES {ticker} {qty} CONFIRM' to override and proceed anyway.")
        return False, 'sector', msg

    return True, None, None


def send_concentration_warning_email(ticker, stock_info, qty, warning_msg):
    subject = f"⚠️ Concentration Warning — {ticker} | Action Required"
    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:#E65100;color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">⚠️ Portfolio Concentration Warning</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{ticker} — {stock_info['name']}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;border-left:4px solid #E65100;">
    <pre style="white-space:pre-wrap;font-family:Arial;font-size:14px;">{warning_msg}</pre>
  </div>
  <div style="background:#FFF3E0;border:2px dashed #E65100;border-radius:8px;padding:16px;text-align:center;">
    To add anyway, reply:<br><br>
    <code style="font-size:18px;color:#E65100;">YES {ticker} {qty} CONFIRM</code>
  </div>
</body></html>"""
    send_report_email(subject, html)


def send_qty_request_email(ticker, stock_info):
    price = stock_info['current_price']
    subject = f"How many shares of {ticker}? | Rs.{price}/share"
    rows = ''
    for q in [10, 25, 50, 100, 200, 500]:
        rows += f"<tr><td style='padding:8px 14px;border-bottom:1px solid #eee;'>{q} shares</td><td style='padding:8px 14px;border-bottom:1px solid #eee;font-weight:bold;'>Rs.{price*q:,.2f}</td></tr>"

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:#2E7D32;color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">How many shares of {ticker}?</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{stock_info['name']} | {stock_info.get('cap_category','')} | {stock_info.get('sector','')}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;">
    <h3 style="margin:0 0 12px;">Current Price: Rs.{price:,.2f} per share</h3>
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#f8f9fa;"><th style="padding:8px 14px;text-align:left;">Quantity</th><th style="padding:8px 14px;text-align:left;">Total Investment</th></tr>
      {rows}
    </table>
  </div>
  <div style="background:#E3F2FD;border:2px dashed #1976D2;border-radius:10px;padding:20px;text-align:center;">
    Reply: <code style="font-size:20px;color:#1976D2;">YES {ticker} [qty]</code>
  </div>
</body></html>"""
    send_report_email(subject, html)
    print(f"[approval] Qty request sent for {ticker}")


def send_confirmation_email(ticker, stock_info, qty, next_row=None, target_price=None):
    buying_price = stock_info['current_price']
    investment   = round(buying_price * qty, 2)
    today        = datetime.now().strftime('%d %B %Y')

    subject = f"{ticker} Added | {qty} shares @ Rs.{buying_price:,.2f} | Total Rs.{investment:,.2f}"
    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:#2E7D32;color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">✅ Stock Added to Portfolio</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{stock_info['name']} | {stock_info.get('cap_category','')} | {stock_info.get('sector','')}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;">
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;width:45%;">Ticker</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;">{ticker}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Buying Price</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;">Rs.{buying_price:,.2f}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Buying Date</td><td style="padding:10px 14px;">{today}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Quantity</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;">{qty} shares</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Investment Amount</td><td style="padding:10px 14px;font-size:20px;font-weight:bold;color:#1B5E20;">Rs.{investment:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Category</td><td style="padding:10px 14px;">{stock_info.get('cap_category','N/A')} | {stock_info.get('sector','N/A')}</td></tr>
      {'<tr style="background:#E8F5E9;"><td style="padding:10px 14px;font-weight:bold;color:#2E7D32;">🎯 AI Target Price</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;color:#2E7D32;">Rs.' + f"{target_price:,.2f}" + '</td></tr>' if target_price else ''}
    </table>
  </div>
</body></html>"""
    send_report_email(subject, html)
    print(f"[approval] Confirmation email sent for {ticker}")


def _add_to_holdings(ticker, stock_info, qty):
    """Add stock to Google Sheets holdings"""
    from config import USE_GOOGLE_SHEETS
    buying_date  = datetime.now().strftime('%Y-%m-%d')
    target_price = stock_info.get('target_price', None)

    if USE_GOOGLE_SHEETS:
        from sheets_handler import add_stock_to_holdings
        sno = add_stock_to_holdings(
            ticker=ticker,
            stock_name=stock_info.get('name', ticker),
            industry=stock_info.get('sector', 'N/A'),
            buying_price=stock_info['current_price'],
            buying_date=buying_date,
            qty=qty,
            target_price=target_price,
            cap_category=stock_info.get('cap_category', 'Mid Cap'),
            sector=stock_info.get('sector', 'N/A'),
        )
        return sno
    else:
        success, sno = add_stock_via_vba(
            ticker=ticker,
            buying_price=stock_info['current_price'],
            buying_date=buying_date,
            qty=qty
        )
        return sno if success else None


def process_approvals():
    print("[approval] Checking Gmail for YES/NO replies...")
    approvals_with_qty, approvals_need_qty, rejections, confirmed_overrides = fetch_approval_replies()

    if not approvals_with_qty and not approvals_need_qty and not rejections and not confirmed_overrides:
        print("[approval] No new replies found.")
        return

    # Process confirmed overrides (bypass concentration checks)
    for ticker, qty in confirmed_overrides:
        stock_info = get_stock_details(ticker)
        if not stock_info:
            print(f"[approval] Could not fetch details for {ticker}. Skipping.")
            continue
        sno = _add_to_holdings(ticker, stock_info, qty)
        send_confirmation_email(ticker, stock_info, qty, sno)
        print(f"[approval] ✅ {ticker} added (override confirmed).")

    # Process normal approvals with concentration checks
    for ticker, qty in approvals_with_qty:
        stock_info = get_stock_details(ticker)
        if not stock_info:
            print(f"[approval] Could not fetch details for {ticker}. Skipping.")
            continue

        # FLAW 3.2 + 3.3: Concentration checks
        ok, warn_type, warn_msg = check_concentration(ticker, stock_info, qty)
        if not ok:
            send_concentration_warning_email(ticker, stock_info, qty, warn_msg)
            print(f"[approval] ⚠️  {ticker} — concentration warning sent ({warn_type}). Waiting for CONFIRM.")
            continue

        sno = _add_to_holdings(ticker, stock_info, qty)
        send_confirmation_email(ticker, stock_info, qty, sno)

    # Process approvals needing qty
    for ticker in approvals_need_qty:
        stock_info = get_stock_details(ticker)
        if not stock_info:
            print(f"[approval] Could not fetch details for {ticker}. Skipping.")
            continue
        send_qty_request_email(ticker, stock_info)

    for ticker in rejections:
        print(f"[approval] {ticker} skipped.")


if __name__ == "__main__":
    process_approvals()