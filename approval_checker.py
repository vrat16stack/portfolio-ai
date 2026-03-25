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
    HOLDINGS_SHEET = "Sheet1"

try:
    from add_stock_macro import add_stock_via_vba
except ImportError:
    def add_stock_via_vba(ticker, buying_price, buying_date, qty):
        return True, 1


def fetch_approval_replies():
    approvals_with_qty = []
    approvals_need_qty = []
    rejections = []

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

            yes_qty_match = re.search(r'YES\s+([A-Z&]+)\s+(\d+)', first_line)
            yes_match = re.search(r'YES\s+([A-Z&]+)', first_line)
            no_match = re.search(r'NO\s+([A-Z&]+)', first_line)

            if yes_qty_match:
                ticker = yes_qty_match.group(1).strip()
                qty = int(yes_qty_match.group(2).strip())
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

    return approvals_with_qty, approvals_need_qty, rejections


def get_stock_details(ticker):
    try:
        ticker_yf = ticker + NSE_SUFFIX
        stock = yf.Ticker(ticker_yf)
        info = stock.info
        current_price = info.get('currentPrice') or info.get('regularMarketPrice')
        if not current_price:
            return None
        return {
            'ticker': ticker,
            'ticker_yf': ticker_yf,
            'name': info.get('longName', ticker),
            'sector': info.get('sector', 'N/A'),
            'current_price': round(float(current_price), 2),
        }
    except Exception as e:
        print(f"[approval] Error fetching details for {ticker}: {e}")
        return None


def send_qty_request_email(ticker, stock_info):
    price = stock_info['current_price']
    subject = f"How many shares of {ticker} do you want? | Rs.{price}/share"

    rows = ''
    for q in [10, 25, 50, 100, 200, 500]:
        rows += f"<tr><td style='padding:8px 14px;border-bottom:1px solid #eee;'>{q} shares</td><td style='padding:8px 14px;border-bottom:1px solid #eee;font-weight:bold;'>Rs.{price*q:,.2f}</td></tr>"

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:#2E7D32;color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">How many shares of {ticker}?</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{stock_info['name']}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;">
    <h3 style="margin:0 0 12px;">Current Price: Rs.{price:,.2f} per share</h3>
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#f8f9fa;">
        <th style="padding:8px 14px;text-align:left;">Quantity</th>
        <th style="padding:8px 14px;text-align:left;">Total Investment</th>
      </tr>
      {rows}
    </table>
  </div>
  <div style="background:#E3F2FD;border:2px dashed #1976D2;border-radius:10px;padding:20px;text-align:center;">
    <strong>Reply to this email with:</strong><br><br>
    <code style="font-size:20px;color:#1976D2;">YES {ticker} [qty]</code><br><br>
    Example: YES {ticker} 50 to buy 50 shares for Rs.{price*50:,.2f}
  </div>
</body></html>"""

    send_report_email(subject, html)
    print(f"[approval] Qty request sent for {ticker}")


def send_confirmation_email(ticker, stock_info, qty, next_row=None, target_price=None):
    buying_price = stock_info['current_price']
    investment = round(buying_price * qty, 2)
    today = datetime.now().strftime('%d %B %Y')
    if next_row is None:
        next_row = "the new"

    subject = f"{ticker} Added | {qty} shares @ Rs.{buying_price:,.2f} | Total Rs.{investment:,.2f}"

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:#2E7D32;color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">Stock Added to Your Portfolio</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{stock_info['name']}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;">
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;width:45%;">Buying Price</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;">Rs.{buying_price:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Buying Date</td><td style="padding:10px 14px;">{today}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Quantity</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;">{qty} shares</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Investment Amount</td><td style="padding:10px 14px;font-size:20px;font-weight:bold;color:#1B5E20;">Rs.{investment:,.2f}</td></tr>
      {'<tr style="background:#E8F5E9;"><td style="padding:10px 14px;font-weight:bold;color:#2E7D32;">🎯 AI Target Price</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;color:#2E7D32;">Rs.' + f"{target_price:,.2f}" + '</td></tr>' if target_price else ''}
    </table>
  </div>
  <div style="background:#E8F5E9;padding:14px;border-radius:6px;text-align:center;">
    Stock added to your Google Sheets portfolio. {'Auto-sell will trigger when price hits Rs.' + f'{target_price:,.2f}' + '!' if target_price else 'Live analysis will include it from next run!'}
  </div>
</body></html>"""

    send_report_email(subject, html)
    print(f"[approval] Confirmation email sent for {ticker}")


def process_approvals():
    print("[approval] Checking Gmail for YES/NO replies...")
    approvals_with_qty, approvals_need_qty, rejections = fetch_approval_replies()

    if not approvals_with_qty and not approvals_need_qty and not rejections:
        print("[approval] No new replies found.")
        return

    for ticker, qty in approvals_with_qty:
        stock_info = get_stock_details(ticker)
        if not stock_info:
            print(f"[approval] Could not fetch details for {ticker}. Skipping.")
            continue
        buying_date = datetime.now().strftime('%Y-%m-%d')

        from config import USE_GOOGLE_SHEETS
        if USE_GOOGLE_SHEETS:
            from sheets_handler import add_stock_to_holdings
            # Get target price from scout data if stored
            target_price = stock_info.get('target_price', None)
            sno = add_stock_to_holdings(
                ticker=ticker,
                stock_name=stock_info.get('name', ticker),
                industry=stock_info.get('sector', 'N/A'),
                buying_price=stock_info['current_price'],
                buying_date=buying_date,
                qty=qty,
                target_price=target_price
            )
            send_confirmation_email(ticker, stock_info, qty, sno, target_price=target_price)
        else:
            success, sno = add_stock_via_vba(
                ticker=ticker,
                buying_price=stock_info['current_price'],
                buying_date=buying_date,
                qty=qty
            )
            if success:
                send_confirmation_email(ticker, stock_info, qty, sno)

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
