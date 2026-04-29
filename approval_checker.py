"""
approval_checker.py
Processes YES/NO email replies for stock additions.

UPDATED FLOW (Pre-Market Buy Logic):
  Previously: YES reply → immediately add to Holdings at current price.
  Now:        YES reply → store as PENDING BUY in PendingBuys sheet.
              Actual buy decision happens at 9:20 AM via pending_buys_handler.py
              using real opening price + gap logic.

  Q2 FIX: Smart overnight news analysis using Groq BEFORE storing PENDING BUY.
           If Groq confirms genuinely harmful news → cancel buy, send alert.
           Groq distinguishes real risk (fraud, probe, insolvency) from noise
           (analyst downgrade, sector headwind, general market fall).

  Existing behaviour preserved for:
  - FLAW 3.2: Position sizing warning > 15%
  - FLAW 3.3: Sector concentration warning > 30%
  - "YES TICKER QTY CONFIRM" override still works
  - Qty request email when YES sent without quantity
"""

import imaplib
import email
import re
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
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

from pending_buys_handler import store_pending_buy


# ── Q2: Smart news analysis using Groq ───────────────────────────────────────

def _fetch_news_headlines(ticker, stock_name):
    """Fetch latest Google News RSS headlines for a stock."""
    headlines = []
    try:
        query = urllib.parse.quote(f"{stock_name} NSE India stock")
        url   = f"https://news.google.com/rss/search?q={query}&hl=en-IN&gl=IN&ceid=IN:en"
        req   = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            content = resp.read()
        root  = ET.fromstring(content)
        items = root.findall('.//item')
        for item in items[:8]:
            title = item.find('title')
            desc  = item.find('description')
            if title is not None and title.text:
                text = title.text
                if desc is not None and desc.text:
                    import re as _re
                    clean = _re.sub(r'<[^>]+>', ' ', desc.text).strip()
                    if clean:
                        text = f"{title.text} — {clean[:200]}"
                headlines.append(text)
    except Exception as e:
        print(f"[approval] News fetch error for {ticker}: {e}")
    return headlines


def _analyse_news_with_groq(ticker, stock_name, headlines):
    """
    Q2 FIX: Ask Groq to judge whether overnight news poses a REAL risk
    to executing this buy order tomorrow morning.

    Groq is instructed to be CONSERVATIVE — only cancel for confirmed,
    severe, company-specific events. When in doubt → PROCEED.

    Returns: (should_cancel: bool, risk_level: str, reasoning: str, key_news: str)
    """
    from groq import Groq
    from config import GROQ_API_KEY, GROQ_MODEL

    if not headlines:
        print(f"[approval] {ticker}: No news found — proceeding with buy.")
        return False, 'NONE', 'No overnight news found. Proceeding with buy.', ''

    news_text = '\n'.join([f"- {h}" for h in headlines])

    prompt = f"""You are an expert NSE India stock market analyst making a critical pre-market buy decision.

Stock: {stock_name} (NSE: {ticker})
A buy order is scheduled for tomorrow morning at market open.

OVERNIGHT NEWS HEADLINES:
{news_text}

TASK: Determine if this news poses a REAL and SEVERE risk that should CANCEL this buy order.

CANCEL the buy ONLY if the news contains CONFIRMED information about:
1. Fraud, financial scam, or accounting irregularities at THIS company
2. SEBI, ED, CBI, Income Tax, or any regulatory raid/investigation on THIS company
3. Promoter arrested, charged, or selling large stake due to distress
4. Debt default, insolvency petition, or bankruptcy filing
5. Major earnings miss (>20% below estimates) with confirmed quarterly numbers
6. Court order stopping operations, government cancelling licence
7. Product/service recall causing major business disruption

DO NOT CANCEL for:
- General market/Nifty falling (affects all stocks equally)
- Sector-level headwinds not specific to this company
- Analyst downgrades or price target cuts
- Unconfirmed rumours or speculative reports
- Minor earnings miss (less than 20%)
- Old news (clearly more than 5 days old)
- Competitor news that may indirectly affect this company
- Global macro concerns (US Fed, crude oil, dollar, etc.)

Be CONSERVATIVE — only cancel if you are CONFIDENT the news is severe and company-specific.
When in doubt, PROCEED. A false cancellation is worse than a bad entry.

Respond in EXACTLY this format, no extra lines:
DECISION: [CANCEL or PROCEED]
RISK_LEVEL: [CRITICAL / HIGH / MEDIUM / LOW / NONE]
KEY_NEWS: [Most relevant headline in one line, or NONE]
REASONING: [2-3 sentences. If CANCEL: what specific confirmed event. If PROCEED: why news does not warrant cancellation.]
CONFIDENCE: [HIGH / MEDIUM / LOW]"""

    try:
        client   = Groq(api_key=GROQ_API_KEY)
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=250,
            temperature=0.1
        )
        text = response.choices[0].message.content.strip()

        decision   = 'PROCEED'
        risk_level = 'NONE'
        key_news   = ''
        reasoning  = ''
        confidence = 'LOW'

        for line in text.split('\n'):
            line = line.strip()
            if line.startswith('DECISION:'):
                decision   = line.split(':', 1)[1].strip().upper()
            elif line.startswith('RISK_LEVEL:'):
                risk_level = line.split(':', 1)[1].strip().upper()
            elif line.startswith('KEY_NEWS:'):
                key_news   = line.split(':', 1)[1].strip()
            elif line.startswith('REASONING:'):
                reasoning  = line.split(':', 1)[1].strip()
            elif line.startswith('CONFIDENCE:'):
                confidence = line.split(':', 1)[1].strip().upper()

        # Only cancel if Groq says CANCEL with at least MEDIUM confidence
        should_cancel = (decision == 'CANCEL' and confidence in ('HIGH', 'MEDIUM'))

        print(f"[approval] Q2 News analysis for {ticker}: "
              f"{decision} | Risk: {risk_level} | Confidence: {confidence}")
        print(f"[approval] Reasoning: {reasoning}")

        return should_cancel, risk_level, reasoning, key_news

    except Exception as e:
        print(f"[approval] Groq news analysis error for {ticker}: {e}")
        # On any Groq failure → PROCEED (don't block buy on technical error)
        return False, 'UNKNOWN', f'News analysis unavailable ({e}) — proceeding.', ''


def _send_news_cancel_email(ticker, stock_name, qty, scout_price,
                             risk_level, reasoning, key_news, headlines):
    """Email sent when buy is cancelled due to confirmed negative news."""
    subject = (f"🚫 Buy CANCELLED: {ticker} | Negative News Detected "
               f"| Risk Level: {risk_level}")

    news_items_html = ''.join(
        f"<li style='margin-bottom:6px;font-size:13px;'>{h}</li>"
        for h in headlines[:5]
    )
    key_news_html = (
        f'<div style="background:#FFF3E0;border-radius:6px;padding:10px 12px;'
        f'margin-bottom:12px;font-size:13px;">'
        f'<strong>Key News:</strong> {key_news}</div>'
        if key_news and key_news.upper() != 'NONE' else ''
    )

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:660px;margin:auto;
             background:#f0f2f5;padding:20px;">

  <div style="background:linear-gradient(135deg,#4A148C,#7B1FA2);color:white;
              padding:28px;border-radius:14px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">🚫 Buy Order Cancelled</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{stock_name} ({ticker})</p>
    <p style="margin:4px 0 0;opacity:0.75;font-size:13px;">
      Pre-market news analysis flagged a critical risk</p>
  </div>

  <div style="background:white;border-radius:12px;padding:20px;margin-bottom:16px;
              border-left:5px solid #C62828;">
    <h3 style="margin:0 0 10px;color:#C62828;">⚠️ Why This Buy Was Cancelled</h3>
    <div style="background:#FFEBEE;border-radius:8px;padding:14px;margin-bottom:14px;">
      <div style="font-size:12px;color:#C62828;font-weight:bold;margin-bottom:6px;">
        Risk Level: {risk_level}
      </div>
      <div style="font-size:13px;line-height:1.7;color:#333;">{reasoning}</div>
    </div>
    {key_news_html}
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:14px;">
      <tr><td style="padding:9px 12px;font-weight:bold;width:40%;">Stock</td>
          <td style="padding:9px 12px;">{ticker} — {stock_name}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:9px 12px;font-weight:bold;">Qty Requested</td>
          <td style="padding:9px 12px;">{qty} shares</td></tr>
      <tr><td style="padding:9px 12px;font-weight:bold;">Scout Price</td>
          <td style="padding:9px 12px;">₹{scout_price:,.2f}</td></tr>
    </table>
    <div style="font-size:13px;font-weight:bold;margin-bottom:8px;">
      Latest News Headlines Analysed:</div>
    <ul style="margin:0;padding-left:18px;color:#444;line-height:1.8;">
      {news_items_html}
    </ul>
  </div>

  <div style="background:#F3E5F5;border-radius:10px;padding:14px;
              text-align:center;font-size:13px;color:#7B1FA2;">
    ⚠️ Stock NOT added to Holdings or PendingBuys.<br>
    Monitor the situation. Watch future scout emails for a fresh entry opportunity.
  </div>

</body></html>"""

    send_report_email(subject, html)
    print(f"[approval] 🚫 News cancel email sent for {ticker}")


# ── Gmail reading (unchanged) ──────────────────────────────────────────────────

def fetch_approval_replies():
    approvals_with_qty   = []
    approvals_need_qty   = []
    rejections           = []
    confirmed_overrides  = []

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

        market_cap   = info.get('marketCap')
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
    try:
        from config import USE_GOOGLE_SHEETS
        if USE_GOOGLE_SHEETS:
            from sheets_handler import read_holdings
            holdings = read_holdings()
        else:
            return 0, {}

        total_value   = 0
        sector_values = {}

        for h in holdings:
            try:
                info  = yf.Ticker(h['ticker'] + NSE_SUFFIX).info
                price = (info.get('currentPrice') or info.get('regularMarketPrice')
                         or h.get('buying_price', 0))
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
    total_value, sector_values = get_portfolio_summary()

    if total_value == 0:
        return True, None, None

    new_position_value = round(stock_info['current_price'] * qty, 2)
    new_total          = total_value + new_position_value
    position_pct       = round((new_position_value / new_total) * 100, 1)

    if position_pct > 15:
        msg = (f"⚠️ CONCENTRATION WARNING: Adding {ticker} ({qty} shares @ "
               f"Rs.{stock_info['current_price']:,.2f}) would make it {position_pct}% "
               f"of your portfolio (max recommended: 15%).\n\n"
               f"Reply 'YES {ticker} {qty} CONFIRM' to override and proceed anyway.")
        return False, 'position', msg

    new_sector       = stock_info.get('sector', 'Unknown')
    current_sector_v = sector_values.get(new_sector, 0)
    new_sector_total = current_sector_v + new_position_value
    sector_pct       = round((new_sector_total / new_total) * 100, 1)

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
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;
             background:#f5f5f5;padding:20px;">
  <div style="background:#E65100;color:white;padding:24px;border-radius:12px;
              margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">⚠️ Portfolio Concentration Warning</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{ticker} — {stock_info['name']}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;
              border-left:4px solid #E65100;">
    <pre style="white-space:pre-wrap;font-family:Arial;font-size:14px;">{warning_msg}</pre>
  </div>
  <div style="background:#FFF3E0;border:2px dashed #E65100;border-radius:8px;
              padding:16px;text-align:center;">
    To add anyway, reply:<br><br>
    <code style="font-size:18px;color:#E65100;">YES {ticker} {qty} CONFIRM</code>
  </div>
</body></html>"""
    send_report_email(subject, html)


def send_qty_request_email(ticker, stock_info):
    price   = stock_info['current_price']
    subject = f"How many shares of {ticker}? | Rs.{price}/share"
    rows    = ''
    for q in [10, 25, 50, 100, 200, 500]:
        rows += (f"<tr><td style='padding:8px 14px;border-bottom:1px solid #eee;'>{q} shares</td>"
                 f"<td style='padding:8px 14px;border-bottom:1px solid #eee;font-weight:bold;'>"
                 f"Rs.{price*q:,.2f}</td></tr>")

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;
             background:#f5f5f5;padding:20px;">
  <div style="background:#2E7D32;color:white;padding:24px;border-radius:12px;
              margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">How many shares of {ticker}?</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{stock_info['name']} |
       {stock_info.get('cap_category','')} | {stock_info.get('sector','')}</p>
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
  <div style="background:#E3F2FD;border:2px dashed #1976D2;border-radius:10px;
              padding:20px;text-align:center;">
    Reply: <code style="font-size:20px;color:#1976D2;">YES {ticker} [qty]</code>
  </div>
</body></html>"""
    send_report_email(subject, html)
    print(f"[approval] Qty request sent for {ticker}")


def _send_pending_buy_queued_email(ticker, stock_info, qty, scout_price, original_target):
    upside_pct     = round(((original_target - scout_price) / scout_price) * 100, 2) \
        if scout_price else 0
    investment_est = round(scout_price * qty, 2)
    today_str      = datetime.now().strftime('%d %B %Y')
    subject        = f"⏳ Buy Queued: {ticker} | {qty} shares | Executing at 9:20 AM open price"

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:640px;margin:auto;
             background:#f0f2f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#1565C0,#1976D2);color:white;
              padding:28px;border-radius:14px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">⏳ Buy Order Queued</h1>
    <p style="margin:6px 0 0;opacity:0.9;font-size:15px;">
      {stock_info['name']} ({ticker})</p>
    <p style="margin:4px 0 0;opacity:0.75;font-size:13px;">{today_str}</p>
  </div>
  <div style="background:white;border-radius:12px;padding:20px;margin-bottom:16px;
              box-shadow:0 2px 8px rgba(0,0,0,0.07);">
    <h3 style="margin:0 0 14px;color:#1565C0;">📋 Order Details</h3>
    <table style="width:100%;border-collapse:collapse;">
      <tr><td style="padding:10px 14px;font-weight:bold;width:45%;">Ticker</td>
          <td style="padding:10px 14px;font-size:18px;font-weight:bold;">{ticker}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Quantity</td>
          <td style="padding:10px 14px;font-size:18px;font-weight:bold;">{qty} shares</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Scout Price (Reference)</td>
          <td style="padding:10px 14px;">₹{scout_price:,.2f}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Original Target</td>
          <td style="padding:10px 14px;">₹{original_target:,.2f}
            <span style="color:#2E7D32;font-size:13px;">(+{upside_pct:.1f}%)</span></td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Estimated Investment</td>
          <td style="padding:10px 14px;font-size:16px;color:#1565C0;">
            ~₹{investment_est:,.2f}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Sector</td>
          <td style="padding:10px 14px;">{stock_info.get('sector','N/A')}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Category</td>
          <td style="padding:10px 14px;">{stock_info.get('cap_category','N/A')}</td></tr>
    </table>
  </div>
  <div style="background:#E8F5E9;border-radius:8px;padding:10px 14px;margin-bottom:12px;
              font-size:13px;color:#2E7D32;">
    ✅ Overnight news check passed — no critical risks detected.
  </div>
  <div style="background:#E3F2FD;border-radius:12px;padding:16px;margin-bottom:16px;">
    <h3 style="margin:0 0 10px;color:#1565C0;">⚙️ What Happens Next</h3>
    <ol style="margin:0;padding-left:18px;line-height:1.9;font-size:14px;">
      <li>At <strong>9:20 AM</strong>, system fetches {ticker}'s actual opening price.</li>
      <li>Gap from yesterday's close is checked against {stock_info.get('sector','sector')} threshold.</li>
      <li>If gap is acceptable → <strong>BUY EXECUTED</strong></li>
      <li>If gap is extreme → <strong>BUY CANCELLED</strong> with alert</li>
    </ol>
  </div>
</body></html>"""

    send_report_email(subject, html)
    print(f"[approval] ⏳ Buy queued email sent for {ticker}")


def _get_scout_price_and_target(ticker):
    try:
        from sheets_handler import get_recommendations_log
        logs = get_recommendations_log()
        ticker_logs = [r for r in logs
                       if str(r.get('Ticker', '')).upper() == ticker.upper()]
        if ticker_logs:
            ticker_logs.sort(key=lambda x: str(x.get('Date', '')), reverse=True)
            latest   = ticker_logs[0]
            scout_p  = float(str(latest.get('Recommended Price', 0) or 0).replace(',', ''))
            target_p = float(str(latest.get('Target Price', 0) or 0).replace(',', ''))
            if scout_p > 0:
                print(f"[approval] Found scout price for {ticker}: ₹{scout_p}, target ₹{target_p}")
                return scout_p, target_p if target_p > 0 else None
    except Exception as e:
        print(f"[approval] Could not fetch scout price for {ticker}: {e}")
    return None, None


def _add_to_holdings_direct(ticker, stock_info, qty):
    from config import USE_GOOGLE_SHEETS
    buying_date  = datetime.now().strftime('%Y-%m-%d')
    target_price = stock_info.get('target_price', None)
    if USE_GOOGLE_SHEETS:
        from sheets_handler import add_stock_to_holdings
        sno = add_stock_to_holdings(
            ticker=ticker, stock_name=stock_info.get('name', ticker),
            industry=stock_info.get('sector', 'N/A'),
            buying_price=stock_info['current_price'],
            buying_date=buying_date, qty=qty,
            target_price=target_price,
            cap_category=stock_info.get('cap_category', 'Mid Cap'),
            sector=stock_info.get('sector', 'N/A'),
        )
        return sno
    else:
        success, sno = add_stock_via_vba(
            ticker=ticker, buying_price=stock_info['current_price'],
            buying_date=buying_date, qty=qty)
        return sno if success else None


def send_confirmation_email(ticker, stock_info, qty, next_row=None, target_price=None):
    buying_price = stock_info['current_price']
    investment   = round(buying_price * qty, 2)
    today        = datetime.now().strftime('%d %B %Y')
    subject      = (f"{ticker} Added (Override) | {qty} shares @ "
                    f"Rs.{buying_price:,.2f} | Total Rs.{investment:,.2f}")
    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;
             background:#f5f5f5;padding:20px;">
  <div style="background:#2E7D32;color:white;padding:24px;border-radius:12px;
              margin-bottom:20px;">
    <h1 style="margin:0;font-size:20px;">✅ Stock Added to Portfolio (Concentration Override)</h1>
    <p style="margin:6px 0 0;opacity:0.9;">{stock_info['name']} |
       {stock_info.get('cap_category','')} | {stock_info.get('sector','')}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;">
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#f8f9fa;">
        <td style="padding:10px 14px;font-weight:bold;width:45%;">Ticker</td>
        <td style="padding:10px 14px;font-size:18px;font-weight:bold;">{ticker}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Buying Price</td>
          <td style="padding:10px 14px;font-size:18px;font-weight:bold;">Rs.{buying_price:,.2f}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Buying Date</td>
          <td style="padding:10px 14px;">{today}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Quantity</td>
          <td style="padding:10px 14px;font-size:18px;font-weight:bold;">{qty} shares</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Investment Amount</td>
          <td style="padding:10px 14px;font-size:20px;font-weight:bold;color:#1B5E20;">
            Rs.{investment:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Category</td>
          <td style="padding:10px 14px;">
            {stock_info.get('cap_category','N/A')} | {stock_info.get('sector','N/A')}</td></tr>
      {'<tr style="background:#E8F5E9;"><td style="padding:10px 14px;font-weight:bold;color:#2E7D32;">🎯 Target Price</td><td style="padding:10px 14px;font-size:18px;font-weight:bold;color:#2E7D32;">Rs.' + f"{target_price:,.2f}" + '</td></tr>' if target_price else ''}
    </table>
  </div>
</body></html>"""
    send_report_email(subject, html)
    print(f"[approval] Confirmation email sent for {ticker} (override)")


def process_approvals():
    print("[approval] === Checking Gmail for YES/NO replies ===")
    approvals_with_qty, approvals_need_qty, rejections, confirmed_overrides = fetch_approval_replies()

    if not approvals_with_qty and not approvals_need_qty and not rejections and not confirmed_overrides:
        print("[approval] No new replies found.")
        return

    # ── CONFIRM overrides: bypass everything, add directly ─────────────────
    for ticker, qty in confirmed_overrides:
        stock_info = get_stock_details(ticker)
        if not stock_info:
            print(f"[approval] Could not fetch details for {ticker}. Skipping.")
            continue
        sno = _add_to_holdings_direct(ticker, stock_info, qty)
        send_confirmation_email(ticker, stock_info, qty, sno)
        print(f"[approval] ✅ {ticker} added directly (concentration override confirmed).")

    # ── Normal approvals ────────────────────────────────────────────────────
    for ticker, qty in approvals_with_qty:
        stock_info = get_stock_details(ticker)
        if not stock_info:
            print(f"[approval] Could not fetch details for {ticker}. Skipping.")
            continue

        # ── Q2: Smart news analysis BEFORE anything else ───────────────────
        print(f"[approval] Q2: Running overnight news analysis for {ticker}...")
        headlines = _fetch_news_headlines(ticker, stock_info['name'])
        should_cancel, risk_level, reasoning, key_news = _analyse_news_with_groq(
            ticker, stock_info['name'], headlines
        )

        if should_cancel:
            print(f"[approval] 🚫 {ticker} CANCELLED — news risk: {risk_level}")
            scout_price, _ = _get_scout_price_and_target(ticker)
            _send_news_cancel_email(
                ticker, stock_info['name'], qty,
                scout_price or stock_info['current_price'],
                risk_level, reasoning, key_news, headlines
            )
            continue

        print(f"[approval] ✅ {ticker} news check passed — proceeding")

        # Concentration checks
        ok, warn_type, warn_msg = check_concentration(ticker, stock_info, qty)
        if not ok:
            send_concentration_warning_email(ticker, stock_info, qty, warn_msg)
            print(f"[approval] ⚠️  {ticker} — concentration warning sent ({warn_type}).")
            continue

        # Get scout price and target
        scout_price, original_target = _get_scout_price_and_target(ticker)
        if not scout_price:
            scout_price     = stock_info['current_price']
            original_target = round(scout_price * 1.10, 2)
            print(f"[approval] {ticker}: No scout record — using live price ₹{scout_price}")

        if not original_target or original_target <= scout_price:
            original_target = round(scout_price * 1.10, 2)

        # Store as PENDING BUY
        stored = store_pending_buy(
            ticker=ticker, stock_name=stock_info['name'], qty=qty,
            scout_price=scout_price, original_target=original_target,
            sector=stock_info.get('sector', 'N/A'),
            cap_category=stock_info.get('cap_category', 'Mid Cap'),
        )

        if stored:
            _send_pending_buy_queued_email(
                ticker, stock_info, qty, scout_price, original_target
            )

    # ── Qty requests ────────────────────────────────────────────────────────
    for ticker in approvals_need_qty:
        stock_info = get_stock_details(ticker)
        if not stock_info:
            print(f"[approval] Could not fetch details for {ticker}. Skipping.")
            continue
        send_qty_request_email(ticker, stock_info)

    for ticker in rejections:
        print(f"[approval] {ticker} skipped (rejected).")


if __name__ == "__main__":
    process_approvals()
