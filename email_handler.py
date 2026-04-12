"""
email_handler.py
Sends the daily portfolio report via Gmail SMTP.
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from config import EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER


def send_report_email(subject, html_body):
    """Send HTML email via Gmail"""
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER

        html_part = MIMEText(html_body, 'html')
        msg.attach(html_part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())

        print(f"[email_handler] ✅ Report sent to {EMAIL_RECEIVER}")
        return True

    except smtplib.SMTPAuthenticationError:
        print("[email_handler] ❌ Gmail authentication failed! Check EMAIL_PASSWORD in config.py")
        return False
    except Exception as e:
        print(f"[email_handler] ❌ Failed to send email: {e}")
        return False


def send_sell_alert(ticker, stock_name, reason, current_price, growth_pct):
    """Send an urgent sell alert immediately"""
    subject = f"🚨 URGENT SELL ALERT: {ticker} | {growth_pct:+.1f}%"
    html = f"""
    <div style="font-family:Arial;max-width:600px;margin:auto;">
      <div style="background:#FF4444;color:white;padding:20px;border-radius:8px;">
        <h1 style="margin:0;">🚨 SELL ALERT</h1>
        <p style="margin:4px 0 0;">{ticker} — {stock_name}</p>
      </div>
      <div style="padding:20px;background:white;border:1px solid #ddd;border-radius:0 0 8px 8px;">
        <table>
          <tr><td><strong>Current Price:</strong></td><td>₹{current_price:,.2f}</td></tr>
          <tr><td><strong>Growth:</strong></td><td style="color:{'red' if growth_pct < 0 else 'green'};">{growth_pct:+.2f}%</td></tr>
          <tr><td><strong>Reason:</strong></td><td>{reason}</td></tr>
        </table>
        <p style="background:#FFF0F0;padding:12px;border-radius:4px;border-left:4px solid red;">
          ⚠️ Please log in to your broker and review this position immediately.
        </p>
      </div>
    </div>
    """
    return send_report_email(subject, html)
