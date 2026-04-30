"""
weekly_monthly_summary.py
Sends:
- Weekly summary every Saturday morning
- Monthly summary on last day of every month

Changes:
  - Q4 FIX: Separate AI accuracy email sent as PDF attachment
             Shows ALL 30-day picks (not just top 10) with sector column
             Sent independently from the weekly report
  - Q5 FIX: "VS NIFTY (from buy)" column removed from holdings table
             VS NIFTY data still computed internally for overall comparison
             but removed from the per-stock row table
  - FLAW 4.1: Data date stamp on every report + fresh yfinance price fetch
  - FLAW 4.2: Per-stock Nifty benchmark from each stock's individual buy date
  - FLAW 2.4: Weekly AI recommendation accuracy (now separate email + PDF)
"""

import yfinance as yf
from datetime import datetime, timedelta, date
import calendar
import io
from config import NSE_SUFFIX
from email_handler import send_report_email


def get_holdings():
    try:
        from config import USE_GOOGLE_SHEETS
        if USE_GOOGLE_SHEETS:
            from sheets_handler import read_holdings
        else:
            from excel_reader import read_holdings
        return read_holdings()
    except Exception as e:
        print(f"[summary] Error reading holdings: {e}")
        return []


def get_live_price(ticker):
    """FLAW 4.1: Always fetch fresh price from yfinance."""
    try:
        info  = yf.Ticker(ticker + NSE_SUFFIX).info
        price = info.get('currentPrice') or info.get('regularMarketPrice')
        return round(float(price), 2) if price else None
    except:
        return None


def get_stock_sector(ticker):
    """Fetch sector name for a ticker — used in accuracy PDF."""
    try:
        info = yf.Ticker(ticker + NSE_SUFFIX).info
        return info.get('sector', 'N/A') or 'N/A'
    except:
        return 'N/A'


def get_last_data_date():
    """FLAW 4.1: Fetch actual last market close date from Nifty data."""
    try:
        nifty = yf.download("^NSEI", period="5d", interval="1d", progress=False)
        if not nifty.empty:
            return nifty.index[-1].strftime('%d %B %Y (%A)')
        return datetime.now().strftime('%d %B %Y')
    except:
        return datetime.now().strftime('%d %B %Y')


def get_nifty_price_on_date(buy_date_str):
    """FLAW 4.2: Fetch Nifty 50 closing price on a specific past date."""
    try:
        nifty = yf.download("^NSEI", start=buy_date_str,
                             period="5d", interval="1d", progress=False)
        if not nifty.empty:
            return round(float(nifty['Close'].iloc[0]), 2)
        return None
    except:
        return None


def get_nifty_current_price():
    try:
        info = yf.Ticker("^NSEI").info
        return info.get('regularMarketPrice') or info.get('currentPrice')
    except:
        return None


def get_nifty_performance(days=7):
    """Period-level Nifty performance for overall comparison."""
    try:
        hist = yf.Ticker("^NSEI").history(period=f"{days+5}d")
        if len(hist) < 2:
            return None
        start = float(hist['Close'].iloc[-(days+1)]
                      if len(hist) > days else hist['Close'].iloc[0])
        end   = float(hist['Close'].iloc[-1])
        return {
            'start':      round(start, 2),
            'end':        round(end, 2),
            'change_pct': round(((end - start) / start) * 100, 2)
        }
    except:
        return None


def get_stock_period_change(ticker_yf, days=7):
    try:
        hist = yf.Ticker(ticker_yf).history(period=f"{days+5}d")
        if len(hist) < 2:
            return None
        start = float(hist['Close'].iloc[-(days+1)]
                      if len(hist) > days else hist['Close'].iloc[0])
        end   = float(hist['Close'].iloc[-1])
        return {
            'start':      round(start, 2),
            'end':        round(end, 2),
            'change_pct': round(((end - start) / start) * 100, 2),
            'change_abs': round(end - start, 2),
        }
    except:
        return None


# ── Recommendation Accuracy ────────────────────────────────────────────────────

def get_recommendation_accuracy():
    """
    Fetch last 30 days of recommendations from RecommendationsLog sheet.
    For each: fetch current price + sector, calculate return.
    Returns accuracy stats dict or None.
    """
    try:
        from sheets_handler import get_recommendations_log
        records = get_recommendations_log()
        if not records:
            return None

        today  = date.today()
        cutoff = today - timedelta(days=30)
        recent = []

        for rec in records:
            try:
                rec_date = datetime.strptime(
                    str(rec.get('Date', ''))[:10], '%Y-%m-%d').date()
                if rec_date >= cutoff:
                    recent.append(rec)
            except:
                continue

        if not recent:
            return None

        profitable = 0
        results    = []

        for rec in recent:
            ticker    = str(rec.get('Ticker', '')).upper()
            rec_price = float(str(rec.get('Recommended Price', 0) or 0))
            target    = float(str(rec.get('Target Price', 0) or 0))

            current = get_live_price(ticker)
            if not current or rec_price <= 0:
                continue

            # Q4: fetch sector for each ticker
            sector     = get_stock_sector(ticker)
            ret_pct    = round(((current - rec_price) / rec_price) * 100, 2)
            target_hit = (current >= target) if target > 0 else False

            if ret_pct > 0:
                profitable += 1

            results.append({
                'ticker':     ticker,
                'sector':     sector,
                'date':       str(rec.get('Date', ''))[:10],
                'rec_price':  rec_price,
                'current':    current,
                'return_pct': ret_pct,
                'target':     target,
                'target_hit': target_hit,
            })

        if not results:
            return None

        return {
            'total':        len(results),
            'profitable':   profitable,
            'accuracy_pct': round((profitable / len(results)) * 100, 1),
            'results':      results,
        }

    except Exception as e:
        print(f"[summary] Error calculating recommendation accuracy: {e}")
        return None


# ── Q4: Generate PDF of all 30-day picks ──────────────────────────────────────

def generate_accuracy_pdf(rec_accuracy):
    """
    Q4 FIX: Generate a PDF containing ALL 30-day picks with sector column.
    Returns PDF bytes or None.
    """
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                        Paragraph, Spacer)
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT

        buffer = io.BytesIO()
        doc    = SimpleDocTemplate(
            buffer,
            pagesize=landscape(A4),
            rightMargin=1.5*cm, leftMargin=1.5*cm,
            topMargin=1.5*cm,   bottomMargin=1.5*cm
        )

        styles   = getSampleStyleSheet()
        elements = []

        # ── Title ──────────────────────────────────────────────────────────
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=16,
            textColor=colors.HexColor('#1B3A6B'),
            spaceAfter=4,
        )
        sub_style = ParagraphStyle(
            'Sub',
            parent=styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#4A5568'),
            spaceAfter=12,
        )

        today_str = date.today().strftime('%d %B %Y')
        elements.append(Paragraph('Portfolio AI — AI Recommendation Accuracy Report', title_style))
        elements.append(Paragraph(
            f'Last 30 Days | Generated: {today_str} | '
            f'Total Picks: {rec_accuracy["total"]} | '
            f'Profitable: {rec_accuracy["profitable"]} | '
            f'Accuracy: {rec_accuracy["accuracy_pct"]}%',
            sub_style
        ))
        elements.append(Spacer(1, 0.3*cm))

        # ── Summary row ────────────────────────────────────────────────────
        acc_color = colors.HexColor('#2E7D32') if rec_accuracy['accuracy_pct'] >= 60 \
            else colors.HexColor('#E65100') if rec_accuracy['accuracy_pct'] >= 50 \
            else colors.HexColor('#C62828')

        summary_data = [[
            'TOTAL PICKS', 'PROFITABLE', 'UNPROFITABLE', 'ACCURACY',
            'TARGET HIT'
        ], [
            str(rec_accuracy['total']),
            str(rec_accuracy['profitable']),
            str(rec_accuracy['total'] - rec_accuracy['profitable']),
            f"{rec_accuracy['accuracy_pct']}%",
            str(sum(1 for r in rec_accuracy['results'] if r['target_hit'])),
        ]]

        summary_table = Table(summary_data, colWidths=[4*cm]*5)
        summary_table.setStyle(TableStyle([
            ('BACKGROUND',   (0, 0), (-1, 0), colors.HexColor('#1B3A6B')),
            ('TEXTCOLOR',    (0, 0), (-1, 0), colors.white),
            ('FONTNAME',     (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE',     (0, 0), (-1, 0), 9),
            ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
            ('BACKGROUND',   (0, 1), (-1, 1), colors.HexColor('#F0F4F8')),
            ('FONTNAME',     (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE',     (0, 1), (-1, 1), 12),
            ('TEXTCOLOR',    (3, 1), (3, 1), acc_color),
            ('GRID',         (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.HexColor('#1B3A6B'), colors.HexColor('#F0F4F8')]),
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 0.5*cm))

        # ── All picks table ────────────────────────────────────────────────
        header = ['#', 'DATE', 'TICKER', 'SECTOR', 'REC PRICE (₹)',
                  'CURRENT (₹)', 'TARGET (₹)', 'RETURN', 'STATUS']

        # Sort: profitable first (descending return), then unprofitable
        sorted_results = sorted(
            rec_accuracy['results'],
            key=lambda x: x['return_pct'],
            reverse=True
        )

        table_data = [header]
        for idx, r in enumerate(sorted_results, 1):
            status = '🎯 Target Hit' if r['target_hit'] else \
                     ('✅ Profit' if r['return_pct'] > 0 else '❌ Loss')
            table_data.append([
                str(idx),
                r['date'],
                r['ticker'],
                r['sector'][:20],
                f"{r['rec_price']:,.2f}",
                f"{r['current']:,.2f}",
                f"{r['target']:,.2f}" if r['target'] > 0 else 'N/A',
                f"{'+' if r['return_pct'] >= 0 else ''}{r['return_pct']}%",
                status,
            ])

        # Column widths for landscape A4
        col_widths = [0.8*cm, 2.2*cm, 2.4*cm, 4.5*cm,
                      3.0*cm, 3.0*cm, 3.0*cm, 2.4*cm, 3.0*cm]

        picks_table = Table(table_data, colWidths=col_widths, repeatRows=1)

        # Build row-level styles
        table_style = [
            # Header
            ('BACKGROUND',  (0, 0), (-1, 0), colors.HexColor('#1565C0')),
            ('TEXTCOLOR',   (0, 0), (-1, 0), colors.white),
            ('FONTNAME',    (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE',    (0, 0), (-1, 0), 8),
            ('ALIGN',       (0, 0), (-1, 0), 'CENTER'),
            # Data rows
            ('FONTSIZE',    (0, 1), (-1, -1), 8),
            ('ALIGN',       (0, 1), (2, -1), 'CENTER'),
            ('ALIGN',       (3, 1), (3, -1), 'LEFT'),
            ('ALIGN',       (4, 1), (-1, -1), 'CENTER'),
            ('GRID',        (0, 0), (-1, -1), 0.3, colors.HexColor('#DDDDDD')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
             [colors.HexColor('#FFFFFF'), colors.HexColor('#F7F9FC')]),
            ('TOPPADDING',  (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]

        # Colour return column per profit/loss
        for row_idx, r in enumerate(sorted_results, 1):
            ret_color = colors.HexColor('#2E7D32') if r['return_pct'] >= 0 \
                else colors.HexColor('#C62828')
            table_style.append(('TEXTCOLOR', (7, row_idx), (7, row_idx), ret_color))
            table_style.append(('FONTNAME',  (7, row_idx), (7, row_idx), 'Helvetica-Bold'))
            # Target hit rows get a light green background
            if r['target_hit']:
                table_style.append(
                    ('BACKGROUND', (0, row_idx), (-1, row_idx),
                     colors.HexColor('#E8F5E9'))
                )

        picks_table.setStyle(TableStyle(table_style))
        elements.append(picks_table)

        # Footer note
        elements.append(Spacer(1, 0.4*cm))
        note_style = ParagraphStyle(
            'Note',
            parent=styles['Normal'],
            fontSize=7,
            textColor=colors.HexColor('#888888'),
        )
        elements.append(Paragraph(
            'Accuracy = picks currently above recommended price / total picks × 100. '
            'Not financial advice. Generated by Portfolio AI.',
            note_style
        ))

        doc.build(elements)
        pdf_bytes = buffer.getvalue()
        buffer.close()
        return pdf_bytes

    except ImportError:
        print("[summary] reportlab not installed — PDF generation skipped. "
              "Install with: pip install reportlab --break-system-packages")
        return None
    except Exception as e:
        print(f"[summary] PDF generation error: {e}")
        return None


# ── Q4: Send separate accuracy email with PDF ─────────────────────────────────

def send_accuracy_email(rec_accuracy):
    """
    Q4 FIX: Sends a completely separate email for AI accuracy.
    Contains summary stats + PDF attachment with all 30-day picks.
    This runs on the same Saturday as the weekly report but is a separate email.
    """
    if not rec_accuracy:
        print("[summary] No accuracy data — skipping accuracy email.")
        return

    today_str = date.today().strftime('%d %B %Y')
    acc       = rec_accuracy['accuracy_pct']
    acc_color = '#2E7D32' if acc >= 60 else '#E65100' if acc >= 50 else '#C62828'
    acc_label = ('🟢 Good' if acc >= 60 else '🟡 Average' if acc >= 50 else '🔴 Needs Improvement')

    # Build sector breakdown
    sector_counts = {}
    sector_profits = {}
    for r in rec_accuracy['results']:
        sec = r.get('sector', 'N/A')
        sector_counts[sec]  = sector_counts.get(sec, 0) + 1
        if sec not in sector_profits:
            sector_profits[sec] = {'profitable': 0, 'total': 0}
        sector_profits[sec]['total'] += 1
        if r['return_pct'] > 0:
            sector_profits[sec]['profitable'] += 1

    sector_rows = ''
    for sec, counts in sorted(sector_profits.items(),
                               key=lambda x: x[1]['total'], reverse=True):
        sec_acc   = round((counts['profitable'] / counts['total']) * 100, 0)
        sec_color = '#2E7D32' if sec_acc >= 60 else '#E65100' if sec_acc >= 50 else '#C62828'
        sector_rows += f"""
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;">{sec}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">
            {counts['total']}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;
              color:#2E7D32;font-weight:bold;">{counts['profitable']}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;
              color:#C62828;">{counts['total'] - counts['profitable']}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;
              color:{sec_color};font-weight:bold;">{int(sec_acc)}%</td>
        </tr>"""

    # Top 5 and bottom 5 picks preview
    sorted_results = sorted(rec_accuracy['results'],
                            key=lambda x: x['return_pct'], reverse=True)
    top5    = sorted_results[:5]
    bottom5 = sorted_results[-5:]

    def mini_rows(picks, is_top=True):
        rows = ''
        for r in picks:
            color = '#2E7D32' if r['return_pct'] >= 0 else '#C62828'
            badge = ' 🎯' if r['target_hit'] else ''
            rows += (f"<tr><td style='padding:6px 10px;border-bottom:1px solid #eee;"
                     f"font-weight:bold;'>{r['ticker']}</td>"
                     f"<td style='padding:6px 10px;border-bottom:1px solid #eee;"
                     f"font-size:11px;color:#666;'>{r['sector'][:15]}</td>"
                     f"<td style='padding:6px 10px;border-bottom:1px solid #eee;"
                     f"color:{color};font-weight:bold;'>"
                     f"{'+' if r['return_pct']>=0 else ''}{r['return_pct']}%{badge}</td>"
                     f"<td style='padding:6px 10px;border-bottom:1px solid #eee;"
                     f"font-size:11px;color:#888;'>{r['date']}</td></tr>")
        return rows

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:820px;margin:auto;
             background:#f5f5f5;padding:20px;">

  <!-- Header -->
  <div style="background:linear-gradient(135deg,#0D47A1,#1565C0);color:white;
              padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">🤖 AI Recommendation Accuracy Report</h1>
    <p style="margin:6px 0 0;opacity:0.85;">Last 30 Days | {today_str}</p>
    <p style="margin:4px 0 0;opacity:0.7;font-size:12px;">
      Full list of all {rec_accuracy['total']} picks attached as PDF</p>
  </div>

  <!-- Summary cards -->
  <div style="display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap;">
    <div style="flex:1;background:white;padding:16px;border-radius:10px;
                box-shadow:0 2px 6px rgba(0,0,0,0.08);text-align:center;min-width:120px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;margin-bottom:6px;">
        Accuracy</div>
      <div style="font-size:32px;font-weight:bold;color:{acc_color};">{acc}%</div>
      <div style="font-size:12px;color:{acc_color};margin-top:4px;">{acc_label}</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:10px;
                box-shadow:0 2px 6px rgba(0,0,0,0.08);text-align:center;min-width:120px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;margin-bottom:6px;">
        Total Picks</div>
      <div style="font-size:32px;font-weight:bold;">{rec_accuracy['total']}</div>
    </div>
    <div style="flex:1;background:#E8F5E9;padding:16px;border-radius:10px;
                box-shadow:0 2px 6px rgba(0,0,0,0.08);text-align:center;min-width:120px;">
      <div style="font-size:11px;color:#2E7D32;text-transform:uppercase;margin-bottom:6px;">
        Profitable</div>
      <div style="font-size:32px;font-weight:bold;color:#2E7D32;">
        {rec_accuracy['profitable']}</div>
    </div>
    <div style="flex:1;background:#FFEBEE;padding:16px;border-radius:10px;
                box-shadow:0 2px 6px rgba(0,0,0,0.08);text-align:center;min-width:120px;">
      <div style="font-size:11px;color:#C62828;text-transform:uppercase;margin-bottom:6px;">
        Unprofitable</div>
      <div style="font-size:32px;font-weight:bold;color:#C62828;">
        {rec_accuracy['total'] - rec_accuracy['profitable']}</div>
    </div>
    <div style="flex:1;background:#FFF8E1;padding:16px;border-radius:10px;
                box-shadow:0 2px 6px rgba(0,0,0,0.08);text-align:center;min-width:120px;">
      <div style="font-size:11px;color:#E65100;text-transform:uppercase;margin-bottom:6px;">
        Targets Hit</div>
      <div style="font-size:32px;font-weight:bold;color:#E65100;">
        {sum(1 for r in rec_accuracy['results'] if r['target_hit'])}</div>
    </div>
  </div>

  <!-- Sector breakdown -->
  <div style="background:white;border-radius:10px;padding:16px;margin-bottom:20px;
              box-shadow:0 2px 6px rgba(0,0,0,0.08);">
    <h3 style="margin:0 0 14px;font-size:15px;color:#1565C0;">
      📊 Accuracy by Sector</h3>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr style="background:#1565C0;color:white;">
        <th style="padding:8px 12px;text-align:left;">Sector</th>
        <th style="padding:8px 12px;text-align:center;">Total</th>
        <th style="padding:8px 12px;text-align:center;">Profitable</th>
        <th style="padding:8px 12px;text-align:center;">Loss</th>
        <th style="padding:8px 12px;text-align:center;">Accuracy</th>
      </tr>
      {sector_rows}
    </table>
  </div>

  <!-- Top 5 and Bottom 5 -->
  <div style="display:flex;gap:16px;margin-bottom:20px;">
    <div style="flex:1;background:white;border-radius:10px;padding:16px;
                box-shadow:0 2px 6px rgba(0,0,0,0.08);">
      <h3 style="margin:0 0 12px;font-size:14px;color:#2E7D32;">🏆 Top 5 Picks</h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <tr style="background:#E8F5E9;">
          <th style="padding:6px 10px;text-align:left;">Ticker</th>
          <th style="padding:6px 10px;text-align:left;">Sector</th>
          <th style="padding:6px 10px;">Return</th>
          <th style="padding:6px 10px;">Date</th>
        </tr>
        {mini_rows(top5, True)}
      </table>
    </div>
    <div style="flex:1;background:white;border-radius:10px;padding:16px;
                box-shadow:0 2px 6px rgba(0,0,0,0.08);">
      <h3 style="margin:0 0 12px;font-size:14px;color:#C62828;">📉 Bottom 5 Picks</h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <tr style="background:#FFEBEE;">
          <th style="padding:6px 10px;text-align:left;">Ticker</th>
          <th style="padding:6px 10px;text-align:left;">Sector</th>
          <th style="padding:6px 10px;">Return</th>
          <th style="padding:6px 10px;">Date</th>
        </tr>
        {mini_rows(bottom5, False)}
      </table>
    </div>
  </div>

  <!-- PDF note -->
  <div style="background:#E3F2FD;border-left:4px solid #1565C0;padding:12px 16px;
              border-radius:4px;margin-bottom:16px;font-size:13px;">
    📎 <strong>Full list of all {rec_accuracy['total']} picks is attached as a PDF.</strong>
    It contains every recommendation from the last 30 days with date, ticker,
    sector, recommended price, current price, target price, and return.
  </div>

  <div style="text-align:center;color:#aaa;font-size:11px;padding:12px;">
    Generated by Portfolio AI | Not financial advice
  </div>

</body></html>"""

    # Generate PDF
    pdf_bytes = generate_accuracy_pdf(rec_accuracy)

    # Send with PDF attachment if available, else plain HTML
    if pdf_bytes:
        _send_email_with_pdf_attachment(
            subject=(f"🤖 AI Accuracy Report | {acc}% | "
                     f"{rec_accuracy['profitable']}/{rec_accuracy['total']} picks profitable "
                     f"| {today_str}"),
            html_body=html,
            pdf_bytes=pdf_bytes,
            pdf_filename=f"AI_Accuracy_{today_str.replace(' ', '_')}.pdf"
        )
    else:
        send_report_email(
            subject=(f"🤖 AI Accuracy Report | {acc}% | "
                     f"{rec_accuracy['profitable']}/{rec_accuracy['total']} profitable"),
            html_body=html
        )

    print("[summary] AI accuracy email sent.")


def _send_email_with_pdf_attachment(subject, html_body, pdf_bytes, pdf_filename):
    """Send HTML email with a PDF attachment via Gmail SMTP."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    from config import EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER

    try:
        msg = MIMEMultipart('mixed')
        msg['Subject'] = subject
        msg['From']    = EMAIL_SENDER
        msg['To']      = EMAIL_RECEIVER

        # HTML body
        alt  = MIMEMultipart('alternative')
        html = MIMEText(html_body, 'html')
        alt.attach(html)
        msg.attach(alt)

        # PDF attachment
        pdf_part = MIMEApplication(pdf_bytes, _subtype='pdf')
        pdf_part.add_header('Content-Disposition', 'attachment',
                             filename=pdf_filename)
        msg.attach(pdf_part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())

        print(f"[summary] Accuracy email with PDF sent to {EMAIL_RECEIVER}")

    except Exception as e:
        print(f"[summary] Error sending accuracy email with PDF: {e}")
        # Fallback — send HTML only
        send_report_email(subject, html_body)


# ── Build summary data ─────────────────────────────────────────────────────────

def build_summary(period='weekly'):
    days         = 7 if period == 'weekly' else 30
    period_label = 'This Week' if period == 'weekly' else 'This Month'

    holdings = get_holdings()
    if not holdings:
        print(f"[summary] No holdings found!")
        return None

    data_date     = get_last_data_date()
    nifty_current = get_nifty_current_price()
    stocks_data   = []

    for stock in holdings:
        live_price   = get_live_price(stock['ticker'])
        buying_price = stock['buying_price']
        qty          = stock['qty']
        buying_date  = str(stock.get('buying_date', ''))[:10]

        if live_price:
            total_investment = round(buying_price * qty, 2)
            current_value    = round(live_price * qty, 2)
            total_profit     = round(current_value - total_investment, 2)
            growth_pct       = round(((live_price - buying_price) / buying_price) * 100, 2)
        else:
            total_investment = round(buying_price * qty, 2)
            current_value    = total_investment
            total_profit     = 0
            growth_pct       = 0

        perf = get_stock_period_change(stock['ticker_yf'], days=days)

        # FLAW 4.2: Compute per-stock vs Nifty (used for overall comparison only)
        nifty_return_since_buy = None
        vs_nifty               = None
        try:
            if buying_date and nifty_current:
                nifty_on_buy = get_nifty_price_on_date(buying_date)
                if nifty_on_buy and nifty_on_buy > 0:
                    nifty_return_since_buy = round(
                        ((nifty_current - nifty_on_buy) / nifty_on_buy) * 100, 2)
                    vs_nifty = round(growth_pct - nifty_return_since_buy, 2)
        except:
            pass

        stocks_data.append({
            'ticker':                 stock['ticker'],
            'stock_name':             stock.get('stock_name', stock['ticker']),
            'industry':               stock.get('industry', 'N/A'),
            'buying_price':           buying_price,
            'buying_date':            buying_date,
            'current_price':          live_price,
            'qty':                    qty,
            'total_investment':       total_investment,
            'current_value':          current_value,
            'total_profit':           total_profit,
            'growth_pct':             growth_pct,
            'period_change_pct':      perf['change_pct'] if perf else None,
            'period_change_abs':      perf['change_abs'] if perf else None,
            'period_start':           perf['start'] if perf else None,
            'period_end':             perf['end'] if perf else None,
            'nifty_return_since_buy': nifty_return_since_buy,
            'vs_nifty':               vs_nifty,
        })

    total_investment = sum(s['total_investment'] for s in stocks_data)
    total_current    = sum(s['current_value'] for s in stocks_data)
    total_profit     = round(total_current - total_investment, 2)
    total_growth     = round((total_profit / total_investment * 100), 2) \
        if total_investment else 0

    valid = [s for s in stocks_data if s['period_change_pct'] is not None]
    best  = max(valid, key=lambda x: x['period_change_pct']) if valid else None
    worst = min(valid, key=lambda x: x['period_change_pct']) if valid else None
    nifty = get_nifty_performance(days=days)

    return {
        'period':           period,
        'period_label':     period_label,
        'stocks':           stocks_data,
        'total_investment': total_investment,
        'total_current':    total_current,
        'total_profit':     total_profit,
        'total_growth':     total_growth,
        'best':             best,
        'worst':            worst,
        'nifty':            nifty,
        'date':             datetime.now().strftime('%d %B %Y'),
        'data_date':        data_date,
    }


# ── Generate summary email HTML ────────────────────────────────────────────────

def generate_summary_email(data):
    """
    Q5 FIX: VS NIFTY (from buy) column removed from holdings table.
    VS NIFTY data is still used in the overall portfolio vs Nifty comparison
    block at the top — just removed from the per-stock row table.
    """
    period_label = data['period_label']
    profit_color = '#2E7D32' if data['total_profit'] >= 0 else '#C62828'
    profit_bg    = '#E8F5E9' if data['total_profit'] >= 0 else '#FFEBEE'
    growth_color = '#2E7D32' if data['total_growth'] >= 0 else '#C62828'

    # Data date stamp
    data_date_html = f"""
    <div style="background:#E3F2FD;border-left:4px solid #1976D2;padding:10px 16px;
                border-radius:4px;margin-bottom:16px;font-size:13px;">
      📅 <strong>Portfolio values as of market close on
      {data.get('data_date', data['date'])}</strong>
    </div>"""

    # Overall vs Nifty (period level)
    nifty = data['nifty']
    if nifty:
        nifty_color = '#2E7D32' if nifty['change_pct'] >= 0 else '#C62828'
        vs_nifty    = round(data['total_growth'] - nifty['change_pct'], 2)
        vs_color    = '#2E7D32' if vs_nifty >= 0 else '#C62828'
        nifty_html  = f"""
        <div style="background:#F8F9FA;border-radius:10px;padding:16px;margin-bottom:20px;">
          <h3 style="margin:0 0 12px;font-size:14px;">
            📈 Your Portfolio vs Nifty 50 ({period_label})</h3>
          <div style="display:flex;gap:12px;">
            <div style="flex:1;text-align:center;background:white;
                        padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;margin-bottom:4px;">Your Portfolio</div>
              <div style="font-size:22px;font-weight:bold;color:{growth_color};">
                {'+' if data['total_growth']>=0 else ''}{data['total_growth']}%</div>
            </div>
            <div style="flex:1;text-align:center;background:white;
                        padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;margin-bottom:4px;">Nifty 50</div>
              <div style="font-size:22px;font-weight:bold;color:{nifty_color};">
                {'+' if nifty['change_pct']>=0 else ''}{nifty['change_pct']}%</div>
            </div>
            <div style="flex:1;text-align:center;background:white;
                        padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;margin-bottom:4px;">vs Nifty</div>
              <div style="font-size:22px;font-weight:bold;color:{vs_color};">
                {'+' if vs_nifty>=0 else ''}{vs_nifty}%</div>
            </div>
          </div>
        </div>"""
    else:
        nifty_html = ""

    # Best / Worst
    best_worst_html = ""
    if data['best'] and data['worst']:
        b = data['best']
        w = data['worst']
        best_worst_html = f"""
        <div style="display:flex;gap:12px;margin-bottom:20px;">
          <div style="flex:1;background:#E8F5E9;border:1px solid #A5D6A7;
                      border-radius:10px;padding:16px;">
            <div style="font-size:11px;color:#2E7D32;font-weight:bold;margin-bottom:6px;">
              🏆 BEST PERFORMER</div>
            <div style="font-size:18px;font-weight:bold;">{b['ticker']}</div>
            <div style="font-size:12px;color:#555;margin-bottom:8px;">
              {b['stock_name'][:30]}</div>
            <div style="font-size:24px;font-weight:bold;color:#2E7D32;">
              +{b['period_change_pct']}%</div>
            <div style="font-size:12px;color:#555;">
              Rs.{b['period_start']} → Rs.{b['period_end']}</div>
          </div>
          <div style="flex:1;background:#FFEBEE;border:1px solid #FFCDD2;
                      border-radius:10px;padding:16px;">
            <div style="font-size:11px;color:#C62828;font-weight:bold;margin-bottom:6px;">
              📉 WORST PERFORMER</div>
            <div style="font-size:18px;font-weight:bold;">{w['ticker']}</div>
            <div style="font-size:12px;color:#555;margin-bottom:8px;">
              {w['stock_name'][:30]}</div>
            <div style="font-size:24px;font-weight:bold;color:#C62828;">
              {w['period_change_pct']}%</div>
            <div style="font-size:12px;color:#555;">
              Rs.{w['period_start']} → Rs.{w['period_end']}</div>
          </div>
        </div>"""

    # Q5 FIX: Stock rows WITHOUT "VS NIFTY (from buy)" column
    stock_rows = ''
    for s in sorted(data['stocks'],
                    key=lambda x: x['period_change_pct'] or 0, reverse=True):
        pc       = s['period_change_pct']
        pc_color = '#2E7D32' if pc and pc >= 0 else '#C62828'
        gc_color = '#2E7D32' if s['growth_pct'] >= 0 else '#C62828'
        tp_color = '#2E7D32' if s['total_profit'] >= 0 else '#C62828'
        cp       = s['current_price']

        stock_rows += f"""
        <tr>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;">
            <strong>{s['ticker']}</strong><br>
            <small style="color:#888;">{s['stock_name'][:25]}</small>
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">
            {s['industry'][:12]}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">
            Rs.{s['buying_price']:,.2f}<br>
            <small style="color:#888;font-size:10px;">{s['buying_date']}</small>
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">
            {('Rs.'+'{:,.2f}'.format(cp)) if cp else '—'}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">
            {s['qty']}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;
              color:{pc_color};font-weight:bold;">
            {('+' if pc >= 0 else '') + str(pc) + '%' if pc is not None else '—'}
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;
              color:{gc_color};font-weight:bold;">
            {'+' if s['growth_pct'] >= 0 else ''}{s['growth_pct']}%
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;
              color:{tp_color};font-weight:bold;">
            {'+' if s['total_profit'] >= 0 else ''}Rs.{s['total_profit']:,.0f}
          </td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:900px;margin:auto;
             background:#f5f5f5;padding:20px;">

  <div style="background:linear-gradient(135deg,#1a1a2e,#16213e);color:white;
              padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">
      {'📅' if data['period']=='weekly' else '📆'} {period_label} Portfolio Summary</h1>
    <p style="margin:6px 0 0;opacity:0.8;">Generated on {data['date']}</p>
  </div>

  {data_date_html}

  <div style="display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap;">
    <div style="flex:1;background:white;padding:16px;border-radius:8px;
                box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;">Total Investment</div>
      <div style="font-size:22px;font-weight:bold;">Rs.{data['total_investment']:,.0f}</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;
                box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;">Current Value</div>
      <div style="font-size:22px;font-weight:bold;">Rs.{data['total_current']:,.0f}</div>
    </div>
    <div style="flex:1;background:{profit_bg};padding:16px;border-radius:8px;
                box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:{profit_color};text-transform:uppercase;">Total P&L</div>
      <div style="font-size:22px;font-weight:bold;color:{profit_color};">
        {'+' if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f}</div>
    </div>
    <div style="flex:1;background:{profit_bg};padding:16px;border-radius:8px;
                box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:{growth_color};text-transform:uppercase;">
        Overall Growth</div>
      <div style="font-size:22px;font-weight:bold;color:{growth_color};">
        {'+' if data['total_growth']>=0 else ''}{data['total_growth']}%</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;
                box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;">Holdings</div>
      <div style="font-size:22px;font-weight:bold;">{len(data['stocks'])} stocks</div>
    </div>
  </div>

  {nifty_html}
  {best_worst_html}

  <!-- Q5 FIX: Holdings table WITHOUT VS NIFTY column -->
  <div style="background:white;border-radius:10px;
              box-shadow:0 2px 6px rgba(0,0,0,0.1);overflow:hidden;">
    <div style="padding:16px;border-bottom:1px solid #eee;">
      <h2 style="margin:0;font-size:16px;">
        📋 Holdings Performance — {period_label}</h2>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr style="background:#f8f9fa;">
          <th style="padding:10px 14px;text-align:left;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">STOCK</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">INDUSTRY</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">BUY PRICE / DATE</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">CURRENT</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">QTY</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">{period_label.upper()} CHANGE</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">TOTAL RETURN</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;
              border-bottom:2px solid #eee;">P&L</th>
        </tr>
      </thead>
      <tbody>{stock_rows}</tbody>
    </table>
  </div>

  <div style="text-align:center;color:#aaa;font-size:11px;padding:12px;">
    Generated by Portfolio AI | Not financial advice
  </div>
</body>
</html>"""
    return html


# ── Public send functions ──────────────────────────────────────────────────────

def send_weekly_summary():
    """
    Saturday job:
    1. Sends the weekly portfolio summary email (no accuracy section inside)
    2. Separately sends the AI accuracy email with PDF attachment
    """
    print("[summary] Generating weekly summary...")
    data = build_summary(period='weekly')
    if not data:
        print("[summary] No data — skipping.")
        return

    # Q5: Weekly summary email — no accuracy section, no VS NIFTY column
    html    = generate_summary_email(data)
    subject = (f"📅 Weekly Summary | "
               f"P&L: {'+'if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f} | "
               f"Data: {data.get('data_date', data['date'])}")
    send_report_email(subject, html)
    print("[summary] Weekly summary email sent.")

    # Q4: Separate accuracy email with PDF
    print("[summary] Generating AI accuracy email...")
    rec_accuracy = get_recommendation_accuracy()
    send_accuracy_email(rec_accuracy)


def send_monthly_summary():
    print("[summary] Generating monthly summary...")
    data = build_summary(period='monthly')
    if not data:
        print("[summary] No data — skipping.")
        return
    html    = generate_summary_email(data)
    subject = (f"📆 Monthly Summary | "
               f"P&L: {'+'if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f} | "
               f"Data: {data.get('data_date', data['date'])}")
    send_report_email(subject, html)
    print("[summary] Monthly summary sent.")


def should_send_weekly():
    return datetime.now().weekday() == 5


def should_send_monthly():
    today    = datetime.now()
    last_day = calendar.monthrange(today.year, today.month)[1]
    return today.day == last_day


if __name__ == "__main__":
    import sys
    if '--weekly' in sys.argv:
        send_weekly_summary()
    elif '--monthly' in sys.argv:
        send_monthly_summary()
    else:
        send_weekly_summary()
