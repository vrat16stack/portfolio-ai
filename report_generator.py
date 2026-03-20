"""
report_generator.py
Generates a beautiful HTML email report with portfolio analysis results.
"""

from datetime import datetime


def _color_for_decision(decision, urgency):
    if decision == 'SELL':
        return '#FF4444', '🔴'
    elif urgency == 'MEDIUM':
        return '#FF9900', '🟡'
    else:
        return '#00AA44', '🟢'


def _color_for_pct(pct):
    if pct is None:
        return '#888888'
    return '#CC0000' if pct < 0 else '#007700'


def generate_html_report(analyzed_stocks, run_date=None, fear_greed=None):
    if run_date is None:
        run_date = datetime.now().strftime('%d %B %Y, %I:%M %p IST')

    sells = [s for s in analyzed_stocks if s.get('decision') == 'SELL']
    watches = [s for s in analyzed_stocks if s.get('urgency') == 'MEDIUM' and s.get('decision') != 'SELL']
    holds = [s for s in analyzed_stocks if s.get('decision') == 'HOLD' and s.get('urgency') == 'LOW']

    total_investment = sum(s.get('investment_amt') or 0 for s in analyzed_stocks)
    total_current = sum(s.get('current_value') or 0 for s in analyzed_stocks)
    total_profit = total_current - total_investment
    total_growth = ((total_profit / total_investment) * 100) if total_investment else 0

    # ── Stock rows ────────────────────────────────────────────
    stock_rows = ''
    for s in analyzed_stocks:
        dec_color, dec_icon = _color_for_decision(s.get('decision'), s.get('urgency'))
        growth = s.get('growth_pct')
        g_color = _color_for_pct(growth)
        growth_str = f"{growth:+.2f}%" if growth is not None else "N/A"
        live = s.get('live_price')
        live_str = f"₹{live:,.2f}" if live else "N/A"
        profit = s.get('total_profit')
        profit_str = f"₹{profit:,.0f}" if profit is not None else "N/A"

        stock_rows += f"""
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;">
            <strong>{s['ticker']}</strong><br>
            <small style="color:#666;">{s.get('stock_name','')[:30]}</small>
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">₹{s.get('buying_price',0):,.2f}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">{live_str}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{g_color};font-weight:bold;">{growth_str}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{g_color};">{profit_str}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">{s.get('technical_signal','N/A')}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{dec_color};font-weight:bold;">{dec_icon} {s.get('decision','HOLD')}</td>
        </tr>"""

    # ── Fear & Greed Widget ──────────────────────────────────
    if fear_greed:
        fg_score   = fear_greed['score']
        fg_rating  = fear_greed['rating']
        fg_emoji   = fear_greed['emoji']
        fg_color   = fear_greed['color']
        fg_advice  = fear_greed['advice']
        fg_prev_c  = fear_greed['prev_close']
        fg_prev_w  = fear_greed['prev_week']
        fg_prev_m  = fear_greed['prev_month']
        bar_width  = int(fg_score)

        fear_greed_html = f"""
        <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 2px 6px rgba(0,0,0,0.1);margin-bottom:20px;">
          <h2 style="margin:0 0 16px;font-size:16px;">Market Mood — Fear & Greed Index</h2>
          <div style="display:flex;align-items:center;gap:20px;margin-bottom:16px;">
            <div style="text-align:center;">
              <div style="font-size:48px;">{fg_emoji}</div>
              <div style="font-size:36px;font-weight:bold;color:{fg_color};">{fg_score}</div>
              <div style="font-size:14px;font-weight:bold;color:{fg_color};">{fg_rating}</div>
            </div>
            <div style="flex:1;">
              <div style="background:#eee;border-radius:20px;height:20px;margin-bottom:8px;position:relative;">
                <div style="background:linear-gradient(to right,#C62828,#FF9800,#2E7D32);border-radius:20px;height:20px;width:100%;"></div>
                <div style="position:absolute;top:-4px;left:{bar_width}%;transform:translateX(-50%);">
                  <div style="width:4px;height:28px;background:#333;border-radius:2px;"></div>
                </div>
              </div>
              <div style="display:flex;justify-content:space-between;font-size:11px;color:#888;">
                <span>Extreme Fear</span><span>Fear</span><span>Neutral</span><span>Greed</span><span>Extreme Greed</span>
              </div>
            </div>
          </div>
          <div style="background:#F8F9FA;padding:12px;border-radius:6px;margin-bottom:12px;">
            <strong>AI Advice:</strong> {fg_advice}
          </div>
          <div style="display:flex;gap:12px;">
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:8px;border-radius:6px;">
              <div style="font-size:11px;color:#888;">Yesterday</div>
              <div style="font-weight:bold;">{fg_prev_c}</div>
            </div>
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:8px;border-radius:6px;">
              <div style="font-size:11px;color:#888;">Last Week</div>
              <div style="font-weight:bold;">{fg_prev_w}</div>
            </div>
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:8px;border-radius:6px;">
              <div style="font-size:11px;color:#888;">Last Month</div>
              <div style="font-weight:bold;">{fg_prev_m}</div>
            </div>
          </div>
        </div>"""
    else:
        fear_greed_html = ""

    # ── Alert sections ────────────────────────────────────────
    alert_section = ''
    if sells:
        alert_section += '<div style="background:#FFF0F0;border-left:4px solid #FF4444;padding:12px 16px;margin:12px 0;border-radius:4px;">'
        alert_section += '<strong>🔴 SELL ALERTS</strong><br>'
        for s in sells:
            alert_section += f"<p><strong>{s['ticker']}</strong> — {s.get('reason','')}<br><em>{s.get('action_detail','')}</em></p>"
        alert_section += '</div>'

    if watches:
        alert_section += '<div style="background:#FFFAEE;border-left:4px solid #FF9900;padding:12px 16px;margin:12px 0;border-radius:4px;">'
        alert_section += '<strong>🟡 WATCH ALERTS</strong><br>'
        for s in watches:
            alert_section += f"<p><strong>{s['ticker']}</strong> — {s.get('reason','')}</p>"
        alert_section += '</div>'

    # ── AI News & Sentiment per stock ─────────────────────────
    ai_insights_html = ''
    for s in analyzed_stocks:
        ticker = s.get('ticker', '')
        name = s.get('stock_name', '')[:35]
        news_sentiment = s.get('news_sentiment', 'NEUTRAL')
        overall_sentiment = s.get('overall_sentiment', 'NEUTRAL')
        recommendation = s.get('recommendation', 'No analysis available.')
        risk = s.get('risk_factors', 'N/A')
        headlines = s.get('headlines', [])

        # Sentiment colors
        sent_color = '#2E7D32' if 'BULL' in overall_sentiment or 'POS' in news_sentiment else '#C62828' if 'BEAR' in overall_sentiment or 'NEG' in news_sentiment else '#E65100'
        sent_bg = '#E8F5E9' if 'BULL' in overall_sentiment or 'POS' in news_sentiment else '#FFEBEE' if 'BEAR' in overall_sentiment or 'NEG' in news_sentiment else '#FFF3E0'

        # News headlines
        headlines_html = ''
        if headlines:
            for h in headlines[:3]:
                headlines_html += f'<li style="margin-bottom:4px;color:#555;">{h["title"][:90]}{"..." if len(h["title"])>90 else ""}</li>'
        else:
            headlines_html = '<li style="color:#999;">No recent news found</li>'

        ai_insights_html += f"""
        <div style="border:1px solid #eee;border-radius:8px;padding:14px;margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
            <div>
              <strong style="font-size:15px;">{ticker}</strong>
              <span style="color:#888;font-size:12px;margin-left:8px;">{name}</span>
            </div>
            <div style="background:{sent_bg};color:{sent_color};padding:4px 10px;border-radius:20px;font-size:12px;font-weight:bold;">
              {overall_sentiment} | News: {news_sentiment}
            </div>
          </div>

          <div style="background:#F8F9FA;padding:10px;border-radius:6px;margin-bottom:10px;">
            <strong style="font-size:12px;color:#333;">📰 Recent News:</strong>
            <ul style="margin:6px 0 0;padding-left:18px;font-size:12px;">
              {headlines_html}
            </ul>
          </div>

          <div style="background:{sent_bg};padding:10px;border-radius:6px;margin-bottom:8px;">
            <strong style="font-size:12px;">🤖 AI Recommendation:</strong>
            <p style="margin:4px 0 0;font-size:13px;color:#333;">{recommendation}</p>
          </div>

          <div style="font-size:11px;color:#888;">
            ⚠️ <strong>Risk:</strong> {risk}
          </div>
        </div>"""

    profit_color = '#007700' if total_profit >= 0 else '#CC0000'
    growth_color = '#007700' if total_growth >= 0 else '#CC0000'

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:900px;margin:auto;background:#f5f5f5;padding:20px;">

  <div style="background:linear-gradient(135deg,#1a1a2e,#16213e);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">📊 Portfolio AI - Daily Analysis Report</h1>
    <p style="margin:6px 0 0;opacity:0.8;">Generated on {run_date}</p>
  </div>

  <div style="display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap;">
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:160px;">
      <div style="font-size:12px;color:#888;text-transform:uppercase;">Total Investment</div>
      <div style="font-size:22px;font-weight:bold;color:#333;">Rs.{total_investment:,.0f}</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:160px;">
      <div style="font-size:12px;color:#888;text-transform:uppercase;">Current Value</div>
      <div style="font-size:22px;font-weight:bold;color:#333;">Rs.{total_current:,.0f}</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:160px;">
      <div style="font-size:12px;color:#888;text-transform:uppercase;">Total P&L</div>
      <div style="font-size:22px;font-weight:bold;color:{profit_color};">Rs.{total_profit:,.0f}</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:160px;">
      <div style="font-size:12px;color:#888;text-transform:uppercase;">Overall Growth</div>
      <div style="font-size:22px;font-weight:bold;color:{growth_color};">{total_growth:+.2f}%</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:160px;">
      <div style="font-size:12px;color:#888;text-transform:uppercase;">Decisions</div>
      <div style="font-size:14px;font-weight:bold;">🔴 {len(sells)} SELL &nbsp; 🟡 {len(watches)} WATCH &nbsp; 🟢 {len(holds)} HOLD</div>
    </div>
  </div>

  {fear_greed_html}
  {alert_section}

  <div style="background:white;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);overflow:hidden;margin-bottom:20px;">
    <div style="padding:16px;border-bottom:1px solid #eee;">
      <h2 style="margin:0;font-size:16px;">📋 Holdings Analysis</h2>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr style="background:#f8f9fa;">
          <th style="padding:10px 12px;text-align:left;font-size:12px;color:#666;border-bottom:2px solid #eee;">STOCK</th>
          <th style="padding:10px 12px;text-align:center;font-size:12px;color:#666;border-bottom:2px solid #eee;">BUY PRICE</th>
          <th style="padding:10px 12px;text-align:center;font-size:12px;color:#666;border-bottom:2px solid #eee;">CURRENT</th>
          <th style="padding:10px 12px;text-align:center;font-size:12px;color:#666;border-bottom:2px solid #eee;">RETURN</th>
          <th style="padding:10px 12px;text-align:center;font-size:12px;color:#666;border-bottom:2px solid #eee;">P&L</th>
          <th style="padding:10px 12px;text-align:center;font-size:12px;color:#666;border-bottom:2px solid #eee;">TECHNICAL</th>
          <th style="padding:10px 12px;text-align:center;font-size:12px;color:#666;border-bottom:2px solid #eee;">DECISION</th>
        </tr>
      </thead>
      <tbody>{stock_rows}</tbody>
    </table>
  </div>

  <div style="background:white;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);padding:16px;margin-bottom:20px;">
    <h2 style="margin:0 0 16px;font-size:16px;">🤖 AI News & Sentiment Analysis</h2>
    {ai_insights_html}
  </div>

  <div style="text-align:center;color:#aaa;font-size:11px;padding:12px;">
    Generated by Portfolio AI. Not financial advice. Always do your own research.
  </div>

</body>
</html>"""

    return html


def generate_subject_line(analyzed_stocks):
    sells = [s['ticker'] for s in analyzed_stocks if s.get('decision') == 'SELL']
    total_profit = sum(s.get('total_profit') or 0 for s in analyzed_stocks)

    if sells:
        subject = f"SELL ALERT: {', '.join(sells)} | Portfolio P&L: Rs.{total_profit:,.0f}"
    elif total_profit >= 0:
        subject = f"Portfolio Update | Profit: Rs.{total_profit:,.0f} | {datetime.now().strftime('%d %b %Y')}"
    else:
        subject = f"Portfolio Update | Loss: Rs.{total_profit:,.0f} | {datetime.now().strftime('%d %b %Y')}"

    return subject