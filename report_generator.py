"""
report_generator.py
Generates HTML email report with:
  - FLAW 2.1: Consecutive BEARISH day count + sentiment flip warning
  - FLAW 2.3: Earnings alert per stock
  - FLAW 3.1: Cap category tag next to stock name
  - FLAW 3.4: PENDING SELL badge in decision column
  - FLAW 1.3: Insufficient indicators note in AI section
  - D.2:  Approaching stop loss warning in return column
"""

from datetime import datetime


def _color_for_decision(decision, urgency, pending_sell=False):
    if pending_sell or decision == 'SELL':
        return '#FF4444', '🔴'
    elif urgency == 'HIGH':
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

    # FLAW 3.4: Include pending sells in sell count
    sells   = [s for s in analyzed_stocks if s.get('decision') == 'SELL' or s.get('pending_sell')]
    watches = [s for s in analyzed_stocks if s.get('urgency') == 'MEDIUM' and s.get('decision') != 'SELL' and not s.get('pending_sell')]
    holds   = [s for s in analyzed_stocks if s.get('decision') == 'HOLD' and s.get('urgency') == 'LOW']

    total_investment = sum(s.get('investment_amt') or 0 for s in analyzed_stocks)
    total_current    = sum(s.get('current_value') or 0 for s in analyzed_stocks)
    total_profit     = total_current - total_investment
    total_growth     = ((total_profit / total_investment) * 100) if total_investment else 0

    # ── Stock rows ─────────────────────────────────────────────────────────────
    stock_rows = ''
    for s in analyzed_stocks:
        pending_sell = s.get('pending_sell', False)
        dec_color, dec_icon = _color_for_decision(s.get('decision'), s.get('urgency'), pending_sell)
        growth     = s.get('growth_pct')
        g_color    = _color_for_pct(growth)
        growth_str = f"{growth:+.2f}%" if growth is not None else "N/A"
        live       = s.get('live_price')
        live_str   = f"₹{live:,.2f}" if live else "N/A"
        profit     = s.get('total_profit')
        profit_str = f"₹{profit:,.0f}" if profit is not None else "N/A"

        # FLAW 3.1: Cap category tag
        cap      = s.get('cap_category', '')
        cap_html = f'<span style="font-size:10px;background:#f0f0f0;color:#666;padding:1px 6px;border-radius:8px;margin-left:4px;">{cap}</span>' if cap else ''

        # FLAW 2.1: Consecutive BEARISH day count in signal column
        consec = s.get('consecutive_bearish', 0)
        signal = s.get('technical_signal', 'N/A')
        if consec >= 2:
            signal_html = f'<span style="color:#CC0000;font-weight:bold;">{signal}<br><small style="font-size:10px;">BEARISH Day {consec} 🔴</small></span>'
        elif consec == 1:
            signal_html = f'<span style="color:#FF6600;">{signal}<br><small style="font-size:10px;">BEARISH Day 1 ⚠️</small></span>'
        else:
            sig_color   = '#007700' if signal == 'BULLISH' else '#CC0000' if signal == 'BEARISH' else '#888'
            signal_html = f'<span style="color:{sig_color};">{signal}</span>'

        # FLAW 3.4: PENDING SELL badge
        if pending_sell:
            dec_badge = '<span style="background:#FF4444;color:white;padding:3px 8px;border-radius:12px;font-size:11px;">⏳ PENDING SELL</span>'
        else:
            dec_badge = f'<span style="color:{dec_color};font-weight:bold;">{dec_icon} {s.get("decision","HOLD")}</span>'

        # D.2: Approaching stop loss warning
        stop_threshold  = s.get('stop_loss_threshold', -20)
        approaching_html = ''
        if growth is not None and stop_threshold and s.get('decision') != 'SELL' and not pending_sell:
            distance = growth - stop_threshold
            if 0 < distance <= 5:
                approaching_html = '<br><small style="color:#E65100;font-weight:bold;">⚠️ NEAR STOP LOSS</small>'

        # FLAW 2.3: Earnings alert in stock name cell
        earnings_html = ''
        ea = s.get('earnings_alert')
        if ea:
            earnings_html = f'<br><small style="color:#1565C0;font-weight:bold;">📅 {ea}</small>'

        stock_rows += f"""
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;">
            <strong>{s['ticker']}</strong>{cap_html}<br>
            <small style="color:#666;">{s.get('stock_name','')[:30]}</small>
            {earnings_html}
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">₹{s.get('buying_price',0):,.2f}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">{live_str}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{g_color};font-weight:bold;">{growth_str}{approaching_html}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{g_color};">{profit_str}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">{signal_html}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">{dec_badge}</td>
        </tr>"""

    # ── Fear & Greed Widget ────────────────────────────────────────────────────
    if fear_greed:
        fg_score  = fear_greed['score']
        fg_rating = fear_greed['rating']
        fg_emoji  = fear_greed['emoji']
        fg_color  = fear_greed['color']
        fg_advice = fear_greed['advice']
        fg_prev_c = fear_greed['prev_close']
        fg_prev_w = fear_greed['prev_week']
        fg_prev_m = fear_greed['prev_month']
        bar_width = int(fg_score)

        # Determine zone description
        if fg_score <= 25:
            zone_desc = "0–25 = Extreme Fear zone. Markets are oversold. Historically a good time to consider buying quality stocks."
            zone_color = '#C62828'
        elif fg_score <= 45:
            zone_desc = "26–45 = Fear zone. Investors are cautious. Selective buying opportunities may exist."
            zone_color = '#E65100'
        elif fg_score <= 55:
            zone_desc = "46–55 = Neutral zone. Normal market conditions. Stick to your trading rules."
            zone_color = '#FF9800'
        elif fg_score <= 75:
            zone_desc = "56–75 = Greed zone. Investors are optimistic. Be selective, avoid chasing rallies."
            zone_color = '#558B2F'
        else:
            zone_desc = "76–100 = Extreme Greed zone. Markets are overheated. Consider booking profits near targets."
            zone_color = '#2E7D32'

        # Trend indicator vs yesterday
        try:
            trend_vs_yesterday = int(fg_score) - int(fg_prev_c)
            if trend_vs_yesterday > 0:
                trend_html = f'<span style="color:#2E7D32;font-size:12px;">▲ +{trend_vs_yesterday} from yesterday</span>'
            elif trend_vs_yesterday < 0:
                trend_html = f'<span style="color:#C62828;font-size:12px;">▼ {trend_vs_yesterday} from yesterday</span>'
            else:
                trend_html = '<span style="color:#888;font-size:12px;">→ Unchanged from yesterday</span>'
        except:
            trend_html = ''

        fear_greed_html = f"""
        <div style="background:white;border-radius:12px;padding:20px;box-shadow:0 2px 6px rgba(0,0,0,0.1);margin-bottom:20px;border-top:4px solid {fg_color};">
          <h2 style="margin:0 0 16px;font-size:16px;">😱 Market Mood — Fear &amp; Greed Index</h2>

          <div style="display:flex;align-items:center;gap:24px;margin-bottom:16px;">
            <!-- Score block -->
            <div style="text-align:center;min-width:100px;">
              <div style="font-size:56px;line-height:1;">{fg_emoji}</div>
              <div style="font-size:44px;font-weight:bold;color:{fg_color};line-height:1.1;">{fg_score}</div>
              <div style="font-size:13px;font-weight:bold;color:{fg_color};margin-top:4px;">{fg_rating}</div>
              <div style="margin-top:6px;">{trend_html}</div>
            </div>

            <!-- Bar + scale -->
            <div style="flex:1;">
              <!-- Zone labels above bar -->
              <div style="display:flex;justify-content:space-between;font-size:10px;color:#888;margin-bottom:4px;">
                <span style="color:#C62828;font-weight:bold;">0<br>Extreme Fear</span>
                <span style="color:#E65100;">25<br>Fear</span>
                <span style="color:#FF9800;">50<br>Neutral</span>
                <span style="color:#7CB342;">75<br>Greed</span>
                <span style="color:#2E7D32;font-weight:bold;text-align:right;">100<br>Extreme Greed</span>
              </div>
              <!-- Gradient bar -->
              <div style="background:#eee;border-radius:20px;height:22px;position:relative;overflow:visible;">
                <div style="background:linear-gradient(to right,#C62828 0%,#FF5722 20%,#FF9800 40%,#FDD835 55%,#8BC34A 70%,#2E7D32 100%);border-radius:20px;height:22px;width:100%;"></div>
                <!-- Needle -->
                <div style="position:absolute;top:-5px;left:{bar_width}%;transform:translateX(-50%);z-index:2;">
                  <div style="width:4px;height:32px;background:#1a1a2e;border-radius:2px;box-shadow:0 2px 4px rgba(0,0,0,0.3);"></div>
                </div>
                <!-- Score bubble on needle -->
                <div style="position:absolute;top:-24px;left:{bar_width}%;transform:translateX(-50%);background:{fg_color};color:white;font-size:10px;font-weight:bold;padding:1px 5px;border-radius:8px;white-space:nowrap;">
                  {fg_score}
                </div>
              </div>
            </div>
          </div>

          <!-- Zone explanation -->
          <div style="background:#F8F9FA;border-left:3px solid {zone_color};padding:10px 14px;border-radius:4px;margin-bottom:12px;font-size:13px;">
            <strong>📊 Scale:</strong> {zone_desc}
          </div>

          <div style="background:#FFF8E1;padding:10px 14px;border-radius:6px;margin-bottom:12px;font-size:13px;">
            <strong>💡 AI Advice:</strong> {fg_advice}
          </div>

          <!-- Historical context -->
          <div style="display:flex;gap:10px;">
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:10px;border-radius:8px;">
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px;">Yesterday</div>
              <div style="font-size:20px;font-weight:bold;margin-top:4px;">{fg_prev_c}</div>
              <div style="font-size:10px;color:#888;">/ 100</div>
            </div>
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:10px;border-radius:8px;">
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px;">Last Week</div>
              <div style="font-size:20px;font-weight:bold;margin-top:4px;">{fg_prev_w}</div>
              <div style="font-size:10px;color:#888;">/ 100</div>
            </div>
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:10px;border-radius:8px;">
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px;">Last Month</div>
              <div style="font-size:20px;font-weight:bold;margin-top:4px;">{fg_prev_m}</div>
              <div style="font-size:10px;color:#888;">/ 100</div>
            </div>
          </div>
        </div>"""
    else:
        fear_greed_html = ""

    # ── Alert sections ─────────────────────────────────────────────────────────
    alert_section = ''
    if sells:
        alert_section += '<div style="background:#FFF0F0;border-left:4px solid #FF4444;padding:12px 16px;margin:12px 0;border-radius:4px;">'
        alert_section += '<strong>🔴 SELL / PENDING SELL ALERTS</strong><br>'
        for s in sells:
            pending_note = ' <em>(will execute at tomorrow\'s opening price)</em>' if s.get('pending_sell') else ''
            alert_section += f"<p><strong>{s['ticker']}</strong>{pending_note} — {s.get('reason','')}<br><em>{s.get('action_detail','')}</em></p>"
        alert_section += '</div>'

    if watches:
        alert_section += '<div style="background:#FFFAEE;border-left:4px solid #FF9900;padding:12px 16px;margin:12px 0;border-radius:4px;">'
        alert_section += '<strong>🟡 WATCH ALERTS</strong><br>'
        for s in watches:
            alert_section += f"<p><strong>{s['ticker']}</strong> — {s.get('reason','')}</p>"
        alert_section += '</div>'

    # ── AI News & Sentiment per stock ──────────────────────────────────────────
    ai_insights_html = ''
    for s in analyzed_stocks:
        ticker            = s.get('ticker', '')
        name              = s.get('stock_name', '')[:35]
        news_sentiment    = s.get('news_sentiment', 'NEUTRAL')
        overall_sentiment = s.get('overall_sentiment', 'NEUTRAL')
        recommendation    = s.get('recommendation', 'No analysis available.')
        risk              = s.get('risk_factors', 'N/A')
        articles          = s.get('headlines', [])
        earnings_alert    = s.get('earnings_alert')
        sentiment_flip    = s.get('sentiment_flip', 'NO')
        consec            = s.get('consecutive_bearish', 0)
        insuff            = s.get('insufficient_indicators', [])

        sent_color = '#2E7D32' if 'BULL' in overall_sentiment or 'POS' in news_sentiment else '#C62828' if 'BEAR' in overall_sentiment or 'NEG' in news_sentiment else '#E65100'
        sent_bg    = '#E8F5E9' if 'BULL' in overall_sentiment or 'POS' in news_sentiment else '#FFEBEE' if 'BEAR' in overall_sentiment or 'NEG' in news_sentiment else '#FFF3E0'

        # FLAW 2.1: Sentiment flip warning badge
        flip_html = ''
        if sentiment_flip == 'YES':
            flip_html = '<span style="background:#FFF3E0;color:#E65100;padding:2px 8px;border-radius:10px;font-size:11px;margin-left:8px;">⚠️ SENTIMENT FLIP</span>'

        # FLAW 2.1: Consecutive BEARISH label
        consec_html = ''
        if consec >= 2:
            consec_html = f'<span style="background:#FFEBEE;color:#C62828;padding:2px 8px;border-radius:10px;font-size:11px;margin-left:4px;">BEARISH Day {consec}</span>'

        # FLAW 2.2: News with tier indicator — full headline shown, no truncation
        headlines_html = ''
        if articles:
            for a in articles[:3]:
                tier_icon = {1: '🟢', 2: '🟡', 3: '🔴'}.get(a.get('tier', 3), '🔴')
                title     = a.get('title', '')
                date_str  = a.get('date', '')[:16] if a.get('date') else ''
                # Show complete headline — no truncation
                headlines_html += f'<li style="margin-bottom:8px;color:#333;line-height:1.5;">{tier_icon} {title}<br><small style="color:#aaa;">{date_str}</small></li>'
        else:
            headlines_html = '<li style="color:#999;">No recent news found</li>'

        # FLAW 1.3: Insufficient indicators note
        insuff_html = ''
        if insuff:
            insuff_html = f'<div style="font-size:11px;color:#888;margin-top:6px;">⚠️ Indicators skipped (insufficient data): {", ".join(insuff)}</div>'

        # FLAW 2.3: Earnings alert banner
        earn_html = ''
        if earnings_alert:
            earn_html = f'<div style="background:#FFF3CD;border:1px solid #FFC107;padding:8px;border-radius:4px;margin-bottom:8px;font-size:12px;font-weight:bold;">📅 {earnings_alert}</div>'

        ai_insights_html += f"""
        <div style="border:1px solid #eee;border-radius:8px;padding:14px;margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
            <div>
              <strong style="font-size:15px;">{ticker}</strong>
              <span style="color:#888;font-size:12px;margin-left:8px;">{name}</span>
              {flip_html}{consec_html}
            </div>
            <div style="background:{sent_bg};color:{sent_color};padding:4px 10px;border-radius:20px;font-size:12px;font-weight:bold;">
              {overall_sentiment} | News: {news_sentiment}
            </div>
          </div>
          {earn_html}
          <div style="background:#F8F9FA;padding:10px;border-radius:6px;margin-bottom:10px;">
            <strong style="font-size:12px;color:#333;">📰 Recent News (🟢 full article  🟡 snippet  🔴 headline only):</strong>
            <ul style="margin:6px 0 0;padding-left:18px;font-size:12px;">
              {headlines_html}
            </ul>
          </div>
          <div style="background:{sent_bg};padding:10px;border-radius:6px;margin-bottom:8px;">
            <strong style="font-size:12px;">🤖 AI Recommendation:</strong>
            <p style="margin:4px 0 0;font-size:13px;color:#333;">{recommendation}</p>
          </div>
          {insuff_html}
          <div style="font-size:11px;color:#888;margin-top:6px;">
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
          <th style="padding:10px 12px;text-align:center;font-size:12px;color:#666;border-bottom:2px solid #eee;">SIGNAL</th>
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
    sells        = [s['ticker'] for s in analyzed_stocks if s.get('decision') == 'SELL' or s.get('pending_sell')]
    total_profit = sum(s.get('total_profit') or 0 for s in analyzed_stocks)

    if sells:
        subject = f"⏳ PENDING SELL: {', '.join(sells)} | P&L: Rs.{total_profit:,.0f} | {datetime.now().strftime('%d %b %Y')}"
    elif total_profit >= 0:
        subject = f"Portfolio Update | Profit: Rs.{total_profit:,.0f} | {datetime.now().strftime('%d %b %Y')}"
    else:
        subject = f"Portfolio Update | Loss: Rs.{total_profit:,.0f} | {datetime.now().strftime('%d %b %Y')}"

    return subject