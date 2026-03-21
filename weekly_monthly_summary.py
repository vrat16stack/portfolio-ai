"""
weekly_monthly_summary.py
Sends:
- Weekly summary every Saturday morning
- Monthly summary on last day of every month
Works with both Google Sheets and Excel.
"""

import yfinance as yf
from datetime import datetime, timedelta
import calendar
from config import NSE_SUFFIX
from email_handler import send_report_email


def get_holdings():
    """Get holdings from Google Sheets or Excel depending on config"""
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


def get_nifty_performance(days=7):
    try:
        nifty = yf.Ticker("^NSEI")
        hist  = nifty.history(period=f"{days+5}d")
        if len(hist) < 2:
            return None
        start_price = float(hist['Close'].iloc[-(days+1)] if len(hist) > days else hist['Close'].iloc[0])
        end_price   = float(hist['Close'].iloc[-1])
        change_pct  = round(((end_price - start_price) / start_price) * 100, 2)
        return {'start': round(start_price,2), 'end': round(end_price,2), 'change_pct': change_pct}
    except:
        return None


def get_live_price(ticker):
    try:
        info  = yf.Ticker(ticker + NSE_SUFFIX).info
        price = info.get('currentPrice') or info.get('regularMarketPrice')
        return round(float(price), 2) if price else None
    except:
        return None


def get_stock_period_change(ticker_yf, days=7):
    try:
        stock = yf.Ticker(ticker_yf)
        hist  = stock.history(period=f"{days+5}d")
        if len(hist) < 2:
            return None
        start = float(hist['Close'].iloc[-(days+1)] if len(hist) > days else hist['Close'].iloc[0])
        end   = float(hist['Close'].iloc[-1])
        return {
            'start':      round(start, 2),
            'end':        round(end, 2),
            'change_pct': round(((end - start) / start) * 100, 2),
            'change_abs': round(end - start, 2),
        }
    except:
        return None


def build_summary(period='weekly'):
    days         = 7 if period == 'weekly' else 30
    period_label = 'This Week' if period == 'weekly' else 'This Month'

    holdings = get_holdings()
    if not holdings:
        print(f"[summary] No holdings found!")
        return None

    stocks_data = []
    for stock in holdings:
        live_price    = get_live_price(stock['ticker'])
        buying_price  = stock['buying_price']
        qty           = stock['qty']

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

        stocks_data.append({
            'ticker':            stock['ticker'],
            'stock_name':        stock.get('stock_name', stock['ticker']),
            'industry':          stock.get('industry', 'N/A'),
            'buying_price':      buying_price,
            'current_price':     live_price,
            'qty':               qty,
            'total_investment':  total_investment,
            'current_value':     current_value,
            'total_profit':      total_profit,
            'growth_pct':        growth_pct,
            'period_change_pct': perf['change_pct'] if perf else None,
            'period_change_abs': perf['change_abs'] if perf else None,
            'period_start':      perf['start'] if perf else None,
            'period_end':        perf['end'] if perf else None,
        })

    total_investment = sum(s['total_investment'] for s in stocks_data)
    total_current    = sum(s['current_value'] for s in stocks_data)
    total_profit     = round(total_current - total_investment, 2)
    total_growth     = round((total_profit / total_investment * 100), 2) if total_investment else 0

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
    }


def generate_summary_email(data):
    period_label  = data['period_label']
    profit_color  = '#2E7D32' if data['total_profit'] >= 0 else '#C62828'
    profit_bg     = '#E8F5E9' if data['total_profit'] >= 0 else '#FFEBEE'
    growth_color  = '#2E7D32' if data['total_growth'] >= 0 else '#C62828'

    # Nifty comparison
    nifty = data['nifty']
    if nifty:
        nifty_color  = '#2E7D32' if nifty['change_pct'] >= 0 else '#C62828'
        vs_nifty     = round(data['total_growth'] - nifty['change_pct'], 2)
        vs_color     = '#2E7D32' if vs_nifty >= 0 else '#C62828'
        nifty_html   = f"""
        <div style="background:#F8F9FA;border-radius:10px;padding:16px;margin-bottom:20px;">
          <h3 style="margin:0 0 12px;font-size:14px;">📈 Your Portfolio vs Nifty 50</h3>
          <div style="display:flex;gap:12px;">
            <div style="flex:1;text-align:center;background:white;padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;margin-bottom:4px;">Your Portfolio</div>
              <div style="font-size:22px;font-weight:bold;color:{growth_color};">{'+' if data['total_growth']>=0 else ''}{data['total_growth']}%</div>
            </div>
            <div style="flex:1;text-align:center;background:white;padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;margin-bottom:4px;">Nifty 50</div>
              <div style="font-size:22px;font-weight:bold;color:{nifty_color};">{'+' if nifty['change_pct']>=0 else ''}{nifty['change_pct']}%</div>
            </div>
            <div style="flex:1;text-align:center;background:white;padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;margin-bottom:4px;">vs Nifty</div>
              <div style="font-size:22px;font-weight:bold;color:{vs_color};">{'+' if vs_nifty>=0 else ''}{vs_nifty}%</div>
            </div>
          </div>
        </div>"""
    else:
        nifty_html = ""

    # Best/Worst
    best_worst_html = ""
    if data['best'] and data['worst']:
        b = data['best']
        w = data['worst']
        best_worst_html = f"""
        <div style="display:flex;gap:12px;margin-bottom:20px;">
          <div style="flex:1;background:#E8F5E9;border:1px solid #A5D6A7;border-radius:10px;padding:16px;">
            <div style="font-size:11px;color:#2E7D32;font-weight:bold;margin-bottom:6px;">🏆 BEST PERFORMER</div>
            <div style="font-size:18px;font-weight:bold;">{b['ticker']}</div>
            <div style="font-size:12px;color:#555;margin-bottom:8px;">{b['stock_name'][:30]}</div>
            <div style="font-size:24px;font-weight:bold;color:#2E7D32;">+{b['period_change_pct']}%</div>
            <div style="font-size:12px;color:#555;">Rs.{b['period_start']} → Rs.{b['period_end']}</div>
          </div>
          <div style="flex:1;background:#FFEBEE;border:1px solid #FFCDD2;border-radius:10px;padding:16px;">
            <div style="font-size:11px;color:#C62828;font-weight:bold;margin-bottom:6px;">📉 WORST PERFORMER</div>
            <div style="font-size:18px;font-weight:bold;">{w['ticker']}</div>
            <div style="font-size:12px;color:#555;margin-bottom:8px;">{w['stock_name'][:30]}</div>
            <div style="font-size:24px;font-weight:bold;color:#C62828;">{w['period_change_pct']}%</div>
            <div style="font-size:12px;color:#555;">Rs.{w['period_start']} → Rs.{w['period_end']}</div>
          </div>
        </div>"""

    # Stock rows
    stock_rows = ''
    for s in sorted(data['stocks'], key=lambda x: x['period_change_pct'] or 0, reverse=True):
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
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">{s['industry'][:15]}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">Rs.{s['buying_price']:,.2f}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">{('Rs.'+'{:,.2f}'.format(cp)) if cp else '—'}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">{s['qty']}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:{pc_color};font-weight:bold;">
            {('+' if pc>=0 else '')+str(pc)+'%' if pc is not None else '—'}
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:{gc_color};font-weight:bold;">
            {'+' if s['growth_pct']>=0 else ''}{s['growth_pct']}%
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:{tp_color};font-weight:bold;">
            {'+' if s['total_profit']>=0 else ''}Rs.{s['total_profit']:,.0f}
          </td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:900px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#1a1a2e,#16213e);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">{'📅' if data['period']=='weekly' else '📆'} {period_label} Portfolio Summary</h1>
    <p style="margin:6px 0 0;opacity:0.8;">Generated on {data['date']}</p>
  </div>

  <div style="display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap;">
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;">Total Investment</div>
      <div style="font-size:22px;font-weight:bold;">Rs.{data['total_investment']:,.0f}</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;">Current Value</div>
      <div style="font-size:22px;font-weight:bold;">Rs.{data['total_current']:,.0f}</div>
    </div>
    <div style="flex:1;background:{profit_bg};padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:{profit_color};text-transform:uppercase;">Total P&L</div>
      <div style="font-size:22px;font-weight:bold;color:{profit_color};">{'+' if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f}</div>
    </div>
    <div style="flex:1;background:{profit_bg};padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:{growth_color};text-transform:uppercase;">Overall Growth</div>
      <div style="font-size:22px;font-weight:bold;color:{growth_color};">{'+' if data['total_growth']>=0 else ''}{data['total_growth']}%</div>
    </div>
    <div style="flex:1;background:white;padding:16px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);min-width:150px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;">Holdings</div>
      <div style="font-size:22px;font-weight:bold;">{len(data['stocks'])} stocks</div>
    </div>
  </div>

  {nifty_html}
  {best_worst_html}

  <div style="background:white;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,0.1);overflow:hidden;">
    <div style="padding:16px;border-bottom:1px solid #eee;">
      <h2 style="margin:0;font-size:16px;">📋 Holdings Performance — {period_label}</h2>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr style="background:#f8f9fa;">
          <th style="padding:10px 14px;text-align:left;font-size:11px;color:#666;border-bottom:2px solid #eee;">STOCK</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">INDUSTRY</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">BUY PRICE</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">CURRENT</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">QTY</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">{period_label.upper()} CHANGE</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">TOTAL RETURN</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">P&L</th>
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


def send_weekly_summary():
    print("[summary] Generating weekly summary...")
    data = build_summary(period='weekly')
    if not data:
        print("[summary] No data — skipping.")
        return
    html    = generate_summary_email(data)
    subject = f"📅 Weekly Portfolio Summary | P&L: {'+'if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f} | {data['date']}"
    send_report_email(subject, html)
    print("[summary] ✅ Weekly summary sent!")


def send_monthly_summary():
    print("[summary] Generating monthly summary...")
    data = build_summary(period='monthly')
    if not data:
        print("[summary] No data — skipping.")
        return
    html    = generate_summary_email(data)
    subject = f"📆 Monthly Portfolio Summary | P&L: {'+'if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f} | {data['date']}"
    send_report_email(subject, html)
    print("[summary] ✅ Monthly summary sent!")


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
