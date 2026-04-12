"""
weekly_monthly_summary.py
Sends:
- Weekly summary every Saturday morning
- Monthly summary on last day of every month

Changes:
  - FLAW 4.1: Data date stamp on every report + fresh yfinance price fetch
  - FLAW 4.2: Per-stock Nifty benchmark from each stock's individual buy date
  - FLAW 2.4: Weekly AI recommendation accuracy section
"""

import yfinance as yf
from datetime import datetime, timedelta, date
import calendar
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
    """FLAW 4.1: Always fetch fresh price from yfinance — not from cached Sheets value"""
    try:
        info  = yf.Ticker(ticker + NSE_SUFFIX).info
        price = info.get('currentPrice') or info.get('regularMarketPrice')
        return round(float(price), 2) if price else None
    except:
        return None


def get_last_data_date():
    """
    FLAW 4.1: Fetch actual last market close date from Nifty data.
    Used to stamp all reports with the correct data date.
    """
    try:
        nifty = yf.download("^NSEI", period="5d", interval="1d", progress=False)
        if not nifty.empty:
            return nifty.index[-1].strftime('%d %B %Y (%A)')
        return datetime.now().strftime('%d %B %Y')
    except:
        return datetime.now().strftime('%d %B %Y')


def get_nifty_price_on_date(buy_date_str):
    """
    FLAW 4.2: Fetch Nifty 50 closing price on a specific past date.
    Used to calculate per-stock Nifty benchmark from buy date.
    """
    try:
        nifty = yf.download("^NSEI", start=buy_date_str, period="5d", interval="1d", progress=False)
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
    """Period-level Nifty performance for overall comparison"""
    try:
        hist = yf.Ticker("^NSEI").history(period=f"{days+5}d")
        if len(hist) < 2:
            return None
        start = float(hist['Close'].iloc[-(days+1)] if len(hist) > days else hist['Close'].iloc[0])
        end   = float(hist['Close'].iloc[-1])
        return {'start': round(start,2), 'end': round(end,2), 'change_pct': round(((end-start)/start)*100, 2)}
    except:
        return None


def get_stock_period_change(ticker_yf, days=7):
    try:
        hist = yf.Ticker(ticker_yf).history(period=f"{days+5}d")
        if len(hist) < 2:
            return None
        start = float(hist['Close'].iloc[-(days+1)] if len(hist) > days else hist['Close'].iloc[0])
        end   = float(hist['Close'].iloc[-1])
        return {
            'start':      round(start, 2),
            'end':        round(end, 2),
            'change_pct': round(((end-start)/start)*100, 2),
            'change_abs': round(end-start, 2),
        }
    except:
        return None


# ── FLAW 2.4: Recommendation Accuracy ────────────────────────────────────────

def get_recommendation_accuracy():
    """
    Fetch last 30 days of recommendations from RecommendationsLog sheet.
    For each: fetch current price, calculate return since recommendation.
    Returns accuracy stats dict or None.
    """
    try:
        from sheets_handler import get_recommendations_log
        records = get_recommendations_log()
        if not records:
            return None

        today   = date.today()
        cutoff  = today - timedelta(days=30)
        recent  = []

        for rec in records:
            try:
                rec_date = datetime.strptime(str(rec.get('Date', ''))[:10], '%Y-%m-%d').date()
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

            ret_pct    = round(((current - rec_price) / rec_price) * 100, 2)
            target_hit = (current >= target) if target > 0 else False
            if ret_pct > 0:
                profitable += 1

            results.append({
                'ticker':     ticker,
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


# ── BUILD SUMMARY DATA ────────────────────────────────────────────────────────

def build_summary(period='weekly'):
    days         = 7 if period == 'weekly' else 30
    period_label = 'This Week' if period == 'weekly' else 'This Month'

    holdings = get_holdings()
    if not holdings:
        print(f"[summary] No holdings found!")
        return None

    # FLAW 4.1: Get actual last market data date
    data_date     = get_last_data_date()
    nifty_current = get_nifty_current_price()
    stocks_data   = []

    for stock in holdings:
        # FLAW 4.1: Always fetch fresh price
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

        # FLAW 4.2: Per-stock Nifty benchmark from buy date
        nifty_return_since_buy = None
        vs_nifty               = None
        try:
            if buying_date and nifty_current:
                nifty_on_buy = get_nifty_price_on_date(buying_date)
                if nifty_on_buy and nifty_on_buy > 0:
                    nifty_return_since_buy = round(((nifty_current - nifty_on_buy) / nifty_on_buy) * 100, 2)
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
            'nifty_return_since_buy': nifty_return_since_buy,  # FLAW 4.2
            'vs_nifty':               vs_nifty,                # FLAW 4.2
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
        'data_date':        data_date,   # FLAW 4.1
    }


# ── GENERATE EMAIL HTML ───────────────────────────────────────────────────────

def generate_summary_email(data, rec_accuracy=None):
    period_label = data['period_label']
    profit_color = '#2E7D32' if data['total_profit'] >= 0 else '#C62828'
    profit_bg    = '#E8F5E9' if data['total_profit'] >= 0 else '#FFEBEE'
    growth_color = '#2E7D32' if data['total_growth'] >= 0 else '#C62828'

    # FLAW 4.1: Data date stamp
    data_date_html = f"""
    <div style="background:#E3F2FD;border-left:4px solid #1976D2;padding:10px 16px;border-radius:4px;margin-bottom:16px;font-size:13px;">
      📅 <strong>Portfolio values as of market close on {data.get('data_date', data['date'])}</strong>
    </div>"""

    # Overall Nifty comparison (period level)
    nifty = data['nifty']
    if nifty:
        nifty_color = '#2E7D32' if nifty['change_pct'] >= 0 else '#C62828'
        vs_nifty    = round(data['total_growth'] - nifty['change_pct'], 2)
        vs_color    = '#2E7D32' if vs_nifty >= 0 else '#C62828'
        nifty_html  = f"""
        <div style="background:#F8F9FA;border-radius:10px;padding:16px;margin-bottom:20px;">
          <h3 style="margin:0 0 12px;font-size:14px;">📈 Your Portfolio vs Nifty 50 ({period_label})</h3>
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

    # FLAW 4.2: Stock rows with per-stock vs Nifty column
    stock_rows = ''
    for s in sorted(data['stocks'], key=lambda x: x['period_change_pct'] or 0, reverse=True):
        pc       = s['period_change_pct']
        pc_color = '#2E7D32' if pc and pc >= 0 else '#C62828'
        gc_color = '#2E7D32' if s['growth_pct'] >= 0 else '#C62828'
        tp_color = '#2E7D32' if s['total_profit'] >= 0 else '#C62828'
        cp       = s['current_price']

        # FLAW 4.2: vs Nifty cell
        vs_n     = s.get('vs_nifty')
        nifty_rb = s.get('nifty_return_since_buy')
        if vs_n is not None:
            vs_color   = '#2E7D32' if vs_n >= 0 else '#C62828'
            vs_nifty_td = (f'<td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;'
                           f'color:{vs_color};font-weight:bold;">'
                           f'{"+"+str(vs_n) if vs_n>=0 else str(vs_n)}%'
                           f'<br><small style="color:#888;font-size:10px;">'
                           f'Nifty: {"+"+str(nifty_rb) if nifty_rb and nifty_rb>=0 else str(nifty_rb)}%'
                           f'</small></td>')
        else:
            vs_nifty_td = '<td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:#aaa;">N/A</td>'

        stock_rows += f"""
        <tr>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;">
            <strong>{s['ticker']}</strong><br>
            <small style="color:#888;">{s['stock_name'][:25]}</small>
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">{s['industry'][:12]}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">
            Rs.{s['buying_price']:,.2f}<br>
            <small style="color:#888;font-size:10px;">{s['buying_date']}</small>
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">{('Rs.'+'{:,.2f}'.format(cp)) if cp else '—'}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">{s['qty']}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:{pc_color};font-weight:bold;">
            {('+' if pc>=0 else '')+str(pc)+'%' if pc is not None else '—'}
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:{gc_color};font-weight:bold;">
            {'+' if s['growth_pct']>=0 else ''}{s['growth_pct']}%
          </td>
          {vs_nifty_td}
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:{tp_color};font-weight:bold;">
            {'+' if s['total_profit']>=0 else ''}Rs.{s['total_profit']:,.0f}
          </td>
        </tr>"""

    # FLAW 2.4: AI Recommendation Accuracy section (weekly only)
    rec_accuracy_html = ''
    if rec_accuracy and data['period'] == 'weekly':
        acc_color = '#2E7D32' if rec_accuracy['accuracy_pct'] >= 50 else '#C62828'
        rec_rows  = ''
        for r in sorted(rec_accuracy['results'], key=lambda x: x['return_pct'], reverse=True)[:10]:
            r_color      = '#2E7D32' if r['return_pct'] >= 0 else '#C62828'
            target_badge = ' 🎯' if r['target_hit'] else ''
            rec_rows += f"""
            <tr>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;">{r['date']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:bold;">{r['ticker']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;">Rs.{r['rec_price']:,.2f}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;">Rs.{r['current']:,.2f}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;color:{r_color};font-weight:bold;">
                {'+' if r['return_pct']>=0 else ''}{r['return_pct']}%{target_badge}
              </td>
            </tr>"""

        rec_accuracy_html = f"""
        <div style="background:white;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,0.1);padding:16px;margin-bottom:20px;">
          <h2 style="margin:0 0 16px;font-size:16px;">🤖 AI Recommendation Accuracy — Last 30 Days</h2>
          <div style="display:flex;gap:12px;margin-bottom:16px;">
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;">Accuracy</div>
              <div style="font-size:28px;font-weight:bold;color:{acc_color};">{rec_accuracy['accuracy_pct']}%</div>
            </div>
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;">Profitable Picks</div>
              <div style="font-size:28px;font-weight:bold;color:#2E7D32;">{rec_accuracy['profitable']}</div>
            </div>
            <div style="flex:1;text-align:center;background:#F8F9FA;padding:12px;border-radius:8px;">
              <div style="font-size:11px;color:#888;">Total Picks</div>
              <div style="font-size:28px;font-weight:bold;">{rec_accuracy['total']}</div>
            </div>
          </div>
          <table style="width:100%;border-collapse:collapse;">
            <thead><tr style="background:#f8f9fa;">
              <th style="padding:8px 12px;text-align:left;font-size:11px;color:#666;">DATE</th>
              <th style="padding:8px 12px;text-align:left;font-size:11px;color:#666;">TICKER</th>
              <th style="padding:8px 12px;text-align:left;font-size:11px;color:#666;">REC PRICE</th>
              <th style="padding:8px 12px;text-align:left;font-size:11px;color:#666;">CURRENT</th>
              <th style="padding:8px 12px;text-align:left;font-size:11px;color:#666;">RETURN</th>
            </tr></thead>
            <tbody>{rec_rows}</tbody>
          </table>
        </div>"""

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:900px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#1a1a2e,#16213e);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">{'📅' if data['period']=='weekly' else '📆'} {period_label} Portfolio Summary</h1>
    <p style="margin:6px 0 0;opacity:0.8;">Generated on {data['date']}</p>
  </div>

  {data_date_html}

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
  {rec_accuracy_html}

  <div style="background:white;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,0.1);overflow:hidden;">
    <div style="padding:16px;border-bottom:1px solid #eee;">
      <h2 style="margin:0;font-size:16px;">📋 Holdings Performance — {period_label}</h2>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr style="background:#f8f9fa;">
          <th style="padding:10px 14px;text-align:left;font-size:11px;color:#666;border-bottom:2px solid #eee;">STOCK</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">INDUSTRY</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">BUY PRICE / DATE</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">CURRENT</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">QTY</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">{period_label.upper()} CHANGE</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">TOTAL RETURN</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;color:#666;border-bottom:2px solid #eee;">VS NIFTY (from buy)</th>
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
    data         = build_summary(period='weekly')
    rec_accuracy = get_recommendation_accuracy()
    if not data:
        print("[summary] No data — skipping.")
        return
    html    = generate_summary_email(data, rec_accuracy=rec_accuracy)
    subject = f"📅 Weekly Summary | P&L: {'+'if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f} | Data: {data.get('data_date', data['date'])}"
    send_report_email(subject, html)
    print("[summary] Weekly summary sent!")


def send_monthly_summary():
    print("[summary] Generating monthly summary...")
    data = build_summary(period='monthly')
    if not data:
        print("[summary] No data — skipping.")
        return
    html    = generate_summary_email(data)
    subject = f"📆 Monthly Summary | P&L: {'+'if data['total_profit']>=0 else ''}Rs.{data['total_profit']:,.0f} | Data: {data.get('data_date', data['date'])}"
    send_report_email(subject, html)
    print("[summary] Monthly summary sent!")


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