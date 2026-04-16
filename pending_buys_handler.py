"""
pending_buys_handler.py
Pre-market buy decision engine for Portfolio AI.

Flow:
  9:00 AM IST — approval_checker.py reads Gmail YES replies
                → stores PENDING BUY in PendingBuys sheet (via store_pending_buy)

  9:20 AM IST — main.py --morning calls execute_pending_buys()
                → fetches actual open price
                → applies gap decision tree
                → sends Buy Executed or Buy Cancelled email
                → adds confirmed buys to Holdings sheet

Gap Decision Tree (per industry-specific thresholds):
  gap_pct < -threshold          → fetch overnight news
                                    bad news  → CANCEL
                                    no bad news → BUY (lower entry = more upside)
  -threshold ≤ gap_pct ≤ +threshold → BUY at open, keep original target
  gap_pct > +threshold AND
      gap_pct < upside_pct×0.8  → BUY, revise target upward, justify fully in email
  gap_pct ≥ upside_pct×0.8     → CANCEL (risk/reward destroyed, target nearly hit)
"""

import yfinance as yf
import time
from datetime import datetime, date, timedelta
from config import NSE_SUFFIX, EMAIL_SENDER, EMAIL_PASSWORD
from sheets_handler import (
    get_or_create_worksheet,
    add_stock_to_holdings,
    read_holdings,
)
from email_handler import send_report_email
from technical_analysis import calculate_indicators
from price_fetcher import get_historical_data

# ── Industry-Specific Gap Thresholds ──────────────────────────────────────────
SECTOR_THRESHOLDS = {
    'Metals':               5.0,
    'Mining':               5.0,
    'Basic Materials':      5.0,
    'Energy':               4.0,
    'Utilities':            4.0,
    'Financial Services':   4.0,
    'Banking':              4.0,
    'Finance':              4.0,
    'Healthcare':           4.0,
    'Pharmaceuticals':      4.0,
    'Consumer Defensive':   3.0,
    'Consumer Staples':     3.0,
    'FMCG':                 3.0,
    'Technology':           3.0,
    'Information Technology': 3.0,
    'Communication Services': 3.0,
    'Industrials':          4.0,
    'Real Estate':          4.0,
    'Consumer Cyclical':    4.0,
}
DEFAULT_THRESHOLD    = 4.0
SMALL_CAP_THRESHOLD  = 6.0   # overrides sector threshold for small caps

PENDING_BUYS_SHEET   = "PendingBuys"
PENDING_BUYS_HEADERS = [
    'Date Added', 'Ticker', 'Stock Name', 'Qty',
    'Scout Price', 'Original Target', 'Upside %',
    'Sector', 'Cap Category',
    'Status',           # PENDING / EXECUTED / CANCELLED
    'Gap %', 'Actual Buy Price', 'Revised Target', 'Reason',
]


# ── Sheet Helpers ──────────────────────────────────────────────────────────────

def _get_pending_buys_ws():
    ws = get_or_create_worksheet(PENDING_BUYS_SHEET)
    if ws is None:
        return None
    existing = ws.get_all_values()
    if not existing or existing[0] != PENDING_BUYS_HEADERS:
        ws.clear()
        ws.insert_row(PENDING_BUYS_HEADERS, 1)
        print(f"[pending_buys] Initialised {PENDING_BUYS_SHEET} sheet with headers.")
    return ws


def store_pending_buy(ticker, stock_name, qty, scout_price, original_target,
                      sector, cap_category):
    """
    Called by approval_checker at 9:00 AM after reading YES reply.
    Writes a PENDING row to PendingBuys sheet.
    Does NOT add to Holdings yet — that happens after gap check at 9:20 AM.
    Skips duplicate if ticker already PENDING.
    """
    ws = _get_pending_buys_ws()
    if ws is None:
        return False

    # Duplicate check — don't add if already PENDING for same ticker
    records = ws.get_all_records()
    for row in records:
        if (str(row.get('Ticker', '')).upper() == ticker.upper()
                and str(row.get('Status', '')).upper() == 'PENDING'):
            print(f"[pending_buys] {ticker} already PENDING — skipping duplicate.")
            return False

    upside_pct = round(((original_target - scout_price) / scout_price) * 100, 2) \
        if scout_price and original_target else 0.0

    ws.append_row([
        date.today().strftime('%Y-%m-%d'),
        ticker.upper(),
        stock_name,
        qty,
        round(scout_price, 2),
        round(original_target, 2),
        upside_pct,
        sector,
        cap_category,
        'PENDING',
        '',   # Gap % — filled at 9:20 AM
        '',   # Actual Buy Price
        '',   # Revised Target
        '',   # Reason
    ])
    print(f"[pending_buys] ✅ Stored PENDING BUY: {ticker} x{qty} @ scout ₹{scout_price}")
    return True


def get_pending_buys():
    """Return all rows with Status == PENDING"""
    ws = _get_pending_buys_ws()
    if ws is None:
        return []
    records = ws.get_all_records()
    pending = [r for r in records
               if str(r.get('Status', '')).upper() == 'PENDING']
    if pending:
        print(f"[pending_buys] Found {len(pending)} pending buy(s): "
              f"{[p['Ticker'] for p in pending]}")
    return pending


def _update_pending_row(ws, ticker, updates: dict):
    """Update specific columns on the first PENDING row matching ticker."""
    try:
        records = ws.get_all_records()
        headers = ws.row_values(1)
        for i, row in enumerate(records, start=2):
            if (str(row.get('Ticker', '')).upper() == ticker.upper()
                    and str(row.get('Status', '')).upper() == 'PENDING'):
                for col_name, value in updates.items():
                    try:
                        col_idx = headers.index(col_name) + 1
                        ws.update_cell(i, col_idx, value)
                    except ValueError:
                        pass
                return True
    except Exception as e:
        print(f"[pending_buys] Error updating row for {ticker}: {e}")
    return False


# ── Price Fetching ─────────────────────────────────────────────────────────────

def _get_open_price(ticker_yf):
    """
    Fetch today's actual opening price.
    Primary: yfinance 1d interval (most reliable after 9:15 AM).
    Fallback: previousClose if open not available yet.
    """
    try:
        stock = yf.Ticker(ticker_yf)

        # Try intraday 1-minute data first — open of first candle = market open
        intraday = stock.history(period='1d', interval='1m')
        if intraday is not None and not intraday.empty:
            open_price = round(float(intraday['Open'].iloc[0]), 2)
            if open_price > 0:
                print(f"[pending_buys] {ticker_yf} open price (intraday): ₹{open_price}")
                return open_price

        # Fallback: day-level history
        hist = stock.history(period='2d', interval='1d')
        if hist is not None and not hist.empty:
            open_price = round(float(hist['Open'].iloc[-1]), 2)
            if open_price > 0:
                print(f"[pending_buys] {ticker_yf} open price (daily): ₹{open_price}")
                return open_price

        # Last resort: current/previous close
        info  = stock.info
        price = info.get('regularMarketOpen') or info.get('currentPrice') or info.get('previousClose')
        if price:
            return round(float(price), 2)

    except Exception as e:
        print(f"[pending_buys] Error fetching open price for {ticker_yf}: {e}")
    return None


def _get_previous_close(ticker_yf):
    """Yesterday's closing price — used as base for gap calculation."""
    try:
        stock = yf.Ticker(ticker_yf)
        hist  = stock.history(period='5d', interval='1d')
        if hist is not None and len(hist) >= 2:
            return round(float(hist['Close'].iloc[-2]), 2)
        info  = stock.info
        prev  = info.get('previousClose')
        return round(float(prev), 2) if prev else None
    except Exception as e:
        print(f"[pending_buys] Error fetching prev close for {ticker_yf}: {e}")
        return None


# ── Gap Threshold Logic ────────────────────────────────────────────────────────

def _get_threshold(sector, cap_category):
    """Return the gap threshold % for this stock's sector and cap category."""
    if cap_category and 'small' in cap_category.lower():
        return SMALL_CAP_THRESHOLD
    for key, val in SECTOR_THRESHOLDS.items():
        if key.lower() in str(sector).lower():
            return val
    return DEFAULT_THRESHOLD


# ── News Fetching for Gap-Down Check ──────────────────────────────────────────

def _fetch_overnight_news(ticker):
    """
    Fetch recent news headlines for ticker via yfinance.
    Returns (has_bad_news: bool, headlines: list[str], bad_keywords_found: list[str])
    """
    BAD_KEYWORDS = [
        'fraud', 'scam', 'bankruptcy', 'bankrupt', 'default', 'insolvency',
        'investigation', 'raid', 'sebi', 'ed ', 'cbi', 'probe', 'penalty',
        'fine', 'downgrade', 'suspend', 'delist', 'loss', 'collapse',
        'shutdown', 'closure', 'resign', 'fired', 'lawsuit', 'legal action',
        'miss', 'missed', 'disappointing', 'warning', 'lower guidance',
        'profit warning', 'negative', 'decline', 'drop', 'fall', 'plunge',
    ]

    try:
        stock    = yf.Ticker(ticker + NSE_SUFFIX)
        news     = stock.news or []
        headlines = []
        bad_found = []

        for item in news[:10]:   # check latest 10 news items
            title = str(item.get('title', '') or item.get('content', {}).get('title', '') or '')
            if title:
                headlines.append(title)
                title_lower = title.lower()
                for kw in BAD_KEYWORDS:
                    if kw in title_lower and kw not in bad_found:
                        bad_found.append(kw)

        has_bad = len(bad_found) > 0
        print(f"[pending_buys] {ticker} news check: {len(headlines)} items, "
              f"bad keywords: {bad_found if bad_found else 'none'}")
        return has_bad, headlines[:5], bad_found

    except Exception as e:
        print(f"[pending_buys] News fetch error for {ticker}: {e}")
        return False, [], []


# ── Target Revision Logic ──────────────────────────────────────────────────────

def _compute_revised_target(ticker_yf, open_price, original_target,
                             scout_price, gap_pct, indicators):
    """
    Compute a data-driven revised target on gap up.
    Formula: original_target + (gap_amount × 0.5)
    Then cross-check against EMA resistance, BB upper, ADX momentum.
    Returns (revised_target, justification_dict)
    """
    gap_amount     = open_price - scout_price
    base_revised   = round(original_target + (gap_amount * 0.5), 2)

    ema50          = indicators.get('ema50')
    ema200         = indicators.get('ema200')
    bb_upper       = indicators.get('bb_upper')
    adx            = indicators.get('adx')
    rsi            = indicators.get('rsi')
    macd           = indicators.get('macd')
    macd_signal    = indicators.get('macd_signal')
    bull_pct       = indicators.get('bull_pct', 50)

    justification  = {
        'formula':        f"Original Target ₹{original_target} + (Gap ₹{round(gap_amount,2)} × 0.5) = ₹{base_revised}",
        'gap_pct':        round(gap_pct, 2),
        'gap_amount':     round(gap_amount, 2),
        'ema50':          ema50,
        'ema200':         ema200,
        'bb_upper':       bb_upper,
        'adx':            adx,
        'rsi':            rsi,
        'macd_bullish':   (macd is not None and macd_signal is not None and macd > macd_signal),
        'bull_pct':       bull_pct,
        'adjustments':    [],
    }

    revised = base_revised

    # EMA200 resistance — cap revised target below EMA200 if it's nearby
    if ema200 and revised > ema200 * 1.02:
        justification['adjustments'].append(
            f"EMA200 at ₹{ema200} acts as strong resistance — target capped near it"
        )
        revised = min(revised, round(ema200 * 1.015, 2))

    # BB Upper — if revised is way above upper band, flag as aggressive
    if bb_upper:
        if revised > bb_upper * 1.05:
            justification['adjustments'].append(
                f"Revised target ₹{revised} exceeds BB Upper ₹{bb_upper} by >5% — "
                f"moderating to BB Upper level"
            )
            revised = min(revised, round(bb_upper * 1.02, 2))
        else:
            justification['adjustments'].append(
                f"BB Upper ₹{bb_upper} — target within acceptable range"
            )

    # Strong ADX — allow more upside
    if adx and adx > 30:
        bonus  = round(open_price * 0.01, 2)   # extra 1% for strong trend
        revised = round(revised + bonus, 2)
        justification['adjustments'].append(
            f"ADX {adx} > 30 (strong trend) — adding ₹{bonus} bonus to target"
        )
    elif adx and adx < 20:
        justification['adjustments'].append(
            f"ADX {adx} < 20 (weak trend) — keeping target conservative"
        )

    # RSI — if already overbought, note the risk
    if rsi and rsi > 70:
        justification['adjustments'].append(
            f"RSI {rsi} is overbought — upside may be limited; monitor closely"
        )
    elif rsi and rsi < 60:
        justification['adjustments'].append(
            f"RSI {rsi} has room to run — target achievable"
        )

    # Overall bull strength
    if bull_pct >= 65:
        justification['adjustments'].append(
            f"Overall technical score {bull_pct}% bullish — supports revised target"
        )
    elif bull_pct < 50:
        justification['adjustments'].append(
            f"Technical score only {bull_pct}% bullish — revised target is conservative"
        )

    justification['final_revised_target'] = revised
    return revised, justification


def _estimate_days_to_target(ticker_yf, open_price, target_price):
    """Estimate days to reach target based on avg daily move (last 20 sessions)."""
    try:
        df = get_historical_data(ticker_yf, days=40)
        if df is None or len(df) < 5:
            return None
        closes      = df['Close'].squeeze()
        daily_moves = closes.pct_change().abs().dropna().tail(20)
        avg_daily   = float(daily_moves.mean())
        if avg_daily <= 0:
            return None
        upside_needed = abs(target_price - open_price) / open_price
        est_days      = round(upside_needed / avg_daily)
        return max(1, est_days)
    except:
        return None


def _get_support_level(ticker_yf, current_price):
    """Find nearest support level for re-entry suggestion in cancel email."""
    try:
        indicators = calculate_indicators(ticker_yf)
        bb_lower   = indicators.get('bb_lower')
        ema50      = indicators.get('ema50')

        candidates = []
        if bb_lower and bb_lower < current_price:
            candidates.append(('BB Lower Band', bb_lower))
        if ema50 and ema50 < current_price:
            candidates.append(('50-day EMA', ema50))

        if candidates:
            # Return the nearest support above zero
            candidates.sort(key=lambda x: x[1], reverse=True)
            return candidates[0]
        return None
    except:
        return None


# ── Email Builders ─────────────────────────────────────────────────────────────

def _send_buy_executed_email(pending, open_price, gap_pct, buy_decision,
                              revised_target, justification, days_to_target,
                              indicators):
    ticker       = pending['Ticker']
    stock_name   = pending.get('Stock Name', ticker)
    qty          = int(pending['Qty'])
    scout_price  = float(pending['Scout Price'])
    orig_target  = float(pending['Original Target'])
    sector       = pending.get('Sector', 'N/A')
    cap_cat      = pending.get('Cap Category', 'N/A')
    investment   = round(open_price * qty, 2)
    upside_new   = round(((revised_target - open_price) / open_price) * 100, 2) \
        if revised_target else 0
    today_str    = date.today().strftime('%d %B %Y')

    gap_color    = '#2E7D32' if gap_pct < 0 else ('#E65100' if gap_pct > 2 else '#1565C0')
    gap_label    = ('Gap Down 🔽' if gap_pct < -0.5
                    else 'Gap Up 🔼' if gap_pct > 0.5
                    else 'Flat Open ➡️')

    # Target revision section
    if revised_target and abs(revised_target - orig_target) > 0.5:
        target_was_revised = True
        revision_reason_html = _build_revision_html(justification, orig_target,
                                                     revised_target, open_price)
    else:
        target_was_revised   = False
        revision_reason_html = ''

    # Indicator summary
    ind_html = _build_indicator_html(indicators)

    days_html = (f"<tr style='background:#f8f9fa;'><td style='padding:10px 14px;"
                 f"font-weight:bold;'>⏱️ Est. Days to Target</td>"
                 f"<td style='padding:10px 14px;font-size:16px;'>"
                 f"~{days_to_target} trading days</td></tr>"
                 if days_to_target else '')

    subject = (f"✅ BUY EXECUTED: {ticker} | {qty} shares @ ₹{open_price:,.2f} | "
               f"Target ₹{revised_target:,.2f}")

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:680px;margin:auto;
             background:#f0f2f5;padding:20px;">

  <!-- Header -->
  <div style="background:linear-gradient(135deg,#1B5E20,#2E7D32);color:white;
              padding:28px;border-radius:14px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">✅ Buy Executed</h1>
    <p style="margin:6px 0 0;opacity:0.9;font-size:15px;">
      {stock_name} ({ticker}) &nbsp;|&nbsp; {cap_cat} &nbsp;|&nbsp; {sector}
    </p>
    <p style="margin:4px 0 0;opacity:0.75;font-size:13px;">{today_str}</p>
  </div>

  <!-- Core trade details -->
  <div style="background:white;border-radius:12px;padding:20px;margin-bottom:16px;
              box-shadow:0 2px 8px rgba(0,0,0,0.07);">
    <h3 style="margin:0 0 14px;color:#1B5E20;">📋 Trade Summary</h3>
    <table style="width:100%;border-collapse:collapse;">
      <tr><td style="padding:10px 14px;font-weight:bold;width:45%;">Ticker</td>
          <td style="padding:10px 14px;font-size:18px;font-weight:bold;">{ticker}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Scout Price (Yesterday)</td>
          <td style="padding:10px 14px;">₹{scout_price:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Actual Buy Price (Open)</td>
          <td style="padding:10px 14px;font-size:18px;font-weight:bold;color:#1565C0;">
            ₹{open_price:,.2f}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Gap at Open</td>
          <td style="padding:10px 14px;font-size:16px;font-weight:bold;color:{gap_color};">
            {gap_pct:+.2f}% &nbsp; {gap_label}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Quantity</td>
          <td style="padding:10px 14px;font-size:18px;font-weight:bold;">{qty} shares</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:10px 14px;font-weight:bold;">Total Investment</td>
          <td style="padding:10px 14px;font-size:20px;font-weight:bold;color:#1B5E20;">
            ₹{investment:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Original Target</td>
          <td style="padding:10px 14px;">₹{orig_target:,.2f}
            {'&nbsp;<span style="background:#FFF9C4;padding:2px 8px;border-radius:4px;'
             'font-size:12px;">REVISED ↗</span>' if target_was_revised else ''}</td></tr>
      <tr style="background:#E8F5E9;">
          <td style="padding:10px 14px;font-weight:bold;color:#1B5E20;">🎯 Final Target</td>
          <td style="padding:10px 14px;font-size:20px;font-weight:bold;color:#1B5E20;">
            ₹{revised_target:,.2f} &nbsp;
            <span style="font-size:14px;color:#2E7D32;">(+{upside_new:.1f}% upside)</span>
          </td></tr>
      {days_html}
    </table>
  </div>

  <!-- Target revision justification (only shown if revised) -->
  {revision_reason_html}

  <!-- Live technical indicators at open -->
  <div style="background:white;border-radius:12px;padding:20px;margin-bottom:16px;
              box-shadow:0 2px 8px rgba(0,0,0,0.07);">
    <h3 style="margin:0 0 14px;color:#1565C0;">📊 Live Technical Indicators at Open</h3>
    {ind_html}
  </div>

  <!-- Footer -->
  <div style="background:#E8F5E9;border-radius:10px;padding:14px;
              text-align:center;font-size:13px;color:#2E7D32;">
    Stock added to Holdings sheet. Monitor daily analysis emails for exit signals.
  </div>

</body></html>"""

    send_report_email(subject, html)
    print(f"[pending_buys] ✅ Buy Executed email sent for {ticker}")


def _build_revision_html(justification, orig_target, revised_target, open_price):
    """Build the detailed target revision justification block."""
    adj_items = ''.join(
        f"<li style='margin-bottom:6px;'>{a}</li>"
        for a in justification.get('adjustments', [])
    )
    ema50    = justification.get('ema50')
    ema200   = justification.get('ema200')
    bb_upper = justification.get('bb_upper')
    adx      = justification.get('adx')
    rsi      = justification.get('rsi')
    gap_pct  = justification.get('gap_pct', 0)

    tech_rows = ''
    for label, val, note in [
        ('EMA 50',      ema50,    'Short-term trend anchor'),
        ('EMA 200',     ema200,   'Long-term resistance/support'),
        ('BB Upper',    bb_upper, 'Volatility ceiling'),
        ('ADX',         adx,      'Trend strength'),
        ('RSI',         rsi,      'Momentum'),
    ]:
        if val is not None:
            tech_rows += (f"<tr><td style='padding:7px 12px;font-weight:bold;'>{label}</td>"
                          f"<td style='padding:7px 12px;'>₹{val:,.2f}" if label not in ('ADX','RSI')
                          else f"<tr><td style='padding:7px 12px;font-weight:bold;'>{label}</td>"
                               f"<td style='padding:7px 12px;'>{val}")
            tech_rows += f"</td><td style='padding:7px 12px;color:#666;font-size:12px;'>{note}</td></tr>"

    return f"""
  <div style="background:white;border-radius:12px;padding:20px;margin-bottom:16px;
              border-left:5px solid #F57F17;box-shadow:0 2px 8px rgba(0,0,0,0.07);">
    <h3 style="margin:0 0 12px;color:#F57F17;">🔄 Target Revised — Full Justification</h3>

    <div style="background:#FFFDE7;border-radius:8px;padding:14px;margin-bottom:14px;">
      <strong>Revision Formula:</strong><br>
      <code style="font-size:14px;">{justification.get('formula','')}</code>
    </div>

    <p style="margin:0 0 8px;font-weight:bold;">Why the gap justified a target revision:</p>
    <ul style="margin:0 0 14px;padding-left:18px;line-height:1.7;">
      <li>The stock gapped <strong>{gap_pct:+.2f}%</strong> from scout price ₹{open_price - justification.get('gap_amount',0):,.2f}
          to open ₹{open_price:,.2f} — this represents real price discovery, not just noise.</li>
      <li>Gap amount of ₹{justification.get('gap_amount',0):,.2f} × 0.5 = ₹{round(justification.get('gap_amount',0)*0.5,2)}
          added to original target ₹{orig_target:,.2f} → base revised target ₹{round(orig_target + justification.get('gap_amount',0)*0.5,2):,.2f}.</li>
      <li>50% of the gap (not 100%) is used to avoid over-projecting momentum that may fade.</li>
    </ul>

    <p style="margin:0 0 8px;font-weight:bold;">Technical factor adjustments:</p>
    <ul style="margin:0 0 14px;padding-left:18px;line-height:1.7;">{adj_items}</ul>

    <p style="margin:0 0 8px;font-weight:bold;">Key levels at open:</p>
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      <tr style="background:#f8f9fa;"><th style="padding:7px 12px;text-align:left;">Indicator</th>
          <th style="padding:7px 12px;text-align:left;">Value</th>
          <th style="padding:7px 12px;text-align:left;">Significance</th></tr>
      {tech_rows}
    </table>

    <div style="background:#E8F5E9;border-radius:6px;padding:10px;margin-top:12px;">
      <strong>Final Revised Target: ₹{revised_target:,.2f}</strong>
      (was ₹{orig_target:,.2f} → revised ₹{revised_target - orig_target:+,.2f})
    </div>
  </div>"""


def _build_indicator_html(indicators):
    """Compact indicator table for the email."""
    rows = ''
    pairs = [
        ('RSI',           indicators.get('rsi'),        '< 35 oversold · > 65 overbought'),
        ('MACD',          indicators.get('macd'),        '> signal = bullish'),
        ('MACD Signal',   indicators.get('macd_signal'), ''),
        ('BB Upper',      indicators.get('bb_upper'),    'Resistance ceiling'),
        ('BB Mid',        indicators.get('bb_mid'),      'Mean reversion level'),
        ('BB Lower',      indicators.get('bb_lower'),    'Support floor'),
        ('ADX',           indicators.get('adx'),         '> 25 = strong trend'),
        ('Stochastic K',  indicators.get('stoch_k'),     '< 20 oversold · > 80 overbought'),
        ('EMA 50',        indicators.get('ema50'),       'Short-term trend'),
        ('EMA 200',       indicators.get('ema200'),      'Long-term trend'),
    ]
    bg = False
    for label, val, note in pairs:
        if val is None:
            continue
        bg_style = "background:#f8f9fa;" if bg else ""
        display  = f"₹{val:,.2f}" if label not in ('RSI', 'ADX', 'Stochastic K', 'MACD', 'MACD Signal') \
                   else f"{val}"
        rows += (f"<tr style='{bg_style}'>"
                 f"<td style='padding:7px 12px;font-weight:bold;'>{label}</td>"
                 f"<td style='padding:7px 12px;'>{display}</td>"
                 f"<td style='padding:7px 12px;color:#666;font-size:12px;'>{note}</td></tr>")
        bg = not bg

    overall = indicators.get('bull_pct', 50)
    color   = '#2E7D32' if overall >= 65 else ('#C62828' if overall <= 35 else '#E65100')
    signal  = indicators.get('technical_signal', 'NEUTRAL')

    return f"""
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      <tr style="background:#1565C0;color:white;">
        <th style="padding:8px 12px;text-align:left;">Indicator</th>
        <th style="padding:8px 12px;text-align:left;">Value</th>
        <th style="padding:8px 12px;text-align:left;">Note</th>
      </tr>
      {rows}
    </table>
    <div style="margin-top:12px;background:#f8f9fa;border-radius:6px;padding:10px;
                text-align:center;">
      Overall Signal: <strong style="color:{color};">{signal}</strong> &nbsp;|&nbsp;
      Bull Score: <strong style="color:{color};">{overall}%</strong>
    </div>"""


def _send_buy_cancelled_email(pending, open_price, gap_pct, cancel_reason,
                               bad_news_headlines, bad_keywords, support_info):
    ticker      = pending['Ticker']
    stock_name  = pending.get('Stock Name', ticker)
    qty         = int(pending['Qty'])
    scout_price = float(pending['Scout Price'])
    orig_target = float(pending['Original Target'])
    sector      = pending.get('Sector', 'N/A')
    threshold   = pending.get('_threshold', DEFAULT_THRESHOLD)

    today_str   = date.today().strftime('%d %B %Y')

    # Support / re-entry level
    if support_info:
        support_html = f"""
    <div style="background:#E3F2FD;border-radius:8px;padding:14px;margin-top:14px;">
      <strong>📌 Re-entry Watch Level:</strong><br>
      Consider re-entering around <strong>₹{support_info[1]:,.2f}</strong>
      ({support_info[0]}) if fundamentals remain intact.
      Wait for the next scout email recommendation for a fresh analysis.
    </div>"""
    else:
        support_html = ""

    # News block
    if bad_news_headlines:
        news_items = ''.join(f"<li style='margin-bottom:6px;'>{h}</li>"
                             for h in bad_news_headlines)
        news_html  = f"""
    <div style="background:#FFEBEE;border-radius:8px;padding:14px;margin-bottom:14px;">
      <p style="margin:0 0 8px;font-weight:bold;color:#C62828;">
        🚨 Negative News Detected (keywords: {', '.join(bad_keywords)}):
      </p>
      <ul style="margin:0;padding-left:18px;line-height:1.7;">{news_items}</ul>
    </div>"""
    else:
        news_html = ""

    # Reason explanation
    reason_map = {
        'gap_down_bad_news':   ('Gap Down + Bad News',
                                f"Stock gapped down {gap_pct:.2f}% AND negative news was detected. "
                                f"Entering on bad news gap-down is high risk."),
        'gap_too_extreme_down':('Gap Down Too Extreme',
                                f"Stock gapped down {gap_pct:.2f}%, exceeding the "
                                f"±{threshold}% threshold for {sector} sector. "
                                f"Something may have changed overnight — buying blind into a "
                                f"severe gap-down without clear news is not advisable."),
        'target_nearly_hit':   ('Target Nearly Hit at Open',
                                f"Stock gapped up {gap_pct:.2f}%, which is ≥80% of the original "
                                f"upside target. The risk/reward ratio is now destroyed — "
                                f"you would be entering a trade where most of the move has already happened."),
    }
    r_title, r_body = reason_map.get(cancel_reason, ('Unknown', cancel_reason))

    subject = f"❌ BUY CANCELLED: {ticker} | {r_title} | Gap {gap_pct:+.2f}%"

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:680px;margin:auto;
             background:#f0f2f5;padding:20px;">

  <!-- Header -->
  <div style="background:linear-gradient(135deg,#B71C1C,#C62828);color:white;
              padding:28px;border-radius:14px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">❌ Buy Cancelled</h1>
    <p style="margin:6px 0 0;opacity:0.9;font-size:15px;">
      {stock_name} ({ticker}) &nbsp;|&nbsp; {sector}
    </p>
    <p style="margin:4px 0 0;opacity:0.75;font-size:13px;">{today_str}</p>
  </div>

  <!-- Why cancelled -->
  <div style="background:white;border-radius:12px;padding:20px;margin-bottom:16px;
              border-left:5px solid #C62828;box-shadow:0 2px 8px rgba(0,0,0,0.07);">
    <h3 style="margin:0 0 12px;color:#C62828;">Why This Buy Was Cancelled</h3>
    <div style="background:#FFEBEE;border-radius:8px;padding:14px;margin-bottom:14px;">
      <strong style="color:#C62828;">{r_title}</strong><br>
      <span style="line-height:1.7;">{r_body}</span>
    </div>
    {news_html}

    <table style="width:100%;border-collapse:collapse;font-size:14px;margin-top:10px;">
      <tr><td style="padding:9px 14px;font-weight:bold;width:45%;">Scout Price</td>
          <td style="padding:9px 14px;">₹{scout_price:,.2f}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:9px 14px;font-weight:bold;">Open Price Today</td>
          <td style="padding:9px 14px;font-size:16px;font-weight:bold;">
            ₹{open_price:,.2f}</td></tr>
      <tr><td style="padding:9px 14px;font-weight:bold;">Gap at Open</td>
          <td style="padding:9px 14px;font-size:16px;font-weight:bold;
              color:{'#C62828' if gap_pct < 0 else '#E65100'};">
            {gap_pct:+.2f}%</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:9px 14px;font-weight:bold;">Sector Threshold</td>
          <td style="padding:9px 14px;">±{threshold}%</td></tr>
      <tr><td style="padding:9px 14px;font-weight:bold;">Original Target</td>
          <td style="padding:9px 14px;">₹{orig_target:,.2f}</td></tr>
      <tr style="background:#f8f9fa;">
          <td style="padding:9px 14px;font-weight:bold;">Qty Requested</td>
          <td style="padding:9px 14px;">{qty} shares</td></tr>
    </table>
    {support_html}
  </div>

  <!-- Footer -->
  <div style="background:#FFF3E0;border-radius:10px;padding:14px;
              text-align:center;font-size:13px;color:#E65100;">
    No action needed. This stock was NOT added to your Holdings.
    Watch tomorrow's scout email for a fresh recommendation.
  </div>

</body></html>"""

    send_report_email(subject, html)
    print(f"[pending_buys] ❌ Buy Cancelled email sent for {ticker} — {cancel_reason}")


# ── Main Execution Function ────────────────────────────────────────────────────

def execute_pending_buys():
    """
    Called at 9:20 AM IST from main.py --morning.
    Processes all PENDING rows in PendingBuys sheet.
    """
    print("[pending_buys] === Executing Pending Buys ===")
    ws      = _get_pending_buys_ws()
    pending_list = get_pending_buys()

    if not pending_list:
        print("[pending_buys] No pending buys to process.")
        return

    for pending in pending_list:
        ticker     = str(pending.get('Ticker', '')).upper()
        if not ticker:
            continue

        ticker_yf  = ticker + NSE_SUFFIX
        scout_price = float(pending.get('Scout Price', 0) or 0)
        orig_target = float(pending.get('Original Target', 0) or 0)
        qty         = int(pending.get('Qty', 0) or 0)
        sector      = str(pending.get('Sector', 'Unknown'))
        cap_cat     = str(pending.get('Cap Category', ''))

        if scout_price <= 0 or qty <= 0:
            print(f"[pending_buys] {ticker}: Invalid scout_price/qty — skipping.")
            _update_pending_row(ws, ticker, {
                'Status': 'CANCELLED', 'Reason': 'Invalid scout price or qty'
            })
            continue

        print(f"\n[pending_buys] Processing: {ticker} x{qty} | Scout: ₹{scout_price}")

        # Step 1 — Fetch actual open price
        open_price = _get_open_price(ticker_yf)
        if open_price is None:
            print(f"[pending_buys] {ticker}: Could not fetch open price — skipping.")
            _update_pending_row(ws, ticker, {
                'Status': 'CANCELLED', 'Reason': 'Could not fetch opening price'
            })
            continue

        # Step 2 — Calculate gap
        prev_close = _get_previous_close(ticker_yf) or scout_price
        gap_pct    = round(((open_price - prev_close) / prev_close) * 100, 2)
        upside_pct = round(((orig_target - scout_price) / scout_price) * 100, 2) \
            if scout_price else 0

        threshold  = _get_threshold(sector, cap_cat)
        pending['_threshold'] = threshold   # pass to email builder

        print(f"[pending_buys] {ticker}: Open ₹{open_price} | Prev Close ₹{prev_close} | "
              f"Gap {gap_pct:+.2f}% | Threshold ±{threshold}% | Upside {upside_pct:.1f}%")

        # Step 3 — Fetch live indicators (fresh data at open)
        print(f"[pending_buys] {ticker}: Fetching live indicators...")
        indicators = calculate_indicators(ticker_yf)
        time.sleep(1)   # gentle rate limit

        # Step 4 — Apply gap decision tree
        cancel_reason       = None
        bad_news_headlines  = []
        bad_keywords_found  = []
        revised_target      = orig_target   # default: keep original
        justification       = {}

        # ── CASE A: Gap down beyond threshold ──
        if gap_pct < -threshold:
            has_bad_news, headlines, bad_kw = _fetch_overnight_news(ticker)
            if has_bad_news:
                cancel_reason      = 'gap_down_bad_news'
                bad_news_headlines = headlines
                bad_keywords_found = bad_kw
            else:
                # No bad news — gap down is actually a better entry
                # Keep original absolute target → more upside now
                revised_target = orig_target
                justification  = {
                    'formula':      f"Gap down {gap_pct:.2f}% with no bad news → "
                                    f"keeping original absolute target ₹{orig_target} "
                                    f"for increased upside",
                    'gap_pct':      gap_pct,
                    'gap_amount':   round(open_price - scout_price, 2),
                    'adjustments':  [
                        f"No negative news found — gap likely due to market/sector weakness",
                        f"Lower entry ₹{open_price} improves upside from "
                        f"{upside_pct:.1f}% to "
                        f"{round(((orig_target-open_price)/open_price)*100,1):.1f}%",
                    ],
                    'final_revised_target': orig_target,
                }
                print(f"[pending_buys] {ticker}: Gap down but no bad news → BUY at ₹{open_price}")

        # ── CASE B: Gap too extreme downward even with no news check
        #    (already handled above — gap_pct < -threshold with bad news = cancel)

        # ── CASE C: Normal range → buy, keep original target ──
        elif -threshold <= gap_pct <= threshold:
            revised_target = orig_target
            justification  = {
                'formula':      f"Gap {gap_pct:+.2f}% within normal ±{threshold}% range — "
                                f"original target maintained",
                'gap_pct':      gap_pct,
                'gap_amount':   round(open_price - scout_price, 2),
                'adjustments':  [f"Gap within threshold — no revision needed"],
                'final_revised_target': orig_target,
            }
            print(f"[pending_buys] {ticker}: Normal gap → BUY at ₹{open_price}, target ₹{orig_target}")

        # ── CASE D: Gap up moderate → buy, revise target upward ──
        elif gap_pct > threshold:
            extreme_threshold = upside_pct * 0.8
            if gap_pct >= extreme_threshold:
                # Gap up ≥ 80% of original upside → CANCEL
                cancel_reason = 'target_nearly_hit'
                print(f"[pending_buys] {ticker}: Extreme gap up {gap_pct:.2f}% ≥ "
                      f"{extreme_threshold:.2f}% (80% of upside) → CANCEL")
            else:
                # Moderate gap up → revise target upward
                revised_target, justification = _compute_revised_target(
                    ticker_yf, open_price, orig_target,
                    scout_price, gap_pct, indicators
                )
                print(f"[pending_buys] {ticker}: Moderate gap up → BUY at ₹{open_price}, "
                      f"revised target ₹{revised_target}")

        # Step 5 — Execute decision
        if cancel_reason:
            # Find support for re-entry suggestion
            support_info = _get_support_level(ticker_yf, open_price)

            _update_pending_row(ws, ticker, {
                'Status':         'CANCELLED',
                'Gap %':          gap_pct,
                'Actual Buy Price': open_price,
                'Revised Target': '',
                'Reason':         cancel_reason,
            })

            _send_buy_cancelled_email(
                pending, open_price, gap_pct, cancel_reason,
                bad_news_headlines, bad_keywords_found, support_info
            )

        else:
            # BUY — add to Holdings
            stock    = yf.Ticker(ticker_yf)
            info     = stock.info
            name     = info.get('longName', pending.get('Stock Name', ticker))

            add_stock_to_holdings(
                ticker       = ticker,
                stock_name   = name,
                industry     = sector,
                buying_price = open_price,
                buying_date  = date.today().strftime('%Y-%m-%d'),
                qty          = qty,
                target_price = revised_target,
                cap_category = cap_cat,
                sector       = sector,
            )

            _update_pending_row(ws, ticker, {
                'Status':           'EXECUTED',
                'Gap %':            gap_pct,
                'Actual Buy Price': open_price,
                'Revised Target':   revised_target,
                'Reason':           f"Gap {gap_pct:+.2f}% — executed at open",
            })

            days_to_target = _estimate_days_to_target(ticker_yf, open_price, revised_target)

            _send_buy_executed_email(
                pending, open_price, gap_pct,
                'BUY', revised_target, justification,
                days_to_target, indicators
            )

        print(f"[pending_buys] {ticker}: Done.\n")

    print("[pending_buys] === Pending Buys Processing Complete ===")


if __name__ == "__main__":
    execute_pending_buys()
