"""
decision_engine.py
Core decision engine with:
  - FLAW 2.1: Minimum 2 consecutive BEARISH verdicts before auto-sell triggers
  - FLAW 2.1: SENTIMENT FLIP WARNING if verdict flips single day with no major move
  - FLAW 3.1: Three stock categories (Large/Mid/Small Cap) with different thresholds
              Auto-classified from yfinance marketCap — no manual column needed
  - FLAW 3.4: Decisions flagged as pending_sell=True — actual execution next morning
"""

import yfinance as yf
from config import NSE_SUFFIX

# ── FLAW 3.1: Cap category thresholds ─────────────────────────────────────────
# marketCap in INR (rupees)
LARGE_CAP_THRESHOLD = 20_000_00_00_000   # Rs. 20,000 Cr
MID_CAP_THRESHOLD   =  5_000_00_00_000   # Rs.  5,000 Cr

CAP_THRESHOLDS = {
    'Large Cap': {'stop_loss': -15, 'profit_target': 40},
    'Mid Cap':   {'stop_loss': -20, 'profit_target': 60},
    'Small Cap': {'stop_loss': -25, 'profit_target': 80},
}


def get_cap_category(market_cap):
    """Classify stock as Large/Mid/Small Cap from marketCap value"""
    if market_cap is None:
        return 'Mid Cap'  # safe default if unknown
    if market_cap >= LARGE_CAP_THRESHOLD:
        return 'Large Cap'
    elif market_cap >= MID_CAP_THRESHOLD:
        return 'Mid Cap'
    else:
        return 'Small Cap'


def fetch_market_cap(ticker_yf):
    """Fetch marketCap from yfinance"""
    try:
        return yf.Ticker(ticker_yf).info.get('marketCap')
    except:
        return None


# ── FLAW 2.1: Consecutive BEARISH count ───────────────────────────────────────

def count_consecutive_bearish(sentiment_history):
    """
    Count how many consecutive BEARISH verdicts appear at the end of history.
    Example: [BULLISH, NEUTRAL, BEARISH, BEARISH] → returns 2
    """
    if not sentiment_history:
        return 0
    count = 0
    for entry in reversed(sentiment_history):
        if str(entry.get('verdict', '')).upper() == 'BEARISH':
            count += 1
        else:
            break
    return count


# ── CORE DECISION LOGIC ────────────────────────────────────────────────────────

def make_decision(stock):
    """
    stock dict must contain:
      - growth_pct, technical_signal, overall_sentiment
      - sentiment_flip: 'YES'/'NO' (from Groq)
      - sentiment_history: list of past verdicts [{date, verdict}]
      - live_price, buying_price, target_price
      - ticker_yf: for market cap lookup
      - market_cap: optional (fetched if missing)
    """

    growth         = stock.get('growth_pct')
    live_price     = stock.get('live_price')
    target_price   = stock.get('target_price')
    tech_signal    = stock.get('technical_signal', 'NEUTRAL').upper()
    ai_sentiment   = stock.get('overall_sentiment', 'NEUTRAL').upper()
    sentiment_flip = stock.get('sentiment_flip', 'NO').upper()
    history        = stock.get('sentiment_history', [])

    # FLAW 3.1: Determine cap category and its thresholds
    market_cap   = stock.get('market_cap') or fetch_market_cap(stock.get('ticker_yf', ''))
    cap_category = get_cap_category(market_cap)
    thresholds   = CAP_THRESHOLDS[cap_category]
    stop_loss_threshold    = thresholds['stop_loss']
    profit_target_threshold = thresholds['profit_target']

    # FLAW 2.1: Count consecutive BEARISH days
    consecutive_bearish = count_consecutive_bearish(history)

    # Combined signal
    if tech_signal == 'BULLISH' or ai_sentiment == 'BULLISH':
        combined = 'BULLISH'
    elif tech_signal == 'BEARISH' and ai_sentiment == 'BEARISH':
        combined = 'BEARISH'
    else:
        combined = 'NEUTRAL'

    decision     = 'HOLD'
    urgency      = 'LOW'
    reason       = ''
    action_detail = ''
    pending_sell  = False

    if growth is None:
        decision      = 'HOLD'
        urgency       = 'MEDIUM'
        reason        = 'Could not fetch live price. Manual check required.'

    # ── Target price hit — always sell regardless of category ─────────────────
    elif target_price and live_price and live_price >= target_price:
        decision      = 'SELL'
        urgency       = 'HIGH'
        pending_sell  = True
        reason        = f'TARGET PRICE HIT! Target was Rs.{target_price:,.2f}, current Rs.{live_price:,.2f}'
        action_detail = 'AI target achieved — booking profit as planned!'

    # ── Stop loss zone ─────────────────────────────────────────────────────────
    elif growth <= stop_loss_threshold:
        if combined == 'BEARISH':
            # FLAW 2.1: Require 2 consecutive BEARISH days before auto-sell
            if consecutive_bearish >= 2:
                decision      = 'SELL'
                urgency       = 'HIGH'
                pending_sell  = True
                reason        = (f'Loss of {abs(growth):.1f}% exceeds {cap_category} stop loss '
                                 f'({abs(stop_loss_threshold)}%) AND BEARISH for {consecutive_bearish} consecutive days.')
                action_detail = 'SELL confirmed — 2+ consecutive BEARISH verdicts with significant loss.'
            else:
                # First BEARISH day — warn but do NOT sell yet
                decision      = 'HOLD'
                urgency       = 'HIGH'
                reason        = (f'Loss of {abs(growth):.1f}% exceeds {cap_category} stop loss '
                                 f'({abs(stop_loss_threshold)}%) — BEARISH Day {consecutive_bearish}/2. '
                                 f'Will SELL if BEARISH again tomorrow.')
                action_detail = 'Waiting for 2nd consecutive BEARISH day to confirm sell.'
        elif combined == 'BULLISH':
            decision      = 'HOLD'
            urgency       = 'MEDIUM'
            reason        = f'Loss of {abs(growth):.1f}% exceeds {cap_category} threshold BUT sentiment BULLISH. Holding for recovery.'
            action_detail = 'Monitor closely. Set stop-loss alert at another -5%.'
        else:
            decision      = 'HOLD'
            urgency       = 'MEDIUM'
            reason        = f'Loss of {abs(growth):.1f}% exceeds {cap_category} threshold. Mixed signals — monitor carefully.'
            action_detail = 'Watch for next 2-3 trading sessions before deciding.'

    # ── Profit target zone ─────────────────────────────────────────────────────
    elif growth >= profit_target_threshold:
        if combined == 'BULLISH':
            decision      = 'HOLD'
            urgency       = 'LOW'
            reason        = f'Profit of {growth:.1f}% achieved ({cap_category} target: {profit_target_threshold}%). BULLISH — holding for peak.'
            action_detail = 'Trail stop-loss to protect at least 50% gains.'
        else:
            decision      = 'SELL'
            urgency       = 'MEDIUM'
            pending_sell  = True
            reason        = f'Profit of {growth:.1f}% achieved ({cap_category} target: {profit_target_threshold}%). No strong uptrend — booking profit.'
            action_detail = 'Consider partial sell (50%) to lock in gains.'

    # ── Normal range ───────────────────────────────────────────────────────────
    else:
        # FLAW 2.1: Sentiment flip warning — don't act on a single-day flip
        if sentiment_flip == 'YES' and combined == 'BEARISH':
            decision      = 'HOLD'
            urgency       = 'MEDIUM'
            reason        = f'SENTIMENT FLIP WARNING: Switched to BEARISH today after recent bullish trend. {abs(growth):.1f}% from buy. Waiting for confirmation.'
            action_detail = 'Single-day flip — monitoring. Will act if BEARISH again tomorrow.'
        elif combined == 'BEARISH' and growth < -10:
            decision      = 'HOLD'
            urgency       = 'MEDIUM'
            reason        = f'Loss of {abs(growth):.1f}% — below {cap_category} stop loss but bearish signals emerging. Watch carefully.'
            action_detail = f'Reassess if loss deepens to {abs(stop_loss_threshold)}%.'
        elif combined == 'BULLISH' and growth > 30:
            decision      = 'HOLD'
            urgency       = 'LOW'
            reason        = f'Good profit of {growth:.1f}% with bullish outlook. Let it run.'
            action_detail = 'Continue monitoring. No action needed today.'
        else:
            decision      = 'HOLD'
            urgency       = 'LOW'
            reason        = f'Return of {growth:.1f}%. Stock within normal parameters ({cap_category}).'
            action_detail = 'No action needed.'

    return {
        'decision':             decision,
        'urgency':              urgency,
        'reason':               reason,
        'action_detail':        action_detail,
        'combined_signal':      combined,
        'cap_category':         cap_category,
        'market_cap':           market_cap,
        'stop_loss_threshold':  stop_loss_threshold,
        'profit_threshold':     profit_target_threshold,
        'consecutive_bearish':  consecutive_bearish,
        'pending_sell':         pending_sell,
    }


def process_all_holdings(enriched_holdings):
    """Run decision engine on all holdings"""
    results = []
    for stock in enriched_holdings:
        dec    = make_decision(stock)
        result = {**stock, **dec}
        results.append(result)

        if dec['decision'] == 'SELL':
            icon = 'PENDING SELL'
        elif dec['urgency'] == 'HIGH':
            icon = 'HIGH ALERT'
        elif dec['urgency'] == 'MEDIUM':
            icon = 'WATCH'
        else:
            icon = 'HOLD'

        print(f"[decision] {stock['ticker']} ({dec['cap_category']}) -> {icon} | {dec['reason'][:70]}")
    return results