"""
decision_engine.py
Core decision engine: applies your trading rules to give BUY / SELL / HOLD decisions.

Rules (as defined by user):
  SELL if loss >= 20% AND bearish  → SELL IMMEDIATELY
  SELL if loss >= 20% AND bullish  → HOLD (wait for recovery)
  SELL if profit >= 70% AND no uptrend → SELL (take profit)
  SELL if profit >= 70% AND bullish    → HOLD (ride to peak)
  Otherwise                            → HOLD / MONITOR
"""

from config import SELL_LOSS_THRESHOLD, SELL_PROFIT_THRESHOLD


def make_decision(stock):
    """
    stock dict must contain:
      - growth_pct: float (e.g. -25.5 or 80.0)
      - technical_signal: 'BULLISH' / 'BEARISH' / 'NEUTRAL'
      - overall_sentiment: 'BULLISH' / 'BEARISH' / 'NEUTRAL'
      - stock_name, ticker, live_price, buying_price, total_profit
    """

    growth = stock.get('growth_pct')
    tech_signal = stock.get('technical_signal', 'NEUTRAL').upper()
    ai_sentiment = stock.get('overall_sentiment', 'NEUTRAL').upper()

    # Combined signal: if either is BULLISH, lean bullish
    if tech_signal == 'BULLISH' or ai_sentiment == 'BULLISH':
        combined = 'BULLISH'
    elif tech_signal == 'BEARISH' and ai_sentiment == 'BEARISH':
        combined = 'BEARISH'
    else:
        combined = 'NEUTRAL'

    decision = 'HOLD'
    urgency = 'LOW'
    reason = ''
    action_detail = ''

    if growth is None:
        decision = 'HOLD'
        reason = 'Could not fetch live price. Manual check required.'
        urgency = 'MEDIUM'

    elif growth <= SELL_LOSS_THRESHOLD:
        # LOSS >= 20%
        if combined == 'BEARISH':
            decision = 'SELL'
            urgency = 'HIGH'
            reason = f'Loss of {abs(growth):.1f}% exceeds {abs(SELL_LOSS_THRESHOLD)}% threshold AND technical/sentiment is BEARISH.'
            action_detail = 'SELL IMMEDIATELY to prevent further loss.'
        elif combined == 'BULLISH':
            decision = 'HOLD'
            urgency = 'MEDIUM'
            reason = f'Loss of {abs(growth):.1f}% exceeds threshold BUT technical/sentiment is BULLISH. Holding for possible recovery.'
            action_detail = 'Monitor closely. Set stop-loss alert at another -5%.'
        else:
            decision = 'HOLD'
            urgency = 'MEDIUM'
            reason = f'Loss of {abs(growth):.1f}% exceeds threshold. Mixed signals — monitor carefully.'
            action_detail = 'Watch for next 2-3 trading sessions before deciding.'

    elif growth >= SELL_PROFIT_THRESHOLD:
        # PROFIT >= 70%
        if combined == 'BULLISH':
            decision = 'HOLD'
            urgency = 'LOW'
            reason = f'Profit of {growth:.1f}% achieved. Sentiment is BULLISH — holding to ride further upside.'
            action_detail = 'Trail stop-loss to protect at least 50% gains.'
        else:
            decision = 'SELL'
            urgency = 'MEDIUM'
            reason = f'Profit of {growth:.1f}% achieved. No strong uptrend signal — booking profit.'
            action_detail = 'Consider partial sell (50%) to lock in gains, hold rest.'

    else:
        # Normal range
        if combined == 'BEARISH' and growth < -10:
            decision = 'HOLD'
            urgency = 'MEDIUM'
            reason = f'Loss of {abs(growth):.1f}% — below sell threshold but bearish signals emerging. Watch carefully.'
            action_detail = 'Reassess if loss deepens to {SELL_LOSS_THRESHOLD}%.'
        elif combined == 'BULLISH' and growth > 30:
            decision = 'HOLD'
            urgency = 'LOW'
            reason = f'Good profit of {growth:.1f}% with bullish outlook. Let it run.'
            action_detail = 'Continue monitoring. No action needed today.'
        else:
            decision = 'HOLD'
            urgency = 'LOW'
            reason = f'Return of {growth:.1f}%. Stock within normal parameters.'
            action_detail = 'No action needed.'

    return {
        'decision': decision,
        'urgency': urgency,
        'reason': reason,
        'action_detail': action_detail,
        'combined_signal': combined,
    }


def process_all_holdings(enriched_holdings):
    """Run decision engine on all holdings"""
    results = []
    for stock in enriched_holdings:
        dec = make_decision(stock)
        results.append({**stock, **dec})
        status_icon = '🔴 SELL' if dec['decision'] == 'SELL' else '🟡 WATCH' if dec['urgency'] == 'MEDIUM' else '🟢 HOLD'
        print(f"[decision] {stock['ticker']} → {status_icon} | {dec['reason'][:60]}")
    return results
