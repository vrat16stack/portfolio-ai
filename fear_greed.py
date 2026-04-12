"""
fear_greed.py
Fetches CNN Fear & Greed Index — free, no API key needed.
Gives overall market mood context for daily report.

Score interpretation:
  0-25  = Extreme Fear  → Good time to buy (market oversold)
  25-45 = Fear          → Cautious buying opportunity
  45-55 = Neutral       → Normal market conditions
  55-75 = Greed         → Be careful, market getting expensive
  75-100= Extreme Greed → Consider booking profits
"""

import urllib.request
import json


def get_fear_greed():
    """Fetch current Fear & Greed Index from alternative API"""
    try:
        # Using alternative Fear & Greed API
        url = "https://api.alternative.me/fng/?limit=30"
        req = urllib.request.Request(
            url,
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read())

        latest    = data['data'][0]
        prev_day  = data['data'][1]
        prev_week = data['data'][6] if len(data['data']) > 6 else data['data'][-1]
        prev_month = data['data'][29] if len(data['data']) > 29 else data['data'][-1]

        score      = int(latest['value'])
        rating     = latest['value_classification'].title()
        prev_close = int(prev_day['value'])
        prev_week_val  = int(prev_week['value'])
        prev_month_val = int(prev_month['value'])

        return {
            'score':       score,
            'rating':      rating,
            'prev_close':  prev_close,
            'prev_week':   prev_week_val,
            'prev_month':  prev_month_val,
            'color':       _get_color(score),
            'emoji':       _get_emoji(score),
            'advice':      _get_advice(score),
        }

    except Exception as e:
        print(f"[fear_greed] Error fetching index: {e}")
        return {
            'score':      50,
            'rating':     'Neutral',
            'prev_close': 50,
            'prev_week':  50,
            'prev_month': 50,
            'color':      '#FF9800',
            'emoji':      '😐',
            'advice':     'Market mood data unavailable today.',
        }


def _get_color(score):
    if score <= 25:   return '#C62828'   # Dark red
    elif score <= 45: return '#FF5722'   # Orange red
    elif score <= 55: return '#FF9800'   # Orange
    elif score <= 75: return '#7CB342'   # Light green
    else:             return '#2E7D32'   # Dark green


def _get_emoji(score):
    if score <= 25:   return '😱'
    elif score <= 45: return '😟'
    elif score <= 55: return '😐'
    elif score <= 75: return '😊'
    else:             return '🤑'


def _get_advice(score):
    if score <= 25:
        return 'Extreme Fear — Market is oversold. Historically a good time to buy quality stocks.'
    elif score <= 45:
        return 'Fear in market — Cautious buying opportunities may exist. Focus on fundamentally strong stocks.'
    elif score <= 55:
        return 'Neutral market — Normal conditions. Stick to your trading rules.'
    elif score <= 75:
        return 'Greed building — Market getting expensive. Be selective, avoid chasing rallies.'
    else:
        return 'Extreme Greed — Market overheated. Consider booking profits on stocks near targets.'


if __name__ == "__main__":
    result = get_fear_greed()
    print(f"Fear & Greed: {result['score']} — {result['rating']} {result['emoji']}")
    print(f"Advice: {result['advice']}")
