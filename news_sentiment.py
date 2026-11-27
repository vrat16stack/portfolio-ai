"""
news_sentiment.py
Groq-powered sentiment for each holding:
  - FLAW 2.1: 5-day sentiment memory fed into the prompt for consistency
  - FLAW 2.2: Tiered news fetching (full article -> RSS snippet -> headline only)
  - FLAW 2.3: Earnings calendar check — warns if earnings within 5 days
  - RELIABILITY: all Groq calls go through llm_client (model fallback, pacing,
                 structured JSON). A failed call is returned with
                 analysis_failed=True instead of a fake NEUTRAL verdict.
"""

import json
import re
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import requests
from datetime import datetime, date

from llm_client import ask_json


# ── FLAW 2.3: Earnings Calendar Check ────────────────────────────────────────

def check_earnings_alert(ticker_yf):
    """
    Returns earnings alert string if earnings are within 5 days.
    Returns None if no upcoming earnings detected.
    """
    try:
        import yfinance as yf
        stock = yf.Ticker(ticker_yf)
        cal   = stock.calendar
        if cal is None or (hasattr(cal, 'empty') and cal.empty):
            return None

        earnings_dates = []
        # Try DataFrame format
        if hasattr(cal, 'columns'):
            for col in cal.columns:
                if 'Earnings' in str(col):
                    val = cal[col].iloc[0] if len(cal) > 0 else None
                    if val is not None:
                        earnings_dates.append(val)
        # Try dict format
        if not earnings_dates and isinstance(cal, dict):
            ed = cal.get('Earnings Date')
            if ed:
                earnings_dates = ed if isinstance(ed, list) else [ed]

        for ed in earnings_dates:
            try:
                ed_date = ed.date() if hasattr(ed, 'date') else datetime.strptime(str(ed)[:10], '%Y-%m-%d').date()
                days_until = (ed_date - date.today()).days
                if 0 <= days_until <= 5:
                    return f"EARNINGS IN {days_until} DAYS ({ed_date.strftime('%d %b %Y')})"
            except:
                continue
        return None
    except:
        return None


# ── FLAW 2.2: Tiered News Fetching ───────────────────────────────────────────

def _fetch_full_article(url):
    """Tier 1: Try to fetch full article text from URL"""
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        resp = requests.get(url, headers=headers, timeout=8)
        if resp.status_code == 200:
            text  = re.sub(r'<[^>]+>', ' ', resp.text)
            text  = re.sub(r'\s+', ' ', text).strip()
            words = text.split()
            return ' '.join(words[:700]) if len(words) > 50 else None
        return None
    except:
        return None


def fetch_news_headlines(stock_name, max_articles=5):
    """
    FLAW 2.2: Tiered news fetching.
    Tier 1: Full article content (best)
    Tier 2: RSS description snippet (medium)
    Tier 3: Headline only (lowest — flagged with low confidence)
    """
    query = urllib.parse.quote(f"{stock_name} NSE stock India")
    url   = f"https://news.google.com/rss/search?q={query}&hl=en-IN&gl=IN&ceid=IN:en"
    articles = []

    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            content = response.read()

        root  = ET.fromstring(content)
        items = root.findall('.//item')

        for item in items[:max_articles]:
            title    = item.find('title')
            pub_date = item.find('pubDate')
            desc     = item.find('description')
            link     = item.find('link')

            if title is None:
                continue

            article = {
                'title':      title.text,
                'date':       pub_date.text if pub_date is not None else 'Unknown',
                'content':    None,
                'tier':       3,
                'confidence': 'low'
            }

            # Tier 2: RSS description snippet
            if desc is not None and desc.text:
                clean_desc = re.sub(r'<[^>]+>', ' ', desc.text).strip()
                if len(clean_desc) > 80:
                    article['content']    = clean_desc[:600]
                    article['tier']       = 2
                    article['confidence'] = 'medium'

            # Tier 1: Full article
            if link is not None and link.text:
                full_text = _fetch_full_article(link.text)
                if full_text:
                    article['content']    = full_text
                    article['tier']       = 1
                    article['confidence'] = 'high'

            articles.append(article)

    except Exception as e:
        print(f"[news_sentiment] Error fetching news for {stock_name}: {e}")

    return articles


# ── FLAW 2.1: Build Memory Context ───────────────────────────────────────────

def _build_memory_context(sentiment_history):
    """Format last 5 verdicts for the prompt"""
    if not sentiment_history:
        return "No previous verdict history available."
    lines = []
    for entry in sentiment_history[-5:]:
        lines.append(f"  - {entry.get('date', 'N/A')}: {entry.get('verdict', 'N/A')}")
    return "\n".join(lines)


# ── Structured output definition ─────────────────────────────────────────────

SENTIMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "news_sentiment":    {"type": "string", "enum": ["POSITIVE", "NEGATIVE", "NEUTRAL"]},
        "overall_sentiment": {"type": "string", "enum": ["BULLISH", "BEARISH", "NEUTRAL"]},
        "ai_decision":       {"type": "string", "enum": ["SELL", "HOLD", "WATCH"]},
        "confidence":        {"type": "string", "enum": ["HIGH", "MEDIUM", "LOW"]},
        "sentiment_flip":    {"type": "string", "enum": ["YES", "NO"]},
        "recommendation":    {"type": "string"},
        "risk_factors":      {"type": "string"},
        "target_price":      {"type": "number"},
        "stop_loss":         {"type": "number"},
    },
    "required": ["news_sentiment", "overall_sentiment", "ai_decision", "confidence",
                 "sentiment_flip", "recommendation", "risk_factors",
                 "target_price", "stop_loss"],
    "additionalProperties": False,
}

_ENUMS = {k: v["enum"] for k, v in SENTIMENT_SCHEMA["properties"].items() if "enum" in v}


def _valid_sentiment(d):
    """Reject replies whose enum fields are missing or outside the allowed values."""
    if not isinstance(d, dict):
        return False
    for key, allowed in _ENUMS.items():
        if str(d.get(key, '')).strip().upper() not in allowed:
            return False
    return True


def _to_price(value):
    try:
        v = float(value)
        return v if v > 0 else None
    except (TypeError, ValueError):
        return None


def _failed_result(reason):
    return {
        'news_sentiment':    'NEUTRAL',
        'overall_sentiment': 'NEUTRAL',
        'ai_decision':       'HOLD',
        'ai_confidence':     'LOW',
        'sentiment_flip':    'NO',
        'recommendation':    f'Analysis failed: {reason}',
        'risk_factors':      'N/A',
        'ai_target_price':   None,
        'ai_stop_loss':      None,
        'analysis_failed':   True,     # callers must NOT store this as a real verdict
        'llm_model':         '',
    }


# ── MAIN GROQ ANALYSIS ────────────────────────────────────────────────────────

def analyze_sentiment_with_groq(stock_name, ticker, articles, technical_data, stock_data, sentiment_history=None):
    # Build news text with tier context
    if not articles:
        news_text = "No recent news found."
    else:
        news_lines = []
        for a in articles:
            tier_label = {1: '[Full Article]', 2: '[Snippet]', 3: '[Headline Only - low confidence]'}.get(a.get('tier', 3), '[Unknown]')
            content    = a.get('content') or a.get('title', '')
            news_lines.append(f"- {tier_label} {a.get('title','')} ({a.get('date','')})\n  {content[:300]}")
        news_text = "\n".join(news_lines)

    rsi      = technical_data.get('rsi', 'N/A')
    macd     = technical_data.get('macd', 'N/A')
    macd_sig = technical_data.get('macd_signal', 'N/A')
    adx      = technical_data.get('adx', 'N/A')
    stoch_k  = technical_data.get('stoch_k', 'N/A')
    ema50    = technical_data.get('ema50', 'N/A')
    ema200   = technical_data.get('ema200', 'N/A')
    bb_upper = technical_data.get('bb_upper', 'N/A')
    bb_lower = technical_data.get('bb_lower', 'N/A')
    bull_pct = technical_data.get('bull_pct', 50)
    tech_notes = '\n'.join([f"  - {n}" for n in technical_data.get('technical_notes', [])])

    # Insufficient indicators note (Flaw 1.3)
    insufficient = technical_data.get('insufficient_indicators', [])
    insuff_note  = f"\nNote: Skipped due to insufficient data: {', '.join(insufficient)}" if insufficient else ""

    buying_price  = stock_data.get('buying_price', 0)
    current_price = stock_data.get('current_price', 0)
    growth_pct    = stock_data.get('growth_pct', 0)
    qty           = stock_data.get('qty', 0)
    total_profit  = round((current_price - buying_price) * qty, 2) if current_price and buying_price else 0

    if growth_pct is not None:
        if growth_pct <= -20:
            situation = f"LOSS ALERT: Stock is down {abs(growth_pct):.1f}% from buying price"
        elif growth_pct >= 70:
            situation = f"PROFIT TARGET: Stock is up {growth_pct:.1f}% from buying price"
        else:
            situation = f"Normal holding: Stock is {'up' if growth_pct >= 0 else 'down'} {abs(growth_pct):.1f}%"
    else:
        situation = "Price data unavailable"

    # FLAW 2.1: Memory context
    memory_context = _build_memory_context(sentiment_history)

    prompt = f"""You are an expert Indian stock market analyst and portfolio manager.
Analyze ALL the following data carefully for {stock_name} (NSE: {ticker}) and give a precise recommendation.

PORTFOLIO SITUATION:
- Buying Price: Rs.{buying_price}
- Current Price: Rs.{current_price}
- Growth: {growth_pct}%
- Quantity: {qty} shares
- Total P&L: Rs.{total_profit}
- Situation: {situation}

PREVIOUS VERDICTS (last 5 trading days — use for consistency):
{memory_context}
Important: If your verdict today differs from the recent trend, explicitly explain why in the recommendation field.
If sentiment flips from BULLISH to BEARISH with no major price move or news catalyst, set sentiment_flip to YES.

TRADING RULES (follow strictly):
- SELL immediately if: Loss >= 20% AND overall sentiment is BEARISH
- HOLD for recovery if: Loss >= 20% BUT sentiment is BULLISH
- SELL if: Profit >= 70% AND no strong uptrend
- HOLD to peak if: Profit >= 70% AND strong BULLISH signals
- HOLD normally if: Between -20% and +70%

TECHNICAL INDICATORS:{insuff_note}
- RSI (14): {rsi}
- MACD: {macd} vs Signal {macd_sig}
- ADX: {adx} (>25 = strong trend)
- Stochastic K: {stoch_k}
- EMA50: {ema50} vs EMA200: {ema200}
- Bollinger Bands: Upper {bb_upper} / Lower {bb_lower}
- Overall Bull Score: {bull_pct}%
Technical Notes:
{tech_notes}

RECENT NEWS (tiered by confidence):
{news_text}

Respond with ONE JSON object only, with these fields:
- news_sentiment: POSITIVE, NEGATIVE or NEUTRAL
- overall_sentiment: BULLISH, BEARISH or NEUTRAL
- ai_decision: SELL, HOLD or WATCH
- confidence: HIGH, MEDIUM or LOW
- sentiment_flip: YES or NO
- recommendation: 2-3 sentence reasoning (explain why if the verdict changed from the recent trend)
- risk_factors: 1-2 key risks
- target_price: realistic target price in Rs. as a number (for example 1850.00)
- stop_loss: stop loss price in Rs. as a number (for example 1200.00)"""

    res = ask_json(prompt, SENTIMENT_SCHEMA, name='stock_sentiment',
                   validate=_valid_sentiment, max_tokens=1200, label=f"sentiment {ticker}")

    if not res.ok:
        print(f"[news_sentiment] Groq FAILED for {stock_name}: {res.error}")
        return _failed_result(res.error[:200])

    d = res.data
    return {
        'news_sentiment':    str(d['news_sentiment']).strip().upper(),
        'overall_sentiment': str(d['overall_sentiment']).strip().upper(),
        'ai_decision':       str(d['ai_decision']).strip().upper(),
        'ai_confidence':     str(d['confidence']).strip().upper(),
        'sentiment_flip':    str(d['sentiment_flip']).strip().upper(),
        'recommendation':    str(d.get('recommendation') or 'No recommendation generated.').strip(),
        'risk_factors':      str(d.get('risk_factors') or 'N/A').strip(),
        'ai_target_price':   _to_price(d.get('target_price')),
        'ai_stop_loss':      _to_price(d.get('stop_loss')),
        'raw_response':      json.dumps(d),
        'analysis_failed':   False,
        'llm_model':         res.model,
    }


def get_full_sentiment(stock_name, ticker, technical_data, stock_data, sentiment_history=None):
    """Main entry point — fetch news, check earnings, run Groq with memory"""
    ticker_yf = ticker + ".NS"

    # FLAW 2.3: Earnings calendar check
    earnings_alert = check_earnings_alert(ticker_yf)
    if earnings_alert:
        print(f"[news_sentiment] {ticker}: {earnings_alert}")

    # FLAW 2.2: Tiered news fetch
    articles = fetch_news_headlines(stock_name)

    # FLAW 2.1: Groq with 5-day memory context
    sentiment = analyze_sentiment_with_groq(
        stock_name, ticker, articles, technical_data, stock_data,
        sentiment_history=sentiment_history
    )
    sentiment['headlines']      = articles
    sentiment['earnings_alert'] = earnings_alert

    return sentiment


if __name__ == "__main__":
    tech = {
        'rsi': 45, 'macd': -2.3, 'macd_signal': -1.8,
        'adx': 28, 'stoch_k': 35, 'ema50': 1380, 'ema200': 1290,
        'bb_upper': 1500, 'bb_lower': 1300, 'bull_pct': 55,
        'technical_notes': ['RSI 45 neutral', 'MACD bearish', 'ADX strong trend'],
        'insufficient_indicators': []
    }
    stock = {
        'buying_price': 1210.60, 'current_price': 1424.70,
        'growth_pct': 17.69, 'qty': 40
    }
    history = [
        {'date': '2026-04-07', 'verdict': 'BULLISH'},
        {'date': '2026-04-08', 'verdict': 'BULLISH'},
        {'date': '2026-04-09', 'verdict': 'NEUTRAL'},
    ]
    result = get_full_sentiment("Reliance Industries", "RELIANCE", tech, stock, sentiment_history=history)
    print(result)
