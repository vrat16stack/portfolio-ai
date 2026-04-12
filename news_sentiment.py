"""
news_sentiment.py
Upgraded Groq prompting with:
  - FLAW 2.1: 5-day sentiment memory fed into Groq prompt for consistency
  - FLAW 2.2: Tiered news fetching (full article -> RSS snippet -> headline only)
  - FLAW 2.3: Earnings calendar check — warns if earnings within 3 trading days
"""

from groq import Groq
from config import GROQ_API_KEY, GROQ_MODEL
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import requests
from datetime import datetime, date

client = Groq(api_key=GROQ_API_KEY)


# ── FLAW 2.3: Earnings Calendar Check ────────────────────────────────────────

def check_earnings_alert(ticker_yf):
    """
    Returns earnings alert string if earnings are within 3 trading days.
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
            import re
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
                import re
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
    """Format last 5 verdicts for Groq prompt"""
    if not sentiment_history:
        return "No previous verdict history available."
    lines = []
    for entry in sentiment_history[-5:]:
        lines.append(f"  - {entry.get('date', 'N/A')}: {entry.get('verdict', 'N/A')}")
    return "\n".join(lines)


# ── MAIN GROQ ANALYSIS ────────────────────────────────────────────────────────

def analyze_sentiment_with_groq(stock_name, ticker, articles, technical_data, stock_data, sentiment_history=None):
    # Build news text with tier context for Groq
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
Important: If your verdict today differs from the recent trend, explicitly explain why in RECOMMENDATION.
If sentiment flips from BULLISH to BEARISH with no major price move or news catalyst, flag SENTIMENT_FLIP as YES.

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

Respond in EXACTLY this format, no extra text:
NEWS_SENTIMENT: [POSITIVE/NEGATIVE/NEUTRAL]
OVERALL_SENTIMENT: [BULLISH/BEARISH/NEUTRAL]
AI_DECISION: [SELL/HOLD/WATCH]
CONFIDENCE: [HIGH/MEDIUM/LOW]
SENTIMENT_FLIP: [YES/NO]
RECOMMENDATION: [2-3 sentence reasoning. If verdict changed from recent trend, explain why.]
RISK_FACTORS: [1-2 key risks]
TARGET_PRICE: [realistic target price in Rs. Just a number like 1850.00]
STOP_LOSS: [stop loss price in Rs. Just a number like 1200.00]"""

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=400,
            temperature=0.2
        )
        text = response.choices[0].message.content.strip()

        result = {
            'news_sentiment':    'NEUTRAL',
            'overall_sentiment': 'NEUTRAL',
            'ai_decision':       'HOLD',
            'ai_confidence':     'LOW',
            'sentiment_flip':    'NO',
            'recommendation':    'No recommendation generated.',
            'risk_factors':      'N/A',
            'ai_target_price':   None,
            'ai_stop_loss':      None,
            'raw_response':      text
        }

        for line in text.split('\n'):
            line = line.strip()
            if line.startswith('NEWS_SENTIMENT:'):
                result['news_sentiment'] = line.split(':', 1)[1].strip()
            elif line.startswith('OVERALL_SENTIMENT:'):
                result['overall_sentiment'] = line.split(':', 1)[1].strip()
            elif line.startswith('AI_DECISION:'):
                result['ai_decision'] = line.split(':', 1)[1].strip()
            elif line.startswith('CONFIDENCE:'):
                result['ai_confidence'] = line.split(':', 1)[1].strip()
            elif line.startswith('SENTIMENT_FLIP:'):
                result['sentiment_flip'] = line.split(':', 1)[1].strip()
            elif line.startswith('RECOMMENDATION:'):
                result['recommendation'] = line.split(':', 1)[1].strip()
            elif line.startswith('RISK_FACTORS:'):
                result['risk_factors'] = line.split(':', 1)[1].strip()
            elif line.startswith('TARGET_PRICE:'):
                try:
                    result['ai_target_price'] = float(line.split(':', 1)[1].strip())
                except:
                    pass
            elif line.startswith('STOP_LOSS:'):
                try:
                    result['ai_stop_loss'] = float(line.split(':', 1)[1].strip())
                except:
                    pass

        return result

    except Exception as e:
        print(f"[news_sentiment] Groq error for {stock_name}: {e}")
        return {
            'news_sentiment':    'NEUTRAL',
            'overall_sentiment': 'NEUTRAL',
            'ai_decision':       'HOLD',
            'ai_confidence':     'LOW',
            'sentiment_flip':    'NO',
            'recommendation':    f'Analysis failed: {e}',
            'risk_factors':      'N/A',
            'ai_target_price':   None,
            'ai_stop_loss':      None,
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