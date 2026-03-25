"""
news_sentiment.py
Upgraded Groq prompting — feeds full technical picture + T&C rules
for much stronger and more accurate AI decisions.
"""

from groq import Groq
from config import GROQ_API_KEY, GROQ_MODEL
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET

client = Groq(api_key=GROQ_API_KEY)


def fetch_news_headlines(stock_name, max_articles=5):
    query = urllib.parse.quote(f"{stock_name} NSE stock India")
    url = f"https://news.google.com/rss/search?q={query}&hl=en-IN&gl=IN&ceid=IN:en"
    headlines = []
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            content = response.read()
        root = ET.fromstring(content)
        items = root.findall('.//item')
        for item in items[:max_articles]:
            title = item.find('title')
            pub_date = item.find('pubDate')
            if title is not None:
                headlines.append({
                    'title': title.text,
                    'date': pub_date.text if pub_date is not None else 'Unknown'
                })
    except Exception as e:
        print(f"[news_sentiment] Error fetching news for {stock_name}: {e}")
    return headlines


def analyze_sentiment_with_groq(stock_name, ticker, headlines, technical_data, stock_data):
    if not headlines:
        headlines_text = "No recent news found."
    else:
        headlines_text = "\n".join([f"- {h['title']} ({h['date']})" for h in headlines])

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

    prompt = f"""You are an expert Indian stock market analyst and portfolio manager.
Analyze ALL the following data carefully for {stock_name} (NSE: {ticker}) and give a precise recommendation.

PORTFOLIO SITUATION:
- Buying Price: Rs.{buying_price}
- Current Price: Rs.{current_price}
- Growth: {growth_pct}%
- Quantity: {qty} shares
- Total P&L: Rs.{total_profit}
- Situation: {situation}

TRADING RULES (follow strictly):
- SELL immediately if: Loss >= 20% AND overall sentiment is BEARISH
- HOLD for recovery if: Loss >= 20% BUT sentiment is BULLISH
- SELL if: Profit >= 70% AND no strong uptrend
- HOLD to peak if: Profit >= 70% AND strong BULLISH signals
- HOLD normally if: Between -20% and +70%

TECHNICAL INDICATORS:
- RSI (14): {rsi}
- MACD: {macd} vs Signal {macd_sig}
- ADX: {adx} (>25 = strong trend)
- Stochastic K: {stoch_k}
- EMA50: {ema50} vs EMA200: {ema200}
- Bollinger Bands: Upper {bb_upper} / Lower {bb_lower}
- Overall Bull Score: {bull_pct}%
Technical Notes:
{tech_notes}

RECENT NEWS:
{headlines_text}

Respond in EXACTLY this format, no extra text:
NEWS_SENTIMENT: [POSITIVE/NEGATIVE/NEUTRAL]
OVERALL_SENTIMENT: [BULLISH/BEARISH/NEUTRAL]
AI_DECISION: [SELL/HOLD/WATCH]
CONFIDENCE: [HIGH/MEDIUM/LOW]
RECOMMENDATION: [2-3 sentence reasoning]
RISK_FACTORS: [1-2 key risks]
TARGET_PRICE: [realistic target price in Rs based on EMA resistance, Bollinger upper band and trend strength. Just a number like 1850.00]
STOP_LOSS: [stop loss price in Rs based on recent support levels. Just a number like 1200.00]"""

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=350,
            temperature=0.2
        )
        text = response.choices[0].message.content.strip()

        result = {
            'news_sentiment':    'NEUTRAL',
            'overall_sentiment': 'NEUTRAL',
            'ai_decision':       'HOLD',
            'ai_confidence':     'LOW',
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
            elif line.startswith('RECOMMENDATION:'):
                result['recommendation'] = line.split(':', 1)[1].strip()
            elif line.startswith('RISK_FACTORS:'):
                result['risk_factors'] = line.split(':', 1)[1].strip()

        return result

    except Exception as e:
        print(f"[news_sentiment] Groq error for {stock_name}: {e}")
        return {
            'news_sentiment':    'NEUTRAL',
            'overall_sentiment': 'NEUTRAL',
            'ai_decision':       'HOLD',
            'ai_confidence':     'LOW',
            'recommendation':    f'Analysis failed: {e}',
            'risk_factors':      'N/A',
            'ai_target_price':   None,
            'ai_stop_loss':      None,
        }


def get_full_sentiment(stock_name, ticker, technical_data, stock_data):
    headlines = fetch_news_headlines(stock_name)
    sentiment = analyze_sentiment_with_groq(stock_name, ticker, headlines, technical_data, stock_data)
    sentiment['headlines'] = headlines
    return sentiment


if __name__ == "__main__":
    tech = {
        'rsi': 45, 'macd': -2.3, 'macd_signal': -1.8,
        'adx': 28, 'stoch_k': 35, 'ema50': 1380, 'ema200': 1290,
        'bb_upper': 1500, 'bb_lower': 1300, 'bull_pct': 55,
        'technical_notes': ['RSI 45 neutral', 'MACD bearish', 'ADX strong trend']
    }
    stock = {
        'buying_price': 1210.60, 'current_price': 1424.70,
        'growth_pct': 17.69, 'qty': 40
    }
    result = get_full_sentiment("Reliance Industries", "RELIANCE", tech, stock)
    print(result)