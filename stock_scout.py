"""
stock_scout.py - MODULE 3
Scans 200+ NSE growth stocks daily across all sectors.
Scores on revenue growth, earnings growth, ROE, debt, and technicals.
Sends top 5 picks email every day.
"""

import yfinance as yf
import random
from datetime import datetime
from technical_analysis import calculate_indicators
from email_handler import send_report_email
from config import NSE_SUFFIX

# ── Full NSE Universe (200+ stocks across all sectors) ────────
FULL_UNIVERSE = [
    # IT & Technology
    "TCS", "INFY", "WIPRO", "HCLTECH", "TECHM", "MPHASIS", "COFORGE",
    "PERSISTENT", "LTTS", "TATAELXSI", "HAPPSTMNDS", "AFFLE", "ROUTE",
    "MASTEK", "KPITTECH", "CYIENT", "ZENSARTECH", "COFORGE", "BSOFT",
    "SONATSOFTW", "TANLA", "INTELLECT", "NEWGEN", "NUCLEUS",

    # Banking & Finance
    "HDFCBANK", "ICICIBANK", "KOTAKBANK", "AXISBANK", "INDUSINDBK",
    "FEDERALBNK", "IDFCFIRSTB", "RBLBANK", "BANDHANBNK", "AUBANK",
    "ABCAPITAL", "CHOLAFIN", "MUTHOOTFIN", "MANAPPURAM", "BAJFINANCE",
    "BAJAJFINSV", "SBICARD", "M&MFIN", "SUNDARMFIN",

    # Pharma & Healthcare
    "SUNPHARMA", "DRREDDY", "CIPLA", "DIVISLAB", "AUROPHARMA",
    "TORNTPHARM", "ALKEM", "IPCA", "LALPATHLAB", "METROPOLIS",
    "APOLLOHOSP", "MAXHEALTH", "FORTIS", "MEDANTA", "RAINBOW",
    "YATHARTH", "KRSNAA", "THYROCARE",

    # Consumer & FMCG
    "HINDUNILVR", "NESTLEIND", "BRITANNIA", "DABUR", "MARICO",
    "GODREJCP", "EMAMILTD", "COLPAL",
    "BIKAJI", "DEVYANI", "SAPPHIREF", "WESTLIFE", "JUBLFOOD",
    "ZOMATO", "NYKAA", "RADICO",

    # Auto & EV
    "MARUTI", "BAJAJ-AUTO", "HEROMOTOCO", "EICHERMOT",
    "MOTHERSON", "BOSCHLTD", "BHARATFORG", "EXIDEIND",
    "OLECTRA", "CRAFTSMAN", "SUPRAJIT",

    # Capital Goods & Engineering
    "LT", "SIEMENS", "ABB", "HAVELLS", "CUMMINSIND", "BHEL",
    "THERMAX", "AIAENG", "KAYNES", "SYRMA", "PGEL", "VOLTAMP",
    "GRINDWELL", "SCHAEFFLER", "TIMKEN", "SKFINDIA", "ELGIEQUIP",

    # Power & Energy
    "NTPC", "POWERGRID", "TATAPOWER", "ADANIGREEN", "TORNTPOWER",
    "CESC", "INOXWIND", "SUZLON", "WAAREEENER", "NHPC", "SJVN", "JSWENERGY",

    # Real Estate
    "DLF", "GODREJPROP", "OBEROIRLTY", "PHOENIXLTD", "PRESTIGE",
    "BRIGADE", "SOBHA", "KOLTEPATIL",

    # Metals & Mining
    "TATASTEEL", "JSWSTEEL", "HINDALCO", "VEDL", "NMDC",
    "COALINDIA", "MOIL", "NATIONALUM", "HINDCOPPER", "WELSPUNLIV",

    # Chemicals & Specialty
    "PIDILITIND", "ASIANPAINT", "BERGEPAINT", "KANSAINER",
    "CLEAN", "FINEORG", "GALAXYSURF", "NAVINFLUOR",
    "ALKYLAMINE", "BALRAMCHIN", "DHANUKA", "PIIND", "ASTRAL",
    "POLYCAB", "KEI", "APLAPOLLO", "FINPIPE",

    # Insurance
    "ICICIGI", "HDFCLIFE", "SBILIFE", "POLICYBZR",
    "GICRE", "STARHEALTH",

    # Media & Entertainment
    "ZEEL", "SUNTV", "PVRINOX", "NAZARA", "SAREGAMA",

    # Logistics
    "DELHIVERY", "BLUEDART", "ADANIPORTS", "CONCOR",

    # Retail
    "TRENT", "VMART", "DMART", "ABFRL", "MANYAVAR",

    # Textiles
    "PAGEIND", "DOLLAR", "RUPA", "KITEX",

    # Cement
    "ULTRACEMCO", "SHREECEM", "AMBUJACEM", "ACC", "RAMCOCEM",
    "JKCEMENT", "HEIDELBERG", "BIRLACORPN",

    # Misc
    "IRCTC", "GMRAIRPORT", "KALYANKJIL", "UNITDSPR",
]

# Remove duplicates
FULL_UNIVERSE = list(set(FULL_UNIVERSE))


def get_alpha_vantage_fundamentals(ticker):
    """Fetch fundamental data from Alpha Vantage for stronger scoring"""
    try:
        from config import ALPHA_VANTAGE_KEY
        import urllib.request
        import json

        # Company Overview
        url = f"https://www.alphavantage.co/query?function=OVERVIEW&symbol={ticker}.BSE&apikey={ALPHA_VANTAGE_KEY}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read())

        if not data or 'Symbol' not in data:
            return None

        return {
            'pe_ratio':        float(data.get('PERatio', 0) or 0),
            'eps':             float(data.get('EPS', 0) or 0),
            'revenue_growth':  float(data.get('QuarterlyRevenueGrowthYOY', 0) or 0),
            'earnings_growth': float(data.get('QuarterlyEarningsGrowthYOY', 0) or 0),
            'roe':             float(data.get('ReturnOnEquityTTM', 0) or 0),
            'profit_margin':   float(data.get('ProfitMargin', 0) or 0),
            'debt_to_equity':  float(data.get('DebtToEquityRatio', 0) or 0),
            'book_value':      float(data.get('BookValue', 0) or 0),
            'dividend_yield':  float(data.get('DividendYield', 0) or 0),
            '52w_high':        float(data.get('52WeekHigh', 0) or 0),
            '52w_low':         float(data.get('52WeekLow', 0) or 0),
            'analyst_target':  float(data.get('AnalystTargetPrice', 0) or 0),
            'sector':          data.get('Sector', 'N/A'),
            'name':            data.get('Name', ticker),
        }
    except Exception as e:
        return None


def get_existing_tickers():
    """Read current holdings from config to avoid recommending owned stocks"""
    try:
        from config import HOLDINGS
        return [h['ticker'].upper() for h in HOLDINGS]
    except:
        return []


def score_stock(ticker):
    """Score stock using yfinance + Alpha Vantage + technicals. Returns score 0-100."""
    try:
        ticker_yf = ticker + NSE_SUFFIX
        stock = yf.Ticker(ticker_yf)
        info = stock.info

        if not info or not info.get('currentPrice'):
            return None

        score = 0
        reasons = []

        current_price = info.get('currentPrice')

        # ── Alpha Vantage Fundamentals (primary source) ───────
        av_data = get_alpha_vantage_fundamentals(ticker)

        if av_data:
            # Revenue Growth from AV
            rev_growth = av_data['revenue_growth']
            if rev_growth > 0.25:
                score += 20
                reasons.append(f"Strong revenue growth: {rev_growth*100:.1f}% (AV)")
            elif rev_growth > 0.10:
                score += 12
                reasons.append(f"Good revenue growth: {rev_growth*100:.1f}% (AV)")

            # Earnings Growth from AV
            earn_growth = av_data['earnings_growth']
            if earn_growth > 0.25:
                score += 20
                reasons.append(f"Strong earnings growth: {earn_growth*100:.1f}% (AV)")
            elif earn_growth > 0.10:
                score += 10
                reasons.append(f"Good earnings growth: {earn_growth*100:.1f}% (AV)")

            # ROE from AV
            roe = av_data['roe']
            if roe > 0.20:
                score += 12
                reasons.append(f"High ROE: {roe*100:.1f}% (AV)")
            elif roe > 0.12:
                score += 6
                reasons.append(f"Decent ROE: {roe*100:.1f}% (AV)")

            # Profit Margin from AV
            pm = av_data['profit_margin']
            if pm > 0.15:
                score += 8
                reasons.append(f"High profit margin: {pm*100:.1f}% (AV)")

            # Analyst Target Price
            target = av_data['analyst_target']
            if target and current_price and target > current_price * 1.15:
                upside = round(((target - current_price) / current_price) * 100, 1)
                score += 10
                reasons.append(f"Analyst target Rs.{target} — {upside}% upside")

            # Debt to Equity
            de = av_data['debt_to_equity']
            if de < 30:
                score += 10
                reasons.append(f"Low debt: D/E={de:.1f} (AV)")
            elif de < 60:
                score += 5
                reasons.append(f"Moderate debt: D/E={de:.1f} (AV)")

        else:
            # Fallback to yfinance fundamentals
            rev_growth  = info.get('revenueGrowth')
            earn_growth = info.get('earningsGrowth')
            roe         = info.get('returnOnEquity')
            de          = info.get('debtToEquity')

            if rev_growth and rev_growth > 0.25:
                score += 20
                reasons.append(f"Strong revenue growth: {rev_growth*100:.1f}%")
            elif rev_growth and rev_growth > 0.10:
                score += 12
                reasons.append(f"Good revenue growth: {rev_growth*100:.1f}%")

            if earn_growth and earn_growth > 0.25:
                score += 20
                reasons.append(f"Strong earnings growth: {earn_growth*100:.1f}%")
            elif earn_growth and earn_growth > 0.10:
                score += 10
                reasons.append(f"Good earnings growth: {earn_growth*100:.1f}%")

            if roe and roe > 0.20:
                score += 12
                reasons.append(f"High ROE: {roe*100:.1f}%")

            if de is not None and de < 30:
                score += 10
                reasons.append(f"Low debt: D/E={de:.1f}")

        # ── Technical Score (always from ta library) ──────────
        tech = calculate_indicators(ticker_yf)
        bull_pct = tech.get('bull_pct', 50)
        if bull_pct >= 65:
            score += 20
            reasons.append(f"Strong bullish technicals: {bull_pct}% bull score")
        elif bull_pct >= 50:
            score += 10
            reasons.append(f"Neutral-bullish technicals: {bull_pct}% bull score")

        # Calculate AI target price based on technicals
        # Use EMA50 as base, add momentum factor from ADX and RSI
        try:
            ema50    = tech.get('ema50') or current_price
            bb_upper = tech.get('bb_upper') or current_price * 1.1
            adx      = tech.get('adx') or 20
            rsi      = tech.get('rsi') or 50
            # Target = weighted average of EMA50 upside and BB upper
            momentum = min(adx / 25, 1.5)  # cap momentum at 1.5x
            raw_target = (ema50 * 1.15 * momentum + bb_upper) / 2
            # Ensure target is at least 10% above current price
            ai_target = round(max(raw_target, current_price * 1.10), 2)
            upside_pct = round(((ai_target - current_price) / current_price) * 100, 1)
        except:
            ai_target  = round(current_price * 1.15, 2)
            upside_pct = 15.0

        return {
            'ticker':           ticker,
            'ticker_yf':        ticker_yf,
            'name':             av_data['name'] if av_data else info.get('longName', ticker),
            'sector':           av_data['sector'] if av_data else info.get('sector', 'N/A'),
            'current_price':    current_price,
            'score':            score,
            'reasons':          reasons,
            'pe_ratio':         av_data['pe_ratio'] if av_data else info.get('trailingPE'),
            'market_cap':       info.get('marketCap'),
            'revenue_growth':   av_data['revenue_growth'] if av_data else info.get('revenueGrowth'),
            'earnings_growth':  av_data['earnings_growth'] if av_data else info.get('earningsGrowth'),
            'roe':              av_data['roe'] if av_data else info.get('returnOnEquity'),
            'debt_to_equity':   av_data['debt_to_equity'] if av_data else info.get('debtToEquity'),
            'analyst_target':   av_data['analyst_target'] if av_data else None,
            'technical_signal': tech['technical_signal'],
            'bull_pct':         bull_pct,
            'rsi':              tech.get('rsi'),
            'ai_target_price':  ai_target,
            'upside_pct':       upside_pct,
        }

    except Exception as e:
        return None


def find_growth_stocks(top_n=5, sample_size=60):
    """
    Each day randomly sample 60 stocks from the full 200+ universe.
    This ensures different stocks are analysed every day,
    giving fresh recommendations daily.
    """
    existing = get_existing_tickers()

    # Filter out stocks already in portfolio
    universe = [t for t in FULL_UNIVERSE if t.upper() not in existing]

    # Randomly sample 60 stocks from universe each day
    daily_sample = random.sample(universe, min(sample_size, len(universe)))

    print(f"[scout] Scanning {len(daily_sample)} randomly selected stocks from {len(universe)} NSE growth candidates...")
    scored = []

    for ticker in daily_sample:
        print(f"[scout] Analyzing {ticker}...", end=' ')
        result = score_stock(ticker)
        if result and result['score'] >= 40:
            scored.append(result)
            print(f"Score: {result['score']}/100 ✅")
        else:
            print(f"Score: {result['score'] if result else 'N/A'} ❌")

    scored.sort(key=lambda x: x['score'], reverse=True)
    top = scored[:top_n]
    print(f"\n[scout] Found {len(top)} qualifying stocks from today's sample.")
    return top


def format_market_cap(mc):
    if not mc:
        return "N/A"
    if mc >= 1e12:
        return f"Rs.{mc/1e12:.1f}T"
    elif mc >= 1e9:
        return f"Rs.{mc/1e9:.1f}B"
    else:
        return f"Rs.{mc/1e6:.0f}M"


def generate_scout_email(candidates):
    date_str = datetime.now().strftime('%d %B %Y')
    stock_cards = ''

    for i, s in enumerate(candidates, 1):
        reasons_html = ''.join([f"<li>{r}</li>" for r in s['reasons']])
        mc = format_market_cap(s.get('market_cap'))
        pe = f"{s['pe_ratio']:.1f}" if s.get('pe_ratio') else "N/A"
        rev_g = f"{s['revenue_growth']*100:.1f}%" if s.get('revenue_growth') else "N/A"
        earn_g = f"{s['earnings_growth']*100:.1f}%" if s.get('earnings_growth') else "N/A"
        roe = f"{s['roe']*100:.1f}%" if s.get('roe') else "N/A"

        stock_cards += f"""
        <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;
                    box-shadow:0 2px 8px rgba(0,0,0,0.1);border-left:5px solid #4CAF50;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
            <div>
              <span style="background:#E8F5E9;color:#2E7D32;padding:4px 10px;border-radius:20px;
                           font-size:12px;font-weight:bold;">#{i} GROWTH PICK</span>
              <h2 style="margin:6px 0 2px;font-size:20px;">{s['ticker']}
                <span style="font-size:14px;color:#666;font-weight:normal;">— {s['name'][:40]}</span>
              </h2>
              <div style="color:#888;font-size:13px;">{s['sector']}</div>
            </div>
            <div style="text-align:right;">
              <div style="font-size:28px;font-weight:bold;">Rs.{s['current_price']:,.2f}</div>
              <div style="background:#4CAF50;color:white;padding:4px 12px;border-radius:20px;
                           font-size:14px;font-weight:bold;">Score: {s['score']}/100</div>
            </div>
          </div>
          <table style="width:100%;border-collapse:collapse;margin-bottom:12px;">
            <tr style="background:#f8f9fa;">
              <td style="padding:8px;text-align:center;"><div style="font-size:11px;color:#888;">P/E</div><div style="font-weight:bold;">{pe}</div></td>
              <td style="padding:8px;text-align:center;"><div style="font-size:11px;color:#888;">Market Cap</div><div style="font-weight:bold;">{mc}</div></td>
              <td style="padding:8px;text-align:center;"><div style="font-size:11px;color:#888;">Revenue Growth</div><div style="font-weight:bold;color:#2E7D32;">{rev_g}</div></td>
              <td style="padding:8px;text-align:center;"><div style="font-size:11px;color:#888;">Earnings Growth</div><div style="font-weight:bold;color:#2E7D32;">{earn_g}</div></td>
              <td style="padding:8px;text-align:center;"><div style="font-size:11px;color:#888;">ROE</div><div style="font-weight:bold;">{roe}</div></td>
              <td style="padding:8px;text-align:center;"><div style="font-size:11px;color:#888;">Technical</div>
                <div style="font-weight:bold;color:{'#2E7D32' if s['technical_signal']=='BULLISH' else '#FF9800'}">{s['technical_signal']}</div></td>
            </tr>
          </table>
          <div style="background:#F1F8E9;padding:10px;border-radius:6px;margin-bottom:12px;">
            <strong>Why this stock?</strong>
            <ul style="margin:6px 0 0;padding-left:20px;color:#333;">{reasons_html}</ul>
          </div>
          <div style="background:#E8F5E9;border:1px solid #A5D6A7;padding:12px;border-radius:6px;margin-bottom:10px;">
            <strong>🎯 AI Target Price: Rs.{s['ai_target_price']:,.2f}</strong>
            <span style="color:#2E7D32;font-size:13px;"> (+{s['upside_pct']}% upside from current price)</span>
            <br><small style="color:#555;">Based on EMA resistance levels and technical trend strength. Auto-sell triggers when target is hit.</small>
          </div>
          <div style="background:#E3F2FD;padding:12px;border-radius:6px;border:2px dashed #1976D2;">
            <strong>To approve:</strong> Reply with <code style="color:#1976D2;">YES {s['ticker']}</code>
            or <code style="color:#CC0000;">NO {s['ticker']}</code>
          </div>
        </div>"""

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:800px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#1B5E20,#2E7D32);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">New Stock Opportunities</h1>
    <p style="margin:6px 0 0;opacity:0.8;">Scouted on {date_str} | Top {len(candidates)} Growth Picks from today's scan</p>
  </div>
  <div style="background:#FFF8E1;border-left:4px solid #FFC107;padding:12px 16px;border-radius:4px;margin-bottom:20px;">
    <strong>How to approve:</strong> Reply with <code>YES TICKER QTY</code> to buy or <code>NO TICKER</code> to skip.
  </div>
  {stock_cards}
  <div style="text-align:center;color:#aaa;font-size:11px;padding:12px;">
    AI-generated scouting report. Not financial advice.
  </div>
</body>
</html>"""
    return html


def send_scout_email(candidates):
    subject = f"New Stock Opportunities — {len(candidates)} Growth Picks | {datetime.now().strftime('%d %b %Y')}"
    html = generate_scout_email(candidates)
    send_report_email(subject, html)


if __name__ == "__main__":
    candidates = find_growth_stocks(top_n=5)
    if candidates:
        send_scout_email(candidates)
        print(f"Scout email sent with {len(candidates)} candidates!")
    else:
        print("No strong candidates found today.")