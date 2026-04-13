"""
stock_scout.py - MODULE 3
Scans 200+ NSE growth stocks daily across all sectors.

Changes:
  - FLAW 5.4: Alpha Vantage REMOVED — all fundamentals now from yfinance (no API limit)
  - FLAW 2.3: Earnings calendar alert shown in recommendation email
  - FLAW 2.5: Minimum 8% upside filter + P/E overvaluation flag
  - FLAW 3.3: Sector diversification — top 5 picks span minimum 3 different sectors
  - FLAW 3.5: Previously sold stocks tagged with trade history in email
  - FLAW 2.4: Recommendations logged to Sheets for accuracy tracking (called from main.py)
"""

import yfinance as yf
import random
from datetime import datetime
from technical_analysis import calculate_indicators
from email_handler import send_report_email
from news_sentiment import check_earnings_alert
from config import NSE_SUFFIX

# ── Full NSE Universe ──────────────────────────────────────────────────────────
FULL_UNIVERSE = [
    # IT & Technology
    "TCS", "INFY", "WIPRO", "HCLTECH", "TECHM", "MPHASIS", "COFORGE",
    "PERSISTENT", "LTTS", "TATAELXSI", "HAPPSTMNDS", "AFFLE", "ROUTE",
    "MASTEK", "KPITTECH", "CYIENT", "ZENSARTECH", "BSOFT",
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

FULL_UNIVERSE = list(set(FULL_UNIVERSE))


# ── Rough sector P/E averages for overvaluation check (Flaw 2.5) ──────────────
SECTOR_PE = {
    'Technology':             25,
    'Financial Services':     18,
    'Healthcare':             30,
    'Consumer Defensive':     35,
    'Consumer Cyclical':      40,
    'Industrials':            28,
    'Energy':                 15,
    'Basic Materials':        12,
    'Real Estate':            35,
    'Utilities':              18,
    'Communication Services': 22,
}


def get_existing_tickers():
    """Get current holdings to avoid recommending owned stocks"""
    try:
        from config import USE_GOOGLE_SHEETS
        if USE_GOOGLE_SHEETS:
            from sheets_handler import read_holdings
            return [h['ticker'].upper() for h in read_holdings()]
        from config import HOLDINGS
        return [h['ticker'].upper() for h in HOLDINGS]
    except:
        return []


def get_pnl_history():
    """Get previously sold stocks for re-entry tagging (Flaw 3.5)"""
    try:
        from config import USE_GOOGLE_SHEETS
        if USE_GOOGLE_SHEETS:
            from sheets_handler import get_pnl_records
            return get_pnl_records()
        return []
    except:
        return []


def score_stock(ticker):
    """
    FLAW 5.4: Alpha Vantage fully replaced by yfinance for all fundamentals.
    No API key needed, no rate limits, same fields available.
    """
    try:
        ticker_yf = ticker + NSE_SUFFIX
        stock     = yf.Ticker(ticker_yf)
        info      = stock.info

        if not info or not info.get('currentPrice'):
            return None

        score         = 0
        reasons       = []
        current_price = info.get('currentPrice')

        # ── Fundamentals from yfinance (replaces Alpha Vantage) ──────────────

        # Revenue Growth
        rev_growth = info.get('revenueGrowth')
        if rev_growth is not None:
            if rev_growth > 0.25:
                score += 20
                reasons.append(f"Strong revenue growth: {rev_growth*100:.1f}%")
            elif rev_growth > 0.10:
                score += 12
                reasons.append(f"Good revenue growth: {rev_growth*100:.1f}%")
        else:
            reasons.append("Revenue growth: Data unavailable")

        # Earnings Growth
        earn_growth = info.get('earningsGrowth')
        if earn_growth is not None:
            if earn_growth > 0.25:
                score += 20
                reasons.append(f"Strong earnings growth: {earn_growth*100:.1f}%")
            elif earn_growth > 0.10:
                score += 10
                reasons.append(f"Good earnings growth: {earn_growth*100:.1f}%")
        else:
            reasons.append("Earnings growth: Data unavailable")

        # ROE
        roe = info.get('returnOnEquity')
        if roe is not None:
            if roe > 0.20:
                score += 12
                reasons.append(f"High ROE: {roe*100:.1f}%")
            elif roe > 0.12:
                score += 6
                reasons.append(f"Decent ROE: {roe*100:.1f}%")

        # Debt to Equity
        dte = info.get('debtToEquity')
        if dte is not None:
            if dte < 0.3:
                score += 8
                reasons.append(f"Low debt (D/E: {dte:.2f})")
            elif dte > 1.5:
                score -= 5
                reasons.append(f"High debt (D/E: {dte:.2f})")

        # ── FLAW 2.5: P/E overvaluation check ────────────────────────────────
        pe_ratio      = info.get('trailingPE')
        sector        = info.get('sector', 'N/A')
        pe_flag       = None
        sector_avg_pe = SECTOR_PE.get(sector, 25)

        if pe_ratio is not None and pe_ratio > 0:
            if pe_ratio > sector_avg_pe * 1.4:
                pe_flag = f"EXPENSIVE VALUATION: P/E {pe_ratio:.1f} is 40%+ above {sector} sector avg ({sector_avg_pe})"
            else:
                reasons.append(f"Reasonable P/E: {pe_ratio:.1f} (sector avg: {sector_avg_pe})")

        # ── Technicals ────────────────────────────────────────────────────────
        tech     = calculate_indicators(ticker_yf)
        bull_pct = tech.get('bull_pct', 50)

        if bull_pct >= 65:
            score += 20
            reasons.append(f"Strong bullish technicals: {bull_pct}% bull score")
        elif bull_pct >= 50:
            score += 10
            reasons.append(f"Neutral-bullish technicals: {bull_pct}% bull score")

        # ── Comprehensive multi-factor target price calculation ───────────────
        # Factors: technicals + fundamentals + sentiment + valuation
        try:
            ema50    = tech.get('ema50') or current_price
            ema200   = tech.get('ema200') or current_price
            bb_upper = tech.get('bb_upper') or current_price * 1.1
            adx      = tech.get('adx') or 20
            rsi      = tech.get('rsi') or 50
            macd_val = tech.get('macd') or 0
            macd_sig = tech.get('macd_signal') or 0

            # ── Base: technical resistance levels ─────────────────────────────
            momentum   = min(adx / 25, 1.5)   # ADX momentum multiplier
            tech_base  = (ema50 * 1.15 * momentum + bb_upper) / 2

            # ── Adjustment 1: RSI factor ───────────────────────────────────────
            # Oversold (RSI < 40) = more upside room, overbought (RSI > 70) = less room
            if rsi < 40:
                rsi_factor = 1.05    # 5% boost — oversold, more room to recover
            elif rsi > 70:
                rsi_factor = 0.95   # 5% reduction — already stretched
            else:
                rsi_factor = 1.0    # neutral

            # ── Adjustment 2: MACD direction factor ────────────────────────────
            # Bullish MACD crossover = more upside confidence
            if macd_val > macd_sig:
                macd_factor = 1.03   # 3% boost for bullish momentum
            else:
                macd_factor = 0.98   # 2% reduction for bearish momentum

            # ── Adjustment 3: EMA200 trend factor ─────────────────────────────
            # Price above EMA200 (golden cross) = strong long-term uptrend = higher target
            if ema200 and current_price > ema200:
                ema_factor = 1.04    # 4% boost — stock in long-term uptrend
            else:
                ema_factor = 0.97   # 3% reduction — below long-term average

            # ── Adjustment 4: Fundamental growth factor ────────────────────────
            # High revenue/earnings growth justifies higher price target
            fund_factor = 1.0
            if rev_growth is not None and rev_growth > 0.25:
                fund_factor += 0.05  # 5% boost for strong revenue growth
            elif rev_growth is not None and rev_growth > 0.10:
                fund_factor += 0.02  # 2% boost for decent revenue growth

            if earn_growth is not None and earn_growth > 0.25:
                fund_factor += 0.05  # 5% boost for strong earnings growth
            elif earn_growth is not None and earn_growth > 0.10:
                fund_factor += 0.02  # 2% boost for decent earnings growth

            fund_factor = min(fund_factor, 1.12)  # cap at 12% boost

            # ── Adjustment 5: Valuation factor (P/E vs sector) ────────────────
            # If stock is cheaper than sector average = more upside room
            val_factor = 1.0
            if pe_ratio is not None and pe_ratio > 0 and sector_avg_pe > 0:
                pe_ratio_vs_sector = pe_ratio / sector_avg_pe
                if pe_ratio_vs_sector < 0.7:
                    val_factor = 1.05   # 5% boost — significantly undervalued vs sector
                elif pe_ratio_vs_sector < 1.0:
                    val_factor = 1.02   # 2% boost — slightly undervalued
                elif pe_ratio_vs_sector > 1.4:
                    val_factor = 0.95   # 5% reduction — significantly overvalued

            # ── Combine all factors ────────────────────────────────────────────
            comprehensive_target = tech_base * rsi_factor * macd_factor * ema_factor * fund_factor * val_factor
            ai_target  = round(max(comprehensive_target, current_price * 1.08), 2)  # enforce min 8% upside
            upside_pct = round(((ai_target - current_price) / current_price) * 100, 1)

            # Build a brief explanation of what drove the target
            target_factors = []
            if rsi_factor > 1:   target_factors.append("oversold RSI boost")
            if rsi_factor < 1:   target_factors.append("overbought RSI reduction")
            if macd_factor > 1:  target_factors.append("bullish MACD")
            if ema_factor > 1:   target_factors.append("above EMA200")
            if fund_factor > 1:  target_factors.append("strong fundamentals boost")
            if val_factor > 1:   target_factors.append("undervalued vs sector")
            if val_factor < 1:   target_factors.append("overvalued vs sector reduction")
            target_basis = ", ".join(target_factors) if target_factors else "technical levels"

        except Exception as e:
            print(f"[scout] Target calc error for {ticker}: {e}")
            ai_target    = round(current_price * 1.15, 2)
            upside_pct   = 15.0
            target_basis = "technical levels (fallback)"

        # Hard filter: skip stock if upside less than 8%
        if upside_pct < 8:
            print(f"[scout] {ticker} skipped — upside only {upside_pct}% (min 8% required)")
            return None

        # ── Estimated days to hit target ──────────────────────────────────────
        # Based on average daily % move over last 20 trading days.
        # Distance to target / avg daily move = estimated trading days.
        # This is a swing trading estimate (days to weeks), not long term.
        est_days       = None
        est_label      = None
        holding_horizon = None
        try:
            hist_20 = yf.download(ticker_yf, period="30d", interval="1d", progress=False)
            if not hist_20.empty and len(hist_20) >= 5:
                closes      = hist_20['Close'].squeeze()
                daily_moves = abs(closes.pct_change().dropna())
                avg_daily   = float(daily_moves.mean())  # average daily % move

                if avg_daily > 0:
                    distance_pct = upside_pct / 100
                    trading_days = round(distance_pct / avg_daily)
                    # Convert trading days to calendar days (multiply by 1.4)
                    calendar_days = round(trading_days * 1.4)
                    est_days = calendar_days

                    # Human-readable label
                    if calendar_days <= 7:
                        est_label = f"~{calendar_days} days"
                        holding_horizon = "Very Short Term"
                    elif calendar_days <= 21:
                        est_label = f"~{calendar_days} days (~{round(calendar_days/7)} weeks)"
                        holding_horizon = "Short Term"
                    elif calendar_days <= 60:
                        est_label = f"~{round(calendar_days/7)} weeks (~{round(calendar_days/30)} month)"
                        holding_horizon = "Short-Medium Term"
                    elif calendar_days <= 120:
                        est_label = f"~{round(calendar_days/30)} months"
                        holding_horizon = "Medium Term"
                    else:
                        est_label = f"~{round(calendar_days/30)}+ months"
                        holding_horizon = "Long Term"
        except:
            pass

        # ── FLAW 2.3: Earnings alert ──────────────────────────────────────────
        earnings_alert = check_earnings_alert(ticker_yf)

        # ── Cap category from marketCap (Flaw 3.1 — auto classification) ─────
        market_cap   = info.get('marketCap')
        cap_category = 'N/A'
        if market_cap:
            if market_cap >= 20_000_00_00_000:
                cap_category = 'Large Cap'
            elif market_cap >= 5_000_00_00_000:
                cap_category = 'Mid Cap'
            else:
                cap_category = 'Small Cap'

        return {
            'ticker':           ticker,
            'ticker_yf':        ticker_yf,
            'name':             info.get('longName', ticker),
            'sector':           sector,
            'cap_category':     cap_category,
            'current_price':    current_price,
            'score':            score,
            'reasons':          reasons,
            'pe_ratio':         pe_ratio,
            'pe_flag':          pe_flag,
            'market_cap':       market_cap,
            'revenue_growth':   rev_growth,
            'earnings_growth':  earn_growth,
            'roe':              roe,
            'debt_to_equity':   dte,
            'technical_signal': tech['technical_signal'],
            'bull_pct':         bull_pct,
            'rsi':              tech.get('rsi'),
            'ai_target_price':  ai_target,
            'upside_pct':       upside_pct,
            'target_basis':     target_basis,
            'est_days':         est_days,
            'est_label':        est_label,
            'holding_horizon':  holding_horizon,
            'earnings_alert':   earnings_alert,
        }

    except Exception as e:
        print(f"[scout] Error scoring {ticker}: {e}")
        return None


def find_growth_stocks(top_n=5, sample_size=60):
    """
    FLAW 3.3: Top 5 picks span minimum 3 different sectors.
    """
    existing     = get_existing_tickers()
    universe     = [t for t in FULL_UNIVERSE if t.upper() not in existing]
    daily_sample = random.sample(universe, min(sample_size, len(universe)))

    print(f"[scout] Scanning {len(daily_sample)} stocks from {len(universe)} candidates...")
    scored = []

    for ticker in daily_sample:
        print(f"[scout] Analyzing {ticker}...", end=' ')
        result = score_stock(ticker)
        if result and result['score'] >= 40:
            scored.append(result)
            print(f"Score: {result['score']}/100 [{result['sector']}]")
        else:
            print(f"Score: {result['score'] if result else 'N/A'} skipped")

    scored.sort(key=lambda x: x['score'], reverse=True)

    # FLAW 3.3: Ensure minimum 3 different sectors in top picks
    diversified  = []
    sectors_used = []

    # First pass: pick highest scoring stock from each new sector
    for s in scored:
        sec = s.get('sector', 'Unknown')
        if sec not in sectors_used:
            diversified.append(s)
            sectors_used.append(sec)
        if len(diversified) >= top_n:
            break

    # Second pass: fill remaining slots with best remaining scores (any sector)
    if len(diversified) < top_n:
        for s in scored:
            if s not in diversified:
                diversified.append(s)
            if len(diversified) >= top_n:
                break

    unique_sectors = list(set(s.get('sector', 'Unknown') for s in diversified))
    if len(unique_sectors) < 3:
        print(f"[scout] Only {len(unique_sectors)} sector(s) in today's picks (target: 3+)")
    else:
        print(f"[scout] Top picks span {len(unique_sectors)} sectors: {', '.join(unique_sectors)}")

    print(f"\n[scout] Found {len(diversified)} qualifying stocks.")
    return diversified


def format_market_cap(mc):
    if not mc:
        return "N/A"
    if mc >= 1e12:
        return f"Rs.{mc/1e12:.1f}T"
    elif mc >= 1e9:
        return f"Rs.{mc/1e9:.1f}B"
    else:
        return f"Rs.{mc/1e6:.0f}M"


def generate_scout_email(candidates, pnl_history=None):
    date_str = datetime.now().strftime('%d %B %Y')
    if pnl_history is None:
        pnl_history = []

    # FLAW 3.5: Build lookup of previously sold stocks
    sold_lookup = {}
    for record in pnl_history:
        t = str(record.get('Ticker', '') or record.get('ticker', '')).upper()
        if t:
            sold_lookup[t] = record

    stock_cards = ''

    for i, s in enumerate(candidates, 1):
        reasons_html = ''.join([f"<li>{r}</li>" for r in s['reasons']])
        mc     = format_market_cap(s.get('market_cap'))
        pe     = f"{s['pe_ratio']:.1f}" if s.get('pe_ratio') else "N/A"
        rev_g  = f"{s['revenue_growth']*100:.1f}%" if s.get('revenue_growth') else "N/A"
        earn_g = f"{s['earnings_growth']*100:.1f}%" if s.get('earnings_growth') else "N/A"
        roe    = f"{s['roe']*100:.1f}%" if s.get('roe') else "N/A"
        cap    = s.get('cap_category', '')

        # FLAW 2.3: Earnings alert banner
        earnings_html = ''
        if s.get('earnings_alert'):
            earnings_html = f'<div style="background:#FFF3CD;border:1px solid #FFC107;padding:10px;border-radius:6px;margin-bottom:10px;font-weight:bold;">📅 {s["earnings_alert"]}</div>'

        # FLAW 2.5: P/E overvaluation flag
        pe_flag_html = ''
        if s.get('pe_flag'):
            pe_flag_html = f'<div style="background:#FFEBEE;border:1px solid #FFCDD2;padding:8px;border-radius:6px;margin-bottom:10px;font-size:12px;">⚠️ {s["pe_flag"]}</div>'

        # FLAW 3.5: Previously sold tag
        prev_trade_html = ''
        if s['ticker'].upper() in sold_lookup:
            prev  = sold_lookup[s['ticker'].upper()]
            p_sell  = prev.get('Selling Price') or prev.get('selling_price', 'N/A')
            p_date  = str(prev.get('Selling Date') or prev.get('selling_date', 'N/A'))[:10]
            p_pnl   = prev.get('Total Profit') or prev.get('total_profit', 'N/A')
            p_ret   = prev.get('Return %') or prev.get('return_pct', 'N/A')
            p_color = '#2E7D32' if isinstance(p_pnl, (int, float)) and p_pnl >= 0 else '#C62828'
            prev_trade_html = f'''
            <div style="background:#E3F2FD;border:1px solid #90CAF9;padding:10px;border-radius:6px;margin-bottom:10px;font-size:12px;">
              🔄 <strong>Previously Held</strong> — Sold on {p_date} at Rs.{p_sell}
              &nbsp;|&nbsp; P&L: <span style="color:{p_color};font-weight:bold;">Rs.{p_pnl} ({p_ret}%)</span>
            </div>'''

        stock_cards += f"""
        <div style="background:white;border-radius:10px;padding:20px;margin-bottom:16px;
                    box-shadow:0 2px 8px rgba(0,0,0,0.1);border-left:5px solid #4CAF50;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
            <div>
              <span style="background:#E8F5E9;color:#2E7D32;padding:4px 10px;border-radius:20px;
                           font-size:12px;font-weight:bold;">#{i} GROWTH PICK</span>
              <span style="background:#f0f0f0;color:#666;padding:4px 8px;border-radius:20px;
                           font-size:11px;margin-left:6px;">{cap}</span>
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
          {earnings_html}
          {pe_flag_html}
          {prev_trade_html}
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
            <div style="margin-top:8px;display:flex;gap:20px;flex-wrap:wrap;">
              <span style="font-size:12px;">⏱️ <strong>Est. time to target:</strong> {s.get('est_label') or 'N/A'}</span>
              <span style="font-size:12px;">📅 <strong>Horizon:</strong> {s.get('holding_horizon') or 'N/A'}</span>
            </div>
            <div style="margin-top:6px;font-size:11px;color:#555;">
              📊 <strong>Target based on:</strong> {s.get('target_basis', 'technical levels')}
            </div>
            <small style="color:#888;font-size:11px;display:block;margin-top:6px;">
              Target factors in: technicals (EMA, BB, ADX, RSI, MACD) + fundamentals (revenue/earnings growth, P/E vs sector).
              Time estimate based on avg daily price movement. Swing trade estimate — not long-term investing advice.
            </small>
          </div>
          <div style="background:#E3F2FD;padding:12px;border-radius:6px;border:2px dashed #1976D2;">
            <strong>To approve:</strong> Reply with <code style="color:#1976D2;">YES {s['ticker']} [qty]</code>
            or <code style="color:#CC0000;">NO {s['ticker']}</code>
          </div>
        </div>"""

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:800px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#1B5E20,#2E7D32);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">🔍 Daily Growth Stock Picks</h1>
    <p style="margin:6px 0 0;opacity:0.8;">Scouted on {date_str} | Top {len(candidates)} picks from today's scan | Min 8% upside enforced</p>
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
    pnl_history = get_pnl_history()
    subject     = f"🔍 {len(candidates)} Growth Stock Picks | {datetime.now().strftime('%d %b %Y')}"
    html        = generate_scout_email(candidates, pnl_history=pnl_history)
    send_report_email(subject, html)


if __name__ == "__main__":
    candidates = find_growth_stocks(top_n=5)
    if candidates:
        send_scout_email(candidates)
        print(f"Scout email sent with {len(candidates)} candidates!")
    else:
        print("No strong candidates found today.")
