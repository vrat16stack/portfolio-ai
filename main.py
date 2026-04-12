"""
main.py  — PORTFOLIO AI MASTER SCRIPT

Pipeline:
  1. Read holdings from Google Sheets
  2. Check if market traded today (yfinance-based — Rule 1.1)
  3. Fetch live prices from NSE (yfinance)
  4. Run technical analysis (RSI, MACD, Bollinger Bands, ADX, Stochastic, EMA, OBV)
  5. Fetch news + AI sentiment (Google News + Groq) with 5-day memory
  6. Apply decision engine (SELL/HOLD rules with 2-consecutive-BEARISH guard)
  7. Generate HTML report
  8. Send email
  9. Flag SELL stocks as PENDING_SELL — actual execution next morning at open price
 10. Scout new growth stocks (market-dependent)
 11. Weekly/Monthly summary (market-independent — runs even on holidays)

Usage:
  python main.py              # Run immediately (3:45 PM job)
  python main.py --morning    # Morning job (9:20 AM) — execute pending sells at open price
  python main.py --schedule   # Run on APScheduler
  python main.py --test       # Test with 2 stocks, no email
"""

import sys
import json
from datetime import datetime, date

from config import USE_GOOGLE_SHEETS
from price_fetcher import enrich_holdings_with_prices
from technical_analysis import calculate_indicators
from news_sentiment import get_full_sentiment
from decision_engine import process_all_holdings
from report_generator import generate_html_report, generate_subject_line
from email_handler import send_report_email
from pnl_updater import process_sell
from stock_scout import find_growth_stocks, send_scout_email
from weekly_monthly_summary import should_send_weekly, should_send_monthly, send_weekly_summary, send_monthly_summary
from fear_greed import get_fear_greed
from approval_checker import process_approvals

if USE_GOOGLE_SHEETS:
    from sheets_handler import (
        read_holdings, update_holdings_prices, update_pnl_prices,
        check_and_fix_stock_splits, log_sentiment_history,
        get_sentiment_history, log_recommendation,
        flag_pending_sell, get_pending_sells, clear_pending_sell_flag
    )
    print("[main] Using Google Sheets as data source")
else:
    from excel_reader import read_holdings
    print("[main] Using Excel as data source")


# ── RULE 1.1: Market status via yfinance (replaces hardcoded holiday list) ────

def did_market_trade_today():
    """
    Fetch Nifty 50 last traded date from yfinance.
    If it matches today -> market was open -> return True.
    Handles ALL cases: weekends, holidays, Muhurat trading, unexpected closures.
    """
    import yfinance as yf
    try:
        nifty = yf.download("^NSEI", period="5d", interval="1d", progress=False)
        if nifty.empty:
            print("[main] Could not fetch Nifty data -- treating as non-trading day.")
            return False
        last_traded = nifty.index[-1].date()
        today = date.today()
        if last_traded == today:
            print(f"[main] Market traded today ({today})")
            return True
        else:
            print(f"[main] Market closed today. Last traded: {last_traded}")
            return False
    except Exception as e:
        print(f"[main] Nifty fetch error: {e} -- treating as non-trading day.")
        return False


# ── MORNING JOB: Execute pending sells at open price (Flaw 3.4 Option A) ─────

def execute_pending_sells():
    """
    Runs at 9:20 AM IST via --morning flag.
    Checks Google Sheets for PENDING_SELL flags set yesterday at 3:45 PM.
    Fetches today's opening price and records it as the actual sell price.
    """
    import yfinance as yf
    print("\n[main] Morning job: Executing pending sells at open price...")

    try:
        pending = get_pending_sells()
        if not pending:
            print("[main] No pending sells to execute.")
            return

        for stock in pending:
            ticker_yf = stock['ticker'] + ".NS"
            try:
                hist = yf.download(ticker_yf, period="1d", interval="1m", progress=False)
                today_str = date.today().strftime('%Y-%m-%d')
                today_data = hist[hist.index.strftime('%Y-%m-%d') == today_str]

                if today_data.empty:
                    print(f"[main] No open price for {stock['ticker']} -- using decision price as fallback.")
                    open_price = stock.get('pending_sell_price') or stock.get('buying_price')
                else:
                    open_price = round(float(today_data['Open'].iloc[0]), 2)

                print(f"[main] {stock['ticker']} -- Open price: Rs.{open_price}")

                pnl_result = process_sell(
                    stock={
                        'ticker':       stock['ticker'],
                        'stock_name':   stock.get('stock_name', stock['ticker']),
                        'industry':     stock.get('industry', 'N/A'),
                        'buying_price': stock['buying_price'],
                        'buying_date':  stock['buying_date'],
                        'qty':          stock['qty'],
                        'target_hit':   stock.get('target_hit', False),
                    },
                    selling_price=open_price
                )

                if pnl_result:
                    clear_pending_sell_flag(stock['ticker'], stock['buying_price'], stock['buying_date'])
                    print(f"[main] {stock['ticker']} sell executed at Rs.{open_price}")

            except Exception as e:
                print(f"[main] Error executing pending sell for {stock['ticker']}: {e}")

    except Exception as e:
        print(f"[main] execute_pending_sells error: {e}")


# ── MAIN ANALYSIS ─────────────────────────────────────────────────────────────

def run_analysis(test_mode=False, morning_job=False):

    if morning_job:
        execute_pending_sells()
        return []

    print("\n" + "="*60)
    print(f"  PORTFOLIO AI -- Starting at {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    print("="*60 + "\n")

    # Market status check (Rule 1.1)
    market_open_today = did_market_trade_today() if not test_mode else True

    # Market-independent tasks -- always run even on holidays
    if not test_mode:
        if should_send_weekly():
            print("Sending weekly portfolio summary...")
            send_weekly_summary()
        if should_send_monthly():
            print("Sending monthly portfolio summary...")
            send_monthly_summary()

    # Gate: stop here if market was closed
    if not market_open_today and not test_mode:
        print("[main] Market closed today -- skipping all market-dependent tasks.")
        print("="*60 + "\n")
        return []

    # Step 0: Split detection (Flaw 1.4)
    if not test_mode:
        print("Step 0: Checking for stock splits...")
        try:
            check_and_fix_stock_splits()
        except Exception as e:
            print(f"[main] Split check error (non-fatal): {e}")
        print()

    # Step 1: Read Holdings
    print("Step 1: Reading holdings from Google Sheets...")
    holdings = read_holdings()
    if test_mode:
        holdings = holdings[:2]
    print(f"   -> {len(holdings)} stocks loaded.\n")

    # Fear & Greed
    print("Fetching Fear & Greed Index...")
    fear_greed = get_fear_greed()
    print(f"   Market Mood: {fear_greed['score']} -- {fear_greed['rating']} {fear_greed['emoji']}\n")

    # Update sheet prices
    if not test_mode:
        print("Updating live prices in Sheets...")
        update_holdings_prices()
        update_pnl_prices()
        print()

    # Step 2: Fetch Live Prices
    print("Step 2: Fetching live prices from NSE...")
    enriched = enrich_holdings_with_prices(holdings)
    print()

    # Step 3: Technical Analysis (Flaw 1.3 -- data sufficiency checks inside calculate_indicators)
    print("Step 3: Running technical analysis...")
    for stock in enriched:
        tech = calculate_indicators(stock['ticker_yf'])
        stock.update(tech)
        insuff = tech.get('insufficient_indicators', [])
        note = f" | Skipped: {', '.join(insuff)}" if insuff else ""
        print(f"   {stock['ticker']}: RSI={stock.get('rsi')} | Signal={stock.get('technical_signal')}{note}")
    print()

    # Step 4: News + AI Sentiment (Flaw 2.1 -- 5-day memory, Flaw 2.2 -- tiered news, Flaw 2.3 -- earnings)
    print("Step 4: Fetching news & AI sentiment with 5-day memory...")
    for stock in enriched:

        sentiment_history = []
        if not test_mode:
            try:
                sentiment_history = get_sentiment_history(stock['ticker'], days=5)
            except Exception as e:
                print(f"   [warning] Could not fetch sentiment history for {stock['ticker']}: {e}")

        technical_data = {
            'rsi':                     stock.get('rsi'),
            'macd':                    stock.get('macd'),
            'macd_signal':             stock.get('macd_signal'),
            'adx':                     stock.get('adx'),
            'stoch_k':                 stock.get('stoch_k'),
            'ema50':                   stock.get('ema50'),
            'ema200':                  stock.get('ema200'),
            'bb_upper':                stock.get('bb_upper'),
            'bb_lower':                stock.get('bb_lower'),
            'bull_pct':                stock.get('bull_pct', 50),
            'technical_notes':         stock.get('technical_notes', []),
            'insufficient_indicators': stock.get('insufficient_indicators', []),
        }
        stock_data = {
            'buying_price':  stock.get('buying_price', 0),
            'current_price': stock.get('live_price', 0),
            'growth_pct':    stock.get('growth_pct', 0),
            'qty':           stock.get('qty', 0),
        }
        sentiment = get_full_sentiment(
            stock['stock_name'],
            stock['ticker'],
            technical_data,
            stock_data,
            sentiment_history=sentiment_history
        )
        stock.update(sentiment)

        # Store today's verdict for future 5-day memory
        if not test_mode:
            try:
                log_sentiment_history(
                    ticker=stock['ticker'],
                    verdict=sentiment.get('overall_sentiment', 'NEUTRAL'),
                    run_date=date.today().strftime('%Y-%m-%d')
                )
            except Exception as e:
                print(f"   [warning] Could not log sentiment for {stock['ticker']}: {e}")

        print(f"   {stock['ticker']}: News={sentiment.get('news_sentiment')} | AI={sentiment.get('overall_sentiment')} | Flip={sentiment.get('sentiment_flip','NO')}")
    print()

    # Step 5: Decision Engine
    # Attach sentiment history so engine can count consecutive BEARISH days (Flaw 2.1)
    print("Step 5: Running decision engine...")
    for stock in enriched:
        if not test_mode:
            try:
                stock['sentiment_history'] = get_sentiment_history(stock['ticker'], days=5)
            except:
                stock['sentiment_history'] = []
        else:
            stock['sentiment_history'] = []

    results = process_all_holdings(enriched)
    print()

    # Step 6: Report
    print("Step 6: Generating report...")
    html    = generate_html_report(results, fear_greed=fear_greed)
    subject = generate_subject_line(results)
    print(f"   Subject: {subject}")

    # Step 7: Send Email + Flag PENDING SELLs (Flaw 3.4)
    if not test_mode:
        print("Step 7: Sending email report...")
        send_report_email(subject, html)

        sell_stocks = [s for s in results if s.get('decision') == 'SELL']
        for s in sell_stocks:
            target_hit = 'TARGET PRICE HIT' in s.get('reason', '')
            print(f"[main] SELL decision for {s['ticker']} -- flagging as PENDING_SELL")
            try:
                flag_pending_sell(
                    ticker=s['ticker'],
                    decision_price=s['live_price'],
                    target_hit=target_hit
                )
            except Exception as e:
                # Fallback: execute immediately if flagging fails
                print(f"[main] Could not flag {s['ticker']} as PENDING_SELL: {e}")
                print(f"[main] Falling back to immediate sell at closing price...")
                process_sell(
                    stock={
                        'ticker':       s['ticker'],
                        'stock_name':   s['stock_name'],
                        'industry':     s['industry'],
                        'buying_price': s['buying_price'],
                        'buying_date':  s['buying_date'],
                        'qty':          s['qty'],
                        'target_hit':   target_hit,
                    },
                    selling_price=s['live_price']
                )
    else:
        print("   [TEST MODE] Email not sent.")
        with open("test_report.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("   Report saved to: test_report.html")

    # Step 8: Check YES/NO Approval Replies
    if not test_mode:
        print("Step 8: Checking Gmail for stock approval replies...")
        process_approvals()
        print()

    # Step 9: Scout New Growth Stocks + log recommendations (Flaw 2.4)
    if not test_mode:
        print("Step 9: Scouting new growth stocks...")
        candidates = find_growth_stocks(top_n=5)
        if candidates:
            for c in candidates:
                try:
                    log_recommendation(
                        ticker=c['ticker'],
                        recommended_price=c['current_price'],
                        target_price=c.get('ai_target_price'),
                        recommended_date=date.today().strftime('%Y-%m-%d')
                    )
                except Exception as e:
                    print(f"   [warning] Could not log recommendation for {c['ticker']}: {e}")
            send_scout_email(candidates)
            print(f"   -> Scout email sent with {len(candidates)} candidates.")
        else:
            print("   -> No strong candidates found today.")
        print()

    # Save dashboard cache
    if not test_mode:
        _save_dashboard_cache(results, fear_greed)

    print("\n" + "="*60)
    print(f"  DONE -- Analysis complete at {datetime.now().strftime('%H:%M:%S')}")
    print("="*60 + "\n")

    return results


def _save_dashboard_cache(results, fear_greed):
    """Save all data needed by index.html to last_analysis.json"""
    try:
        import yfinance as yf
        from weekly_monthly_summary import get_recommendation_accuracy

        idx_list = [
            {"symbol": "^NSEI",             "name": "NIFTY 50"},
            {"symbol": "^BSESN",            "name": "SENSEX"},
            {"symbol": "^NSEBANK",          "name": "BANK NIFTY"},
            {"symbol": "NIFTYMIDCAP150.NS", "name": "MIDCAP 150"},
        ]
        indices_data = []
        for idx in idx_list:
            try:
                info  = yf.Ticker(idx['symbol']).info
                price = info.get('regularMarketPrice') or info.get('currentPrice') or 0
                prev  = info.get('previousClose') or price
                chg   = round(price - prev, 2)
                pct   = round((chg / prev * 100), 2) if prev else 0
                indices_data.append({"name": idx['name'], "price": round(price,2), "change": chg, "change_pct": pct})
            except:
                indices_data.append({"name": idx['name'], "price": 0, "change": 0, "change_pct": 0})

        total_investment = sum(s.get('investment_amt') or 0 for s in results)
        total_current    = sum(s.get('current_value') or 0 for s in results)
        total_profit     = round(total_current - total_investment, 2)
        total_growth     = round((total_profit / total_investment * 100), 2) if total_investment else 0

        clean_stocks = []
        for s in results:
            clean_stocks.append({
                'ticker':                 s.get('ticker', ''),
                'stock_name':             s.get('stock_name', ''),
                'industry':               s.get('industry', ''),
                'sector':                 s.get('sector', ''),
                'cap_category':           s.get('cap_category', ''),
                'buying_price':           s.get('buying_price'),
                'buying_date':            str(s.get('buying_date', '')),
                'live_price':             s.get('live_price'),
                'growth_pct':             s.get('growth_pct'),
                'total_profit':           s.get('total_profit'),
                'investment_amt':         s.get('investment_amt'),
                'current_value':          s.get('current_value'),
                'qty':                    s.get('qty'),
                'decision':               s.get('decision', 'HOLD'),
                'urgency':                s.get('urgency', 'LOW'),
                'reason':                 s.get('reason', ''),
                'technical_signal':       s.get('technical_signal', 'NEUTRAL'),
                'rsi':                    s.get('rsi'),
                'overall_sentiment':      s.get('overall_sentiment', 'NEUTRAL'),
                'news_sentiment':         s.get('news_sentiment', 'NEUTRAL'),
                'ai_decision':            s.get('ai_decision', 'HOLD'),
                'ai_confidence':          s.get('ai_confidence', 'LOW'),
                'recommendation':         s.get('recommendation', ''),
                'risk_factors':           s.get('risk_factors', ''),
                'consecutive_bearish':    s.get('consecutive_bearish', 0),
                'stop_loss_threshold':    s.get('stop_loss_threshold'),
                'profit_threshold':       s.get('profit_threshold'),
                'pending_sell':           s.get('pending_sell', False),
                'nifty_return_since_buy': s.get('nifty_return_since_buy'),
                'sentiment_flip':         s.get('sentiment_flip', 'NO'),
            })

        rec_accuracy = None
        try:
            rec_accuracy = get_recommendation_accuracy()
        except Exception as e:
            print(f"[main] Could not fetch recommendation accuracy: {e}")

        cache = {
            'stocks':                  clean_stocks,
            'indices':                 indices_data,
            'fear_greed':              fear_greed,
            'total_investment':        round(total_investment, 2),
            'total_current':           round(total_current, 2),
            'total_profit':            total_profit,
            'total_growth':            total_growth,
            'last_updated':            datetime.now().strftime('%d-%m-%Y %H:%M'),
            'data_date':               date.today().strftime('%d %B %Y'),
            'recommendation_accuracy': rec_accuracy,
        }

        with open('last_analysis.json', 'w') as f:
            json.dump(cache, f, default=str)
        print("[main] Dashboard cache saved.")

    except Exception as e:
        print(f"[main] Cache save failed: {e}")


def run_scheduled():
    from apscheduler.schedulers.blocking import BlockingScheduler
    from config import RUN_HOUR, RUN_MINUTE
    scheduler = BlockingScheduler(timezone="Asia/Kolkata")
    scheduler.add_job(run_analysis, 'cron', hour=RUN_HOUR, minute=RUN_MINUTE)
    print(f"Scheduler started. Will run daily at {RUN_HOUR:02d}:{RUN_MINUTE:02d} IST")
    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        print("\nScheduler stopped.")


if __name__ == "__main__":
    if "--schedule" in sys.argv:
        run_scheduled()
    elif "--test" in sys.argv:
        run_analysis(test_mode=True)
    elif "--morning" in sys.argv:
        run_analysis(morning_job=True)
    else:
        run_analysis(test_mode=False)