"""
main.py  — PORTFOLIO AI MASTER SCRIPT
Run this file manually or via Windows Task Scheduler at 3:45 PM IST daily.

Pipeline:
  1. Read holdings from Google Sheets (Sheet1)
  2. Fetch live prices from NSE (yfinance)
  3. Run technical analysis (RSI, MACD, Bollinger Bands)
  4. Fetch news + AI sentiment (Google News + Gemini)
  5. Apply decision engine (BUY / SELL / HOLD rules)
  6. Generate HTML report
  7. Send email
  8. (If SELL) → Update P&L sheet (Sheet3) after confirmation

Usage:
  python main.py                  # Run immediately
  python main.py --schedule       # Run on schedule (3:45 PM daily)
  python main.py --test           # Test with 2 stocks, no email
"""

import sys
import time
from datetime import datetime

from config import USE_GOOGLE_SHEETS
from price_fetcher import enrich_holdings_with_prices
from technical_analysis import calculate_indicators
from news_sentiment import get_full_sentiment
from decision_engine import process_all_holdings
from report_generator import generate_html_report, generate_subject_line
from email_handler import send_report_email, send_sell_alert
from pnl_updater import process_sell
from stock_scout import find_growth_stocks, send_scout_email
from weekly_monthly_summary import should_send_weekly, should_send_monthly, send_weekly_summary, send_monthly_summary
from fear_greed import get_fear_greed
from approval_checker import process_approvals

if USE_GOOGLE_SHEETS:
    from sheets_handler import read_holdings, update_holdings_prices, update_pnl_prices
    print("[main] Using Google Sheets as data source")
else:
    from excel_reader import read_holdings
    from excel_updater import update_sheet1_prices as update_holdings_prices, update_sheet3_prices as update_pnl_prices
    print("[main] Using Excel as data source")


def run_analysis(test_mode=False):
    print("\n" + "="*60)
    print(f"  PORTFOLIO AI — Starting analysis at {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    print("="*60 + "\n")

    # ── STEP 1: Read Holdings ─────────────────────────────────
    print("📂 Step 1: Reading holdings from Excel...")
    holdings = read_holdings()
    if test_mode:
        holdings = holdings[:2]  # Only test with first 2 stocks
    print(f"   → {len(holdings)} stocks loaded.\n")

    # ── FEAR & GREED INDEX ───────────────────────────────────
    print("😱 Fetching Fear & Greed Index...")
    fear_greed = get_fear_greed()
    print(f"   Market Mood: {fear_greed['score']} — {fear_greed['rating']} {fear_greed['emoji']}")
    print()

    # ── UPDATE EXCEL PRICES ──────────────────────────────────
    if not test_mode:
        print("📊 Updating live prices...")
        update_holdings_prices()
        update_pnl_prices()
        print()

    # ── STEP 2: Fetch Live Prices ─────────────────────────────
    print("💰 Step 2: Fetching live prices from NSE...")
    enriched = enrich_holdings_with_prices(holdings)
    print()

    # ── STEP 3: Technical Analysis ────────────────────────────
    print("📈 Step 3: Running technical analysis...")
    for stock in enriched:
        tech = calculate_indicators(stock['ticker_yf'])
        stock.update(tech)
        print(f"   {stock['ticker']}: RSI={stock.get('rsi')} | MACD={stock.get('technical_signal')}")
    print()

    # ── STEP 4: News + AI Sentiment ───────────────────────────
    print("📰 Step 4: Fetching news & AI sentiment...")
    for stock in enriched:
        technical_data = {
            'rsi':             stock.get('rsi'),
            'macd':            stock.get('macd'),
            'macd_signal':     stock.get('macd_signal'),
            'adx':             stock.get('adx'),
            'stoch_k':         stock.get('stoch_k'),
            'ema50':           stock.get('ema50'),
            'ema200':          stock.get('ema200'),
            'bb_upper':        stock.get('bb_upper'),
            'bb_lower':        stock.get('bb_lower'),
            'bull_pct':        stock.get('bull_pct', 50),
            'technical_notes': stock.get('technical_notes', []),
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
            stock_data
        )
        stock.update(sentiment)
        print(f"   {stock['ticker']}: News={sentiment.get('news_sentiment')} | AI={sentiment.get('overall_sentiment')}")
    print()

    # ── STEP 5: Decision Engine ───────────────────────────────
    print("🧠 Step 5: Running decision engine...")
    results = process_all_holdings(enriched)
    print()

    # ── STEP 6: Report ───────────────────────────────────────
    print("📄 Step 6: Generating report...")
    html = generate_html_report(results, fear_greed=fear_greed)
    subject = generate_subject_line(results)
    print(f"   Subject: {subject}")

    # ── STEP 7: Send Email ────────────────────────────────────
    if not test_mode:
        print("📧 Step 7: Sending email report...")
        send_report_email(subject, html)

        # Auto-trigger Module 2 for all SELL decisions
        sell_stocks = [s for s in results if s.get('decision') == 'SELL']
        for s in sell_stocks:
            print(f"[main] SELL triggered for {s['ticker']} — updating Sheet3 and sending P&L...")
            process_sell(
                stock={
                    'ticker': s['ticker'],
                    'stock_name': s['stock_name'],
                    'industry': s['industry'],
                    'buying_price': s['buying_price'],
                    'buying_date': s['buying_date'],
                    'qty': s['qty'],
                },
                selling_price=s['live_price']
            )
    else:
        print("   [TEST MODE] Email not sent. Report generated successfully.")
        # Save report to file for inspection
        with open("test_report.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("   Report saved to: test_report.html")

    # ── STEP 8: Check for YES/NO Approval Replies ────────────
    if not test_mode:
        print("📬 Step 8: Checking Gmail for stock approval replies...")
        process_approvals()
        print()

    # ── STEP 9: Scout New Growth Stocks (daily) ─────────────
    if not test_mode:
        print("🔍 Step 9: Scouting new growth stocks...")
        candidates = find_growth_stocks(top_n=5)
        if candidates:
            send_scout_email(candidates)
            print(f"   → Scout email sent with {len(candidates)} candidates.")
        else:
            print("   → No strong candidates found today.")
        print()

    # ── Save analysis cache for dashboard ───────────────────
    if not test_mode:
        try:
            total_investment = sum(s.get('investment_amt') or 0 for s in results)
            total_current    = sum(s.get('current_value') or 0 for s in results)
            total_profit     = round(total_current - total_investment, 2)
            total_growth     = round((total_profit / total_investment * 100), 2) if total_investment else 0
            cache = {
                'stocks':           results,
                'total_investment': round(total_investment, 2),
                'total_current':    round(total_current, 2),
                'total_profit':     total_profit,
                'total_growth':     total_growth,
                'last_updated':     datetime.now().strftime('%d-%m-%Y %H:%M'),
            }
            import json
            with open('last_analysis.json', 'w') as f:
                json.dump(cache, f, default=str)
            print("[main] ✅ Dashboard cache updated")
        except Exception as e:
            print(f"[main] Cache save error: {e}")

    # ── Save dashboard cache ─────────────────────────────────
    if not test_mode:
        try:
            total_investment = sum(s.get('investment_amt') or 0 for s in results)
            total_current    = sum(s.get('current_value') or 0 for s in results)
            total_profit     = round(total_current - total_investment, 2)
            total_growth     = round((total_profit / total_investment * 100), 2) if total_investment else 0
            # Build clean serializable stock list with all AI fields
            clean_stocks = []
            for s in results:
                clean_stocks.append({
                    'ticker':            s.get('ticker', ''),
                    'stock_name':        s.get('stock_name', ''),
                    'industry':          s.get('industry', ''),
                    'buying_price':      s.get('buying_price'),
                    'live_price':        s.get('live_price'),
                    'growth_pct':        s.get('growth_pct'),
                    'total_profit':      s.get('total_profit'),
                    'investment_amt':    s.get('investment_amt'),
                    'current_value':     s.get('current_value'),
                    'decision':          s.get('decision', 'HOLD'),
                    'urgency':           s.get('urgency', 'LOW'),
                    'reason':            s.get('reason', ''),
                    'technical_signal':  s.get('technical_signal', 'NEUTRAL'),
                    'rsi':               s.get('rsi'),
                    'overall_sentiment': s.get('overall_sentiment', 'NEUTRAL'),
                    'news_sentiment':    s.get('news_sentiment', 'NEUTRAL'),
                    'ai_decision':       s.get('ai_decision', 'HOLD'),
                    'ai_confidence':     s.get('ai_confidence', 'LOW'),
                    'recommendation':    s.get('recommendation', ''),
                    'risk_factors':      s.get('risk_factors', ''),
                })

            cache = {
                'stocks':           clean_stocks,
                'total_investment': round(total_investment, 2),
                'total_current':    round(total_current, 2),
                'total_profit':     total_profit,
                'total_growth':     total_growth,
            }
            import json
            with open('last_analysis.json', 'w') as f:
                json.dump(cache, f, default=str)
            print("[main] Dashboard cache saved.")
        except Exception as e:
            print(f"[main] Cache save failed: {e}")

    # ── STEP 10: Weekly & Monthly Summary ────────────────────
    if not test_mode:
        if should_send_weekly():
            print("📅 Step 10: Sending weekly portfolio summary...")
            send_weekly_summary()
        elif should_send_monthly():
            print("📆 Step 10: Sending monthly portfolio summary...")
            send_monthly_summary()

    print("\n" + "="*60)
    print(f"  ✅ DONE — Analysis complete at {datetime.now().strftime('%H:%M:%S')}")
    print("="*60 + "\n")

    return results


def run_scheduled():
    """Run on APScheduler — every day at 3:45 PM"""
    from apscheduler.schedulers.blocking import BlockingScheduler
    from config import RUN_HOUR, RUN_MINUTE

    scheduler = BlockingScheduler(timezone="Asia/Kolkata")
    scheduler.add_job(run_analysis, 'cron', hour=RUN_HOUR, minute=RUN_MINUTE)
    print(f"⏰ Scheduler started. Will run daily at {RUN_HOUR:02d}:{RUN_MINUTE:02d} IST")
    print("   Press Ctrl+C to stop.\n")

    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        print("\n⏹ Scheduler stopped.")


if __name__ == "__main__":
    if "--schedule" in sys.argv:
        run_scheduled()
    elif "--test" in sys.argv:
        run_analysis(test_mode=True)
    else:
        run_analysis(test_mode=False)
