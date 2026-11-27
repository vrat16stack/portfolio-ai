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

Reliability upgrade (this version):
  * A failed AI verdict (analysis_failed) is NEVER written to SentimentHistory, so it can no
    longer reset the consecutive-BEARISH streak.
  * flag_pending_sell is retried and its False return is treated as a failure (immediate-sell
    fallback + health issue). The Holdings sheet columns it needs are checked at start-up.
  * Morning job: robust open-price lookup, and a loud alert email if a pending sell is not executed.
  * Scout: recommendations are logged ONLY when the scout email can actually go out (after the
    15:30 IST gate), so the accuracy stats never include picks you were never shown.
  * Health email at the end of the run when anything went wrong (see HEALTH_EMAIL_ALWAYS).

Usage:
  python main.py              # Run immediately (3:45 PM job)
  python main.py --morning    # Morning job (9:20 AM) — execute pending sells at open price
  python main.py --schedule   # Run on APScheduler
  python main.py --test       # Test with 2 stocks, no email
"""

import sys
import json
import time
import html as _html
from datetime import datetime, date

from config import USE_GOOGLE_SHEETS
from price_fetcher import enrich_holdings_with_prices
from technical_analysis import calculate_indicators
from news_sentiment import get_full_sentiment
from decision_engine import process_all_holdings
from report_generator import generate_html_report, generate_subject_line
from email_handler import send_report_email
from pnl_updater import process_sell
from stock_scout import find_growth_stocks, send_scout_email, send_no_picks_email, is_after_market_close
from weekly_monthly_summary import should_send_weekly, should_send_monthly, send_weekly_summary, send_monthly_summary
from fear_greed import get_fear_greed
from approval_checker import process_approvals
from llm_client import get_stats as llm_get_stats, reset_stats as llm_reset_stats

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


# ── HEALTH TRACKING ───────────────────────────────────────────────────────────

HEALTH = {'issues': [], 'failed_verdicts': []}
HEALTH_EMAIL_ALWAYS = False     # True = send the health email every run, even when all is well
REQUIRED_SELL_COLUMNS = ['Pending Sell', 'Decision Price', 'Target Hit']


def _note_issue(message):
    print(f"[health] {message}")
    HEALTH['issues'].append(message)


def _send_alert(subject, lines):
    """Immediate alert email. Never raises."""
    try:
        items = ''.join(f"<li style='margin-bottom:6px;'>{_html.escape(str(l))}</li>" for l in lines)
        body = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:640px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:#B71C1C;color:white;padding:20px;border-radius:12px;margin-bottom:16px;">
    <h2 style="margin:0;">{_html.escape(subject)}</h2>
    <p style="margin:6px 0 0;opacity:0.85;font-size:13px;">{datetime.now().strftime('%d %B %Y %H:%M')}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:18px;font-size:14px;line-height:1.6;">
    <ul style="margin:0;padding-left:18px;">{items}</ul>
  </div>
</body></html>"""
        send_report_email(subject, body)
    except Exception as e:
        print(f"[main] Could not send alert email '{subject}': {e}")


def _send_health_email(stats, issues, failed_verdicts):
    """Sent only when something needs attention (or HEALTH_EMAIL_ALWAYS is True)."""
    try:
        problems = list(issues)
        if stats.get('failed'):
            problems.append(f"{stats['failed']} Groq call(s) failed on every model.")
        if stats.get('fallbacks'):
            problems.append(f"{stats['fallbacks']} call(s) were answered by a FALLBACK model "
                            f"(primary model failed) — check the Groq model list.")
        if failed_verdicts:
            problems.append(f"{len(failed_verdicts)} stock verdict(s) FAILED and were not stored: "
                            + '; '.join(v[:90] for v in failed_verdicts))
        if not problems and not HEALTH_EMAIL_ALWAYS:
            print("[health] All good — no health email needed.")
            return
        ok = not problems
        rows = ''.join(f"<li style='margin-bottom:6px;'>{_html.escape(p)}</li>" for p in problems) \
            or "<li>No issues.</li>"
        models = ', '.join(f"{m}: {n}" for m, n in stats.get('models', {}).items()) or 'none'
        errs = ''.join(f"<li style='font-size:12px;color:#666;'>{_html.escape(e)}</li>"
                       for e in stats.get('errors', [])[-5:])
        banner = '#2E7D32' if ok else '#E65100'
        body = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:680px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:{banner};color:white;padding:20px;border-radius:12px;margin-bottom:16px;">
    <h2 style="margin:0;">Portfolio AI health report</h2>
    <p style="margin:6px 0 0;opacity:0.9;font-size:13px;">{datetime.now().strftime('%d %B %Y %H:%M')}</p>
  </div>
  <div style="background:white;border-radius:10px;padding:18px;margin-bottom:14px;">
    <h3 style="margin:0 0 10px;font-size:15px;">Needs attention</h3>
    <ul style="margin:0;padding-left:18px;font-size:14px;line-height:1.6;">{rows}</ul>
  </div>
  <div style="background:white;border-radius:10px;padding:18px;font-size:13px;line-height:1.7;">
    <h3 style="margin:0 0 10px;font-size:15px;">Groq usage this run</h3>
    Calls: {stats.get('calls', 0)} | OK: {stats.get('ok', 0)} | Failed: {stats.get('failed', 0)} |
    429 retries: {stats.get('retries_429', 0)} | Fallbacks: {stats.get('fallbacks', 0)}<br>
    Models used: {_html.escape(models)}
    <ul style="margin:8px 0 0;padding-left:18px;">{errs}</ul>
  </div>
</body></html>"""
        send_report_email(("⚠️ Portfolio AI health: attention needed" if not ok
                           else "✅ Portfolio AI health: all good"), body)
    except Exception as e:
        print(f"[health] Could not send health email: {e}")


def _check_sell_columns():
    """Holdings must have the PENDING SELL columns, otherwise sells cannot be flagged."""
    try:
        from sheets_handler import get_sheets
        holdings_sheet, _ = get_sheets()
        headers = [h.strip() for h in holdings_sheet.row_values(1)]
        missing = [c for c in REQUIRED_SELL_COLUMNS if c not in headers]
        if missing:
            _note_issue(f"Holdings sheet is MISSING column(s) {missing} — SELL decisions cannot be "
                        f"flagged for the morning job (falling back to immediate sell at the close). "
                        f"Add these headers to row 1 of the Holdings tab.")
    except Exception as e:
        _note_issue(f"Could not verify Holdings columns ({e}).")


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

def _flatten_columns(df):
    """Newer yfinance returns MultiIndex columns (Price, Ticker); keep only the Price level."""
    if df is None:
        return None
    if getattr(df.columns, 'nlevels', 1) > 1:
        df = df.copy()
        df.columns = df.columns.get_level_values(0)
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]
    return df


def _get_open_price(ticker_yf):
    """Today's opening price: first 1-minute candle's Open, else the daily candle's Open. None if unknown."""
    import yfinance as yf
    today = date.today()

    try:
        hist = _flatten_columns(yf.download(ticker_yf, period="1d", interval="1m", progress=False))
        if hist is not None and not hist.empty and 'Open' in hist.columns:
            mask = [ts.date() == today for ts in hist.index]
            todays = hist[mask]
            if not todays.empty:
                value = float(todays['Open'].iloc[0])
                if value > 0:
                    return round(value, 2)
    except Exception as e:
        print(f"[main] 1-minute open lookup failed for {ticker_yf}: {e}")

    try:
        daily = _flatten_columns(yf.Ticker(ticker_yf).history(period="5d", interval="1d"))
        if daily is not None and not daily.empty and 'Open' in daily.columns:
            mask = [ts.date() == today for ts in daily.index]
            todays = daily[mask]
            if not todays.empty:
                value = float(todays['Open'].iloc[0])
                if value > 0:
                    return round(value, 2)
    except Exception as e:
        print(f"[main] Daily open lookup failed for {ticker_yf}: {e}")
    return None


def execute_pending_sells():
    """
    Runs at 9:20 AM IST via --morning flag.
    Checks Google Sheets for PENDING_SELL flags set yesterday at 3:45 PM.
    Fetches today's opening price and records it as the actual sell price.
    A sale that cannot be executed keeps its PENDING_SELL flag (it retries next morning)
    and triggers an alert email.
    """
    print("\n[main] Morning job: Executing pending sells at open price...")
    failures, notes = [], []

    try:
        pending = get_pending_sells()
        if not pending:
            print("[main] No pending sells to execute.")
            return

        for stock in pending:
            ticker = stock['ticker']
            try:
                open_price = _get_open_price(ticker + ".NS")
                if not open_price:
                    open_price = stock.get('pending_sell_price') or None
                    if open_price:
                        notes.append(f"{ticker}: no open price available — sold at yesterday's decision "
                                     f"price Rs.{open_price}.")
                    else:
                        failures.append(f"{ticker}: no open price AND no decision price — NOT sold, "
                                        f"flag kept, will retry next morning.")
                        continue

                print(f"[main] {ticker} -- sell price: Rs.{open_price}")

                pnl_result = process_sell(
                    stock={
                        'ticker':       ticker,
                        'stock_name':   stock.get('stock_name', ticker),
                        'industry':     stock.get('industry', 'N/A'),
                        'buying_price': stock['buying_price'],
                        'buying_date':  stock['buying_date'],
                        'qty':          stock['qty'],
                        'target_hit':   stock.get('target_hit', False),
                    },
                    selling_price=open_price
                )

                if pnl_result:
                    clear_pending_sell_flag(ticker, stock['buying_price'], stock['buying_date'])
                    print(f"[main] {ticker} sell executed at Rs.{open_price}")
                else:
                    failures.append(f"{ticker}: process_sell returned nothing (P&L record failed) — "
                                    f"holding and PENDING flag kept, will retry next morning.")

            except Exception as e:
                failures.append(f"{ticker}: error while selling — {e}")

    except Exception as e:
        failures.append(f"execute_pending_sells crashed: {e}")

    if failures:
        _send_alert("⚠️ Pending sell NOT executed", failures + notes)
    elif notes:
        _send_alert("ℹ️ Pending sell executed at decision price", notes)


def _flag_sell_with_retry(ticker, decision_price, target_hit):
    """flag_pending_sell up to 2 times (5 s apart). A False return counts as a failure."""
    for attempt in (1, 2):
        try:
            if flag_pending_sell(ticker=ticker, decision_price=decision_price, target_hit=target_hit):
                return True
            print(f"[main] flag_pending_sell({ticker}) returned False (attempt {attempt}/2)")
        except Exception as e:
            print(f"[main] flag_pending_sell({ticker}) raised {e} (attempt {attempt}/2)")
        if attempt == 1:
            time.sleep(5)
    return False


# ── MAIN ANALYSIS ─────────────────────────────────────────────────────────────

def run_analysis(test_mode=False, morning_job=False):

    if morning_job:
        # Part 1: Execute pending sells at open price
        execute_pending_sells()

        # Part 2: Execute pending buys with gap decision tree
        try:
            from pending_buys_handler import execute_pending_buys
            execute_pending_buys()
        except Exception as e:
            print(f"[main] Pending buys execution error: {e}")
            _send_alert("⚠️ Pending buys execution failed", [str(e)])

        return []

    HEALTH['issues'].clear()
    HEALTH['failed_verdicts'].clear()
    llm_reset_stats()

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

    if not test_mode and USE_GOOGLE_SHEETS:
        _check_sell_columns()

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

        if sentiment.get('analysis_failed'):
            # A failed call must NOT be stored: it would look like a NEUTRAL day and reset the BEARISH streak.
            HEALTH['failed_verdicts'].append(f"{stock['ticker']}: {sentiment.get('recommendation', 'analysis failed')}")
            print(f"   [warning] {stock['ticker']}: AI analysis FAILED — verdict not stored.")
        elif not test_mode:
            # Store today's verdict for future 5-day memory
            try:
                log_sentiment_history(
                    ticker=stock['ticker'],
                    verdict=sentiment.get('overall_sentiment', 'NEUTRAL'),
                    run_date=date.today().strftime('%Y-%m-%d')
                )
            except Exception as e:
                print(f"   [warning] Could not log sentiment for {stock['ticker']}: {e}")

        print(f"   {stock['ticker']}: News={sentiment.get('news_sentiment')} | AI={sentiment.get('overall_sentiment')} | Flip={sentiment.get('sentiment_flip','NO')}"
              + (" | FAILED" if sentiment.get('analysis_failed') else ""))
    print()

    # Step 5: Decision Engine
    # Attach sentiment history so engine can count consecutive BEARISH days (Flaw 2.1)
    print("Step 5: Running decision engine...")
    for stock in enriched:
        if not test_mode:
            try:
                stock['sentiment_history'] = get_sentiment_history(stock['ticker'], days=5)
            except Exception:
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
            if _flag_sell_with_retry(s['ticker'], s['live_price'], target_hit):
                continue

            # Fallback: execute immediately if flagging fails
            _note_issue(f"{s['ticker']}: could not flag PENDING_SELL — selling immediately at the closing price.")
            try:
                sold = process_sell(
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
                if not sold:
                    _note_issue(f"{s['ticker']}: immediate-sell fallback ALSO failed — SELL NOT RECORDED. Act manually.")
            except Exception as e:
                _note_issue(f"{s['ticker']}: immediate-sell fallback raised {e} — SELL NOT RECORDED. Act manually.")
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
        after_close = is_after_market_close()

        if not after_close:
            # The scout email is withheld before 15:30 IST, so these picks must NOT be logged:
            # the accuracy report would otherwise count picks you were never shown.
            _note_issue("Scout ran before the 15:30 IST gate — scout email withheld and "
                        f"{len(candidates or [])} pick(s) NOT logged. Re-run the job after 15:30 IST.")
        elif candidates:
            try:
                send_scout_email(candidates)
                print(f"   -> Scout email sent with {len(candidates)} candidates.")
                for c in candidates:
                    try:
                        try:
                            log_recommendation(
                                ticker=c['ticker'],
                                recommended_price=c['current_price'],
                                target_price=c.get('ai_target_price'),
                                recommended_date=date.today().strftime('%Y-%m-%d'),
                                sector=c.get('sector')
                            )
                        except TypeError:
                            # sheets_handler.py not patched yet (old signature has no sector)
                            log_recommendation(
                                ticker=c['ticker'],
                                recommended_price=c['current_price'],
                                target_price=c.get('ai_target_price'),
                                recommended_date=date.today().strftime('%Y-%m-%d')
                            )
                    except Exception as e:
                        print(f"   [warning] Could not log recommendation for {c['ticker']}: {e}")
            except Exception as e:
                _note_issue(f"Scout email failed ({e}) — picks NOT logged.")
        else:
            send_no_picks_email()
            print("   -> No strong candidates found today. Notification email sent.")
        print()

    # Save dashboard cache
    if not test_mode:
        _save_dashboard_cache(results, fear_greed)

    # Health email (only when something needs attention)
    if not test_mode:
        _send_health_email(llm_get_stats(), HEALTH['issues'], HEALTH['failed_verdicts'])

    print("\n" + "="*60)
    print(f"  DONE -- Analysis complete at {datetime.now().strftime('%H:%M:%S')}")
    print("="*60 + "\n")

    return results


def _save_dashboard_cache(results, fear_greed):
    """Save all data needed by index.html to last_analysis.json"""
    try:
        import yfinance as yf
        from accuracy_tracker import get_recommendation_accuracy

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
            except Exception:
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
                'analysis_failed':        s.get('analysis_failed', False),
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
