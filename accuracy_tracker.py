"""
accuracy_tracker.py
===================
Measures how the scout's recommendations actually performed.  Replaces the accuracy
functions in weekly_monthly_summary.py (those become unused).

What changed vs the old metric ("price today > price at recommendation"):
  * records are de-duplicated by (date, ticker); a ticker already seen is flagged is_repeat
  * headline numbers count UNIQUE tickers (first recommendation in the window) so one stock
    repeated five days does not count five times
  * each pick is walked day by day AFTER its recommendation date:
        Low  <= stop   -> STOP     (exit at min(stop, that day's open))
        High >= target -> TARGET   (exit at target)
        30 calendar days passed -> EXPIRED (exit at that day's close)
        same day stop AND target -> STOP (conservative)
        otherwise -> OPEN (marked at the latest close)
  * every pick is compared with Nifty 50 over the same dates (excess return)
  * ONE yfinance download per ticker (+1 for ^NSEI) instead of one .info call per pick

TRACK_STOP_PCT is a MEASUREMENT convention only. It is NOT the live sell rule.

Public API
    get_recommendation_accuracy(force=False)   -> dict | None  (memoised per day)
    compute_accuracy(picks, frames, nifty, sectors=None, today=None)   (pure, unit-testable)
    generate_accuracy_pdf(acc)                 -> bytes | None
    send_accuracy_email(acc)
    send_accuracy_report()                     (what the Saturday job calls)
"""

import html as _html
import io
from bisect import bisect_right
from datetime import date, datetime, timedelta

import pandas as pd

TRACK_STOP_PCT = 10.0     # tracking-only stop, % below the recommended price
LOOKBACK_DAYS  = 30       # recommendations younger than this are measured
EXPIRY_DAYS    = 30       # an open pick is closed out after this many calendar days
NIFTY_SYMBOL   = '^NSEI'

_CACHE         = {'key': None, 'value': None, 'ready': False}
_SECTOR_CACHE  = {}


# ── small helpers ─────────────────────────────────────────────────────────────

def _f(value, default=None):
    """Tolerant float parser ('1,234.5', '₹12', None, NaN)."""
    try:
        x = float(str(value).replace(',', '').replace('₹', '').strip())
        if x != x:
            return default
        return x
    except (TypeError, ValueError):
        return default


def _parse_date(value):
    try:
        return datetime.strptime(str(value)[:10], '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return None


def _flatten(df):
    """yfinance can return MultiIndex columns (Price, Ticker); keep the Price level."""
    if df is None:
        return None
    if getattr(df.columns, 'nlevels', 1) > 1:
        df = df.copy()
        df.columns = df.columns.get_level_values(0)
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]
    return df


def _normalise(frame):
    """Return a clean Open/High/Low/Close frame with a tz-naive DatetimeIndex, or None."""
    df = _flatten(frame)
    if df is None or len(df) == 0:
        return None
    need = ['Open', 'High', 'Low', 'Close']
    if not all(c in df.columns for c in need):
        return None
    df = df[need].dropna()
    if df.empty:
        return None
    idx = pd.to_datetime(df.index)
    if getattr(idx, 'tz', None) is not None:
        idx = idx.tz_localize(None)
    df = df.copy()
    df.index = idx
    return df.sort_index()


def _close_pairs(obj):
    """(dates, closes) lists from a Series or an OHLC DataFrame; ([], []) if unusable."""
    if obj is None:
        return [], []
    try:
        if hasattr(obj, 'columns'):
            df = _flatten(obj)
            if 'Close' not in df.columns:
                return [], []
            s = df['Close']
        else:
            s = obj
        s = s.dropna()
        idx = pd.to_datetime(s.index)
        if getattr(idx, 'tz', None) is not None:
            idx = idx.tz_localize(None)
        pairs = sorted(zip([ts.date() for ts in idx], [float(v) for v in s.values]))
        return [p[0] for p in pairs], [p[1] for p in pairs]
    except Exception:
        return [], []


def _nifty_return(dates, closes, d0, d1):
    """Nifty % change from the close on/before d0 to the close on/before d1."""
    if not dates:
        return None
    i0 = bisect_right(dates, d0) - 1
    i1 = bisect_right(dates, d1) - 1
    if i0 < 0 or i1 < 0 or closes[i0] <= 0:
        return None
    return round((closes[i1] / closes[i0] - 1) * 100, 2)


# ── step 1: turn log rows into picks ──────────────────────────────────────────

def load_picks(records, today=None, lookback=LOOKBACK_DAYS):
    """
    RecommendationsLog rows -> list of picks (oldest first).
    Drops unreadable rows, de-duplicates by (date, ticker), flags is_repeat.
    """
    today  = today or date.today()
    cutoff = today - timedelta(days=lookback)
    rows   = []
    for r in records or []:
        d = _parse_date(r.get('Date'))
        t = str(r.get('Ticker', '')).strip().upper()
        p = _f(r.get('Recommended Price'))
        if not d or not t or not p or p <= 0:
            continue
        if d < cutoff or d > today:
            continue
        target = _f(r.get('Target Price'), 0.0) or 0.0
        if target <= p:                 # a target at/below the entry price is meaningless
            target = 0.0
        rows.append((d, t, p, target, str(r.get('Sector', '') or '').strip()))
    rows.sort(key=lambda x: (x[0], x[1]))

    picks, seen_keys, seen_tickers = [], set(), set()
    for d, t, p, target, sector in rows:
        if (d, t) in seen_keys:
            continue
        seen_keys.add((d, t))
        picks.append({'date': d, 'ticker': t, 'rec_price': p, 'target': target,
                      'sector': sector, 'is_repeat': t in seen_tickers})
        seen_tickers.add(t)
    return picks


# ── step 2: walk one pick forward through daily candles ───────────────────────

def evaluate_pick(pick, df, today=None):
    today = today or date.today()
    d0, price, target = pick['date'], pick['rec_price'], pick['target']
    stop   = price * (1 - TRACK_STOP_PCT / 100.0)
    dates  = [ts.date() for ts in df.index]
    latest = float(df['Close'].iloc[-1])

    status, exit_price, exit_date = 'OPEN', latest, dates[-1]
    for day, (o, h, l, c) in zip(dates, df[['Open', 'High', 'Low', 'Close']].itertuples(index=False)):
        if day <= d0:
            continue
        if day > today:
            break
        if l <= stop:                                   # checked first: same-day both -> STOP
            status, exit_price, exit_date = 'STOP', min(stop, float(o)), day
            break
        if target > 0 and h >= target:
            status, exit_price, exit_date = 'TARGET', target, day
            break
        if (day - d0).days >= EXPIRY_DAYS:
            status, exit_price, exit_date = 'EXPIRED', float(c), day
            break
    return {'status': status, 'exit_price': exit_price, 'exit_date': exit_date,
            'current': latest, 'stop': stop}


# ── step 3: everything together (pure function) ───────────────────────────────

def compute_accuracy(picks, frames, nifty=None, sectors=None, today=None):
    """
    picks   : output of load_picks
    frames  : {ticker: OHLC DataFrame}
    nifty   : Nifty OHLC DataFrame or Close Series
    sectors : {ticker: sector} fallback when the log row has no sector
    """
    today   = today or date.today()
    sectors = sectors or {}
    n_dates, n_closes = _close_pairs(nifty)

    results, nodata = [], 0
    for p in picks:
        df = _normalise((frames or {}).get(p['ticker']))
        if df is None:
            nodata += 1
            continue
        ev       = evaluate_pick(p, df, today)
        price    = p['rec_price']
        ret      = round((ev['exit_price'] / price - 1) * 100, 2)
        cur_ret  = round((ev['current'] / price - 1) * 100, 2)
        n_ret    = _nifty_return(n_dates, n_closes, p['date'], ev['exit_date'])
        excess   = round(ret - n_ret, 2) if n_ret is not None else None
        results.append({
            'ticker':             p['ticker'],
            'sector':             p['sector'] or sectors.get(p['ticker']) or 'N/A',
            'date':               p['date'].strftime('%Y-%m-%d'),
            'rec_price':          round(price, 2),
            'current':            round(ev['current'], 2),
            'target':             round(p['target'], 2),
            'target_hit':         ev['status'] == 'TARGET',
            'status':             ev['status'],
            'exit_price':         round(ev['exit_price'], 2),
            'exit_date':          ev['exit_date'].strftime('%Y-%m-%d'),
            'return_pct':         ret,             # outcome return (exit vs rec price)
            'current_return_pct': cur_ret,         # latest close vs rec price
            'nifty_return_pct':   n_ret,
            'excess_pct':         excess,
            'is_repeat':          p['is_repeat'],
        })

    if not results:
        return None

    uniq = [r for r in results if not r['is_repeat']]
    total = len(uniq)
    profitable = sum(1 for r in uniq if r['current_return_pct'] > 0)     # old metric kept
    with_n = [r for r in uniq if r['excess_pct'] is not None]
    count = lambda s: sum(1 for r in uniq if r['status'] == s)
    return {
        'total':          total,
        'total_picks':    len(results),
        'profitable':     profitable,
        'accuracy_pct':   round(profitable / total * 100, 1) if total else 0.0,
        'n_target':       count('TARGET'),
        'n_stop':         count('STOP'),
        'n_open':         count('OPEN'),
        'n_expired':      count('EXPIRED'),
        'beat_nifty_pct': round(sum(1 for r in with_n if r['excess_pct'] > 0) / len(with_n) * 100, 1)
                          if with_n else None,
        'avg_return':     round(sum(r['return_pct'] for r in uniq) / total, 2) if total else 0.0,
        'avg_excess':     round(sum(r['excess_pct'] for r in with_n) / len(with_n), 2) if with_n else None,
        'no_data':        nodata,
        'generated':      today.strftime('%Y-%m-%d'),
        'results':        results,
    }


def unique_results(acc):
    return [r for r in acc['results'] if not r.get('is_repeat')]


# ── data access (yfinance / Sheets) ───────────────────────────────────────────

def _suffix():
    try:
        from config import NSE_SUFFIX
        return NSE_SUFFIX
    except Exception:
        return '.NS'


def _download(symbol, start, end):
    import yfinance as yf
    return yf.download(symbol, start=start.isoformat(), end=end.isoformat(),
                       interval='1d', progress=False, auto_adjust=False)


def _download_frames(picks, today):
    first = {}
    for p in picks:
        first[p['ticker']] = min(first.get(p['ticker'], p['date']), p['date'])
    frames = {}
    for ticker, d in first.items():
        try:
            frames[ticker] = _download(ticker + _suffix(), d - timedelta(days=5),
                                       today + timedelta(days=1))
        except Exception as e:
            print(f"[accuracy] Price download failed for {ticker}: {e}")
    return frames


def _lookup_sector(ticker):
    if ticker in _SECTOR_CACHE:
        return _SECTOR_CACHE[ticker]
    sector = 'N/A'
    try:
        import yfinance as yf
        sector = yf.Ticker(ticker + _suffix()).info.get('sector') or 'N/A'
    except Exception:
        pass
    _SECTOR_CACHE[ticker] = sector
    return sector


def get_recommendation_accuracy(force=False):
    """Accuracy stats for the last 30 days, or None. Memoised per calendar day."""
    key = date.today().isoformat()
    if not force and _CACHE['ready'] and _CACHE['key'] == key:
        return _CACHE['value']
    try:
        from sheets_handler import get_recommendations_log
        today = date.today()
        picks = load_picks(get_recommendations_log(), today=today)
        if not picks:
            acc = None
        else:
            frames  = _download_frames(picks, today)
            first   = min(p['date'] for p in picks)
            try:
                nifty = _download(NIFTY_SYMBOL, first - timedelta(days=7), today + timedelta(days=1))
            except Exception as e:
                print(f"[accuracy] Nifty download failed (excess return skipped): {e}")
                nifty = None
            sectors = {p['ticker']: _lookup_sector(p['ticker'])
                       for p in picks if not p['is_repeat'] and not p['sector']}
            acc = compute_accuracy(picks, frames, nifty, sectors, today)
        _CACHE.update({'key': key, 'value': acc, 'ready': True})
        return acc
    except Exception as e:
        print(f"[accuracy] Error calculating recommendation accuracy: {e}")
        return None


# ── PDF (plain text only: Helvetica cannot draw emoji or the rupee sign) ──────

def _signed(x, suffix='%'):
    return '-' if x is None else f"{'+' if x >= 0 else ''}{x:.1f}{suffix}"


_OUTCOME_TEXT = {'TARGET': 'Target hit', 'STOP': 'Stopped (-10%)',
                 'EXPIRED': 'Expired', 'OPEN': 'Open'}


def generate_accuracy_pdf(acc):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError:
        print("[accuracy] reportlab not installed - PDF skipped (add it to the YML pip line)")
        return None
    try:
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=1.5*cm,
                                leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
        styles = getSampleStyleSheet()
        title = ParagraphStyle('T', parent=styles['Title'], fontSize=16,
                               textColor=colors.HexColor('#1B3A6B'), spaceAfter=4)
        sub   = ParagraphStyle('S', parent=styles['Normal'], fontSize=10,
                               textColor=colors.HexColor('#4A5568'), spaceAfter=10)
        note  = ParagraphStyle('N', parent=styles['Normal'], fontSize=7,
                               textColor=colors.HexColor('#777777'))
        el = [Paragraph('Portfolio AI - Recommendation Accuracy Report', title)]
        beat = f"{acc['beat_nifty_pct']}%" if acc.get('beat_nifty_pct') is not None else 'n/a'
        el.append(Paragraph(
            f"Last {LOOKBACK_DAYS} days | Generated {acc['generated']} | "
            f"{acc['total']} unique tickers ({acc['total_picks']} picks incl. repeats) | "
            f"Profitable now: {acc['profitable']} ({acc['accuracy_pct']}%) | Beat Nifty: {beat}", sub))

        head = [['UNIQUE', 'PROFITABLE', 'TARGET', 'STOP', 'OPEN', 'EXPIRED', 'AVG RETURN', 'AVG VS NIFTY'],
                [acc['total'], acc['profitable'], acc['n_target'], acc['n_stop'], acc['n_open'],
                 acc['n_expired'], _signed(acc['avg_return']), _signed(acc.get('avg_excess'))]]
        ht = Table([[str(c) for c in row] for row in head], colWidths=[3.2*cm]*8)
        ht.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1B3A6B')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#F0F4F8')),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'), ('FONTSIZE', (0, 1), (-1, 1), 11),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC'))]))
        el += [ht, Spacer(1, 0.5*cm)]

        rows = [['#', 'DATE', 'TICKER', 'SECTOR', 'REC (Rs.)', 'CURRENT (Rs.)', 'TARGET (Rs.)',
                 'RETURN', 'VS NIFTY', 'OUTCOME']]
        ordered = sorted(acc['results'], key=lambda r: (r['date'], r['ticker']))
        for i, r in enumerate(ordered, 1):
            rows.append([str(i), r['date'], r['ticker'] + ('*' if r['is_repeat'] else ''),
                         str(r['sector'])[:22], f"{r['rec_price']:,.2f}", f"{r['current']:,.2f}",
                         f"{r['target']:,.2f}" if r['target'] > 0 else 'N/A',
                         _signed(r['return_pct']), _signed(r['excess_pct']),
                         _OUTCOME_TEXT.get(r['status'], r['status'])])
        widths = [0.8*cm, 2.2*cm, 2.6*cm, 4.2*cm, 2.5*cm, 2.7*cm, 2.6*cm, 2.1*cm, 2.2*cm, 2.9*cm]
        t = Table(rows, colWidths=widths, repeatRows=1)
        style = [('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
                 ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                 ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                 ('FONTSIZE', (0, 0), (-1, -1), 8), ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                 ('ALIGN', (3, 1), (3, -1), 'LEFT'),
                 ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#DDDDDD')),
                 ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7F9FC')]),
                 ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]
        for i, r in enumerate(ordered, 1):
            good = r['return_pct'] >= 0
            style.append(('TEXTCOLOR', (7, i), (7, i),
                          colors.HexColor('#2E7D32') if good else colors.HexColor('#C62828')))
            if r['status'] == 'TARGET':
                style.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#E8F5E9')))
        t.setStyle(TableStyle(style))
        el += [t, Spacer(1, 0.4*cm), Paragraph(
            "* = repeat recommendation of a ticker already listed earlier; headline numbers count each "
            f"ticker once. The {TRACK_STOP_PCT:.0f}% stop is a tracking convention for this report, not the "
            "live sell rule. RETURN = result at exit (target / stop / expiry) or latest close if still open. "
            "VS NIFTY = RETURN minus Nifty 50 over the same dates. Not financial advice.", note)]
        doc.build(el)
        return buf.getvalue()
    except Exception as e:
        print(f"[accuracy] PDF generation error: {e}")
        return None


# ── email ─────────────────────────────────────────────────────────────────────

def _send_html_with_pdf(subject, html_body, pdf_bytes, filename):
    """Own copy of the HTML-email-with-PDF helper (no cross-import)."""
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from config import EMAIL_PASSWORD, EMAIL_RECEIVER, EMAIL_SENDER
    try:
        msg = MIMEMultipart('mixed')
        msg['Subject'], msg['From'], msg['To'] = subject, EMAIL_SENDER, EMAIL_RECEIVER
        alt = MIMEMultipart('alternative')
        alt.attach(MIMEText(html_body, 'html', 'utf-8'))
        msg.attach(alt)
        part = MIMEApplication(pdf_bytes, _subtype='pdf')
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())
        print(f"[accuracy] Accuracy email with PDF sent to {EMAIL_RECEIVER}")
        return True
    except Exception as e:
        print(f"[accuracy] Could not send PDF email ({e}) - falling back to HTML only")
        from email_handler import send_report_email
        send_report_email(subject, html_body)
        return False


def build_accuracy_html(acc, has_pdf):
    esc = _html.escape
    uniq = unique_results(acc)
    a = acc['accuracy_pct']
    color = '#2E7D32' if a >= 60 else '#E65100' if a >= 50 else '#C62828'
    label = 'Good' if a >= 60 else 'Average' if a >= 50 else 'Needs improvement'
    beat = f"{acc['beat_nifty_pct']}%" if acc.get('beat_nifty_pct') is not None else 'n/a'

    sec = {}
    for r in uniq:
        s = sec.setdefault(r['sector'], {'n': 0, 'win': 0, 'tgt': 0, 'stop': 0})
        s['n'] += 1
        s['win'] += r['current_return_pct'] > 0
        s['tgt'] += r['status'] == 'TARGET'
        s['stop'] += r['status'] == 'STOP'
    sector_rows = ''.join(
        f"<tr><td style='padding:8px 12px;border-bottom:1px solid #eee;'>{esc(str(k))}</td>"
        f"<td style='padding:8px;text-align:center;border-bottom:1px solid #eee;'>{v['n']}</td>"
        f"<td style='padding:8px;text-align:center;border-bottom:1px solid #eee;color:#2E7D32;'>{v['win']}</td>"
        f"<td style='padding:8px;text-align:center;border-bottom:1px solid #eee;'>{v['tgt']}</td>"
        f"<td style='padding:8px;text-align:center;border-bottom:1px solid #eee;color:#C62828;'>{v['stop']}</td>"
        f"<td style='padding:8px;text-align:center;border-bottom:1px solid #eee;font-weight:bold;'>"
        f"{round(v['win'] / v['n'] * 100)}%</td></tr>"
        for k, v in sorted(sec.items(), key=lambda kv: kv[1]['n'], reverse=True))

    ranked = sorted(uniq, key=lambda r: r['return_pct'], reverse=True)

    def mini(picks):
        out = ''
        for r in picks:
            c = '#2E7D32' if r['return_pct'] >= 0 else '#C62828'
            out += (f"<tr><td style='padding:6px 10px;border-bottom:1px solid #eee;font-weight:bold;'>"
                    f"{esc(r['ticker'])}</td><td style='padding:6px 10px;border-bottom:1px solid #eee;"
                    f"font-size:11px;color:#666;'>{esc(str(r['sector'])[:15])}</td>"
                    f"<td style='padding:6px 10px;border-bottom:1px solid #eee;color:{c};font-weight:bold;'>"
                    f"{_signed(r['return_pct'])}</td><td style='padding:6px 10px;border-bottom:1px solid #eee;"
                    f"font-size:11px;color:#888;'>{_OUTCOME_TEXT.get(r['status'], r['status'])}</td></tr>")
        return out

    def card(title, value, bg='white', fg='#222'):
        return (f"<div style='flex:1;background:{bg};padding:14px;border-radius:10px;text-align:center;"
                f"min-width:110px;box-shadow:0 2px 6px rgba(0,0,0,0.08);'><div style='font-size:11px;"
                f"color:#888;text-transform:uppercase;margin-bottom:6px;'>{title}</div>"
                f"<div style='font-size:28px;font-weight:bold;color:{fg};'>{value}</div></div>")

    pdf_note = (f"<div style='background:#E3F2FD;border-left:4px solid #1565C0;padding:12px 16px;"
                f"border-radius:4px;margin-bottom:16px;font-size:13px;'><strong>Full list of all "
                f"{acc['total_picks']} picks is attached as a PDF.</strong></div>") if has_pdf else (
                "<div style='background:#FFF3E0;border-left:4px solid #E65100;padding:12px 16px;"
                "border-radius:4px;margin-bottom:16px;font-size:13px;'>PDF could not be generated this "
                "week (reportlab missing or an error) - check the run log.</div>")
    nodata = (f"<div style='font-size:12px;color:#E65100;margin-bottom:12px;'>{acc['no_data']} pick(s) "
              f"skipped: no price data.</div>") if acc.get('no_data') else ''
    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:820px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#0D47A1,#1565C0);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">AI Recommendation Accuracy Report</h1>
    <p style="margin:6px 0 0;opacity:0.85;">Last {LOOKBACK_DAYS} days | {acc['generated']} | {acc['total']} unique tickers ({acc['total_picks']} picks incl. repeats)</p>
  </div>
  <div style="display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap;">
    {card('Accuracy', f"{a}%", fg=color)}
    {card('Unique picks', acc['total'])}
    {card('Profitable now', acc['profitable'], '#E8F5E9', '#2E7D32')}
    {card('Targets hit', acc['n_target'], '#FFF8E1', '#E65100')}
    {card('Stopped', acc['n_stop'], '#FFEBEE', '#C62828')}
    {card('Beat Nifty', beat)}
  </div>
  <div style="background:white;border-radius:10px;padding:14px 16px;margin-bottom:20px;font-size:13px;box-shadow:0 2px 6px rgba(0,0,0,0.08);">
    Accuracy ({label}) = unique picks currently above their recommended price. Average outcome return
    {_signed(acc['avg_return'])}; average vs Nifty {_signed(acc.get('avg_excess'))}. Open: {acc['n_open']}, expired: {acc['n_expired']}.
  </div>
  {nodata}
  <div style="background:white;border-radius:10px;padding:16px;margin-bottom:20px;box-shadow:0 2px 6px rgba(0,0,0,0.08);">
    <h3 style="margin:0 0 12px;font-size:15px;color:#1565C0;">Accuracy by sector (unique picks)</h3>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr style="background:#1565C0;color:white;"><th style="padding:8px 12px;text-align:left;">Sector</th><th>Picks</th><th>Profitable</th><th>Target</th><th>Stop</th><th>Accuracy</th></tr>
      {sector_rows}
    </table>
  </div>
  <div style="display:flex;gap:16px;margin-bottom:20px;">
    <div style="flex:1;background:white;border-radius:10px;padding:16px;box-shadow:0 2px 6px rgba(0,0,0,0.08);">
      <h3 style="margin:0 0 10px;font-size:14px;color:#2E7D32;">Top 5 picks</h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">{mini(ranked[:5])}</table></div>
    <div style="flex:1;background:white;border-radius:10px;padding:16px;box-shadow:0 2px 6px rgba(0,0,0,0.08);">
      <h3 style="margin:0 0 10px;font-size:14px;color:#C62828;">Bottom 5 picks</h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">{mini(ranked[-5:][::-1])}</table></div>
  </div>
  {pdf_note}
  <div style="font-size:11px;color:#888;line-height:1.6;margin-bottom:12px;">
    The {TRACK_STOP_PCT:.0f}% stop used here is a measurement convention for this report - it is NOT your live sell rule.
    If a stop and the target were both touched on the same day, the pick is counted as stopped.
  </div>
  <div style="text-align:center;color:#aaa;font-size:11px;padding:12px;">Generated by Portfolio AI | Not financial advice</div>
</body></html>"""


def send_accuracy_email(acc):
    if not acc:
        print("[accuracy] No accuracy data - skipping accuracy email.")
        return
    pdf = generate_accuracy_pdf(acc)
    body = build_accuracy_html(acc, has_pdf=bool(pdf))
    today_str = datetime.now().strftime('%d %B %Y')
    subject = (f"AI Accuracy Report | {acc['accuracy_pct']}% | {acc['profitable']}/{acc['total']} "
               f"unique picks profitable | {acc['n_target']} targets hit | {today_str}")
    if pdf:
        _send_html_with_pdf(subject, body, pdf, f"AI_Accuracy_{today_str.replace(' ', '_')}.pdf")
    else:
        from email_handler import send_report_email
        send_report_email(subject, body)
    print("[accuracy] AI accuracy email sent.")


def send_accuracy_report():
    """What the Saturday job calls. Never raises."""
    try:
        send_accuracy_email(get_recommendation_accuracy(force=True))
    except Exception as e:
        print(f"[accuracy] send_accuracy_report failed: {e}")


if __name__ == "__main__":
    import sys
    result = get_recommendation_accuracy(force=True)
    if result is None:
        print("No accuracy data.")
    else:
        print({k: v for k, v in result.items() if k != 'results'})
        if '--send' in sys.argv:
            send_accuracy_email(result)
