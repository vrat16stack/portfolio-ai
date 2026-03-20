"""
technical_analysis.py
7 indicators for strong technical signals:
1. RSI — momentum / overbought / oversold
2. MACD — trend direction and momentum
3. Bollinger Bands — volatility and price extremes
4. ADX — trend strength (is the trend strong or weak?)
5. Stochastic — overbought / oversold confirmation
6. EMA Cross — short vs long term trend (50 EMA vs 200 EMA)
7. OBV — volume momentum (is smart money buying or selling?)
"""

import ta
import pandas as pd
from price_fetcher import get_historical_data


def calculate_indicators(ticker_yf):
    df = get_historical_data(ticker_yf, days=250)

    if df is None or len(df) < 50:
        return {
            'rsi': None,
            'macd': None,
            'macd_signal': None,
            'bb_upper': None,
            'bb_lower': None,
            'bb_mid': None,
            'adx': None,
            'stoch_k': None,
            'ema50': None,
            'ema200': None,
            'technical_signal': 'NEUTRAL',
            'technical_summary': 'Insufficient data',
            'technical_notes': [],
            'signal_score': 0,
        }

    close  = df['Close'].squeeze()
    high   = df['High'].squeeze()
    low    = df['Low'].squeeze()
    volume = df['Volume'].squeeze()

    current_price = round(float(close.iloc[-1]), 2)

    bullish = 0
    bearish = 0
    notes   = []

    # ── 1. RSI ────────────────────────────────────────────────
    try:
        rsi = round(float(ta.momentum.RSIIndicator(close, window=14).rsi().iloc[-1]), 2)
        if rsi < 35:
            bullish += 1.5
            notes.append(f"RSI {rsi} → Oversold (Strong Bullish)")
        elif rsi < 50:
            bullish += 0.5
            notes.append(f"RSI {rsi} → Neutral-Bullish")
        elif rsi < 65:
            bearish += 0.5
            notes.append(f"RSI {rsi} → Neutral-Bearish")
        else:
            bearish += 1.5
            notes.append(f"RSI {rsi} → Overbought (Strong Bearish)")
    except:
        rsi = None

    # ── 2. MACD ───────────────────────────────────────────────
    try:
        macd_obj  = ta.trend.MACD(close, window_fast=12, window_slow=26, window_sign=9)
        macd_val  = round(float(macd_obj.macd().iloc[-1]), 4)
        macd_sig  = round(float(macd_obj.macd_signal().iloc[-1]), 4)
        macd_hist = round(float(macd_obj.macd_diff().iloc[-1]), 4)
        prev_hist = round(float(macd_obj.macd_diff().iloc[-2]), 4)

        if macd_val > macd_sig and macd_hist > prev_hist:
            bullish += 2
            notes.append(f"MACD bullish crossover + increasing momentum")
        elif macd_val > macd_sig:
            bullish += 1
            notes.append(f"MACD above signal → Bullish")
        elif macd_val < macd_sig and macd_hist < prev_hist:
            bearish += 2
            notes.append(f"MACD bearish crossover + decreasing momentum")
        else:
            bearish += 1
            notes.append(f"MACD below signal → Bearish")
    except:
        macd_val = macd_sig = None

    # ── 3. Bollinger Bands ────────────────────────────────────
    try:
        bb_obj    = ta.volatility.BollingerBands(close, window=20, window_dev=2)
        bb_upper  = round(float(bb_obj.bollinger_hband().iloc[-1]), 2)
        bb_mid    = round(float(bb_obj.bollinger_mavg().iloc[-1]), 2)
        bb_lower  = round(float(bb_obj.bollinger_lband().iloc[-1]), 2)
        bb_pct    = round(float(bb_obj.bollinger_pband().iloc[-1]), 4)

        if bb_pct < 0.2:
            bullish += 1.5
            notes.append(f"Price near lower BB → Oversold bounce likely")
        elif bb_pct > 0.8:
            bearish += 1.5
            notes.append(f"Price near upper BB → Overbought pullback likely")
        else:
            notes.append(f"Price within BB bands (neutral zone)")
    except:
        bb_upper = bb_mid = bb_lower = None

    # ── 4. ADX — Trend Strength ───────────────────────────────
    try:
        adx_obj = ta.trend.ADXIndicator(high, low, close, window=14)
        adx     = round(float(adx_obj.adx().iloc[-1]), 2)
        adx_pos = round(float(adx_obj.adx_pos().iloc[-1]), 2)
        adx_neg = round(float(adx_obj.adx_neg().iloc[-1]), 2)

        if adx > 25:
            if adx_pos > adx_neg:
                bullish += 2
                notes.append(f"ADX {adx} → Strong uptrend confirmed")
            else:
                bearish += 2
                notes.append(f"ADX {adx} → Strong downtrend confirmed")
        else:
            notes.append(f"ADX {adx} → Weak trend (sideways market)")
    except:
        adx = None

    # ── 5. Stochastic ─────────────────────────────────────────
    try:
        stoch_obj = ta.momentum.StochasticOscillator(high, low, close, window=14, smooth_window=3)
        stoch_k   = round(float(stoch_obj.stoch().iloc[-1]), 2)
        stoch_d   = round(float(stoch_obj.stoch_signal().iloc[-1]), 2)

        if stoch_k < 20 and stoch_k > stoch_d:
            bullish += 1.5
            notes.append(f"Stochastic {stoch_k} → Oversold + bullish crossover")
        elif stoch_k < 20:
            bullish += 1
            notes.append(f"Stochastic {stoch_k} → Oversold zone")
        elif stoch_k > 80 and stoch_k < stoch_d:
            bearish += 1.5
            notes.append(f"Stochastic {stoch_k} → Overbought + bearish crossover")
        elif stoch_k > 80:
            bearish += 1
            notes.append(f"Stochastic {stoch_k} → Overbought zone")
        else:
            notes.append(f"Stochastic {stoch_k} → Neutral zone")
    except:
        stoch_k = None

    # ── 6. EMA Cross (50 vs 200) ──────────────────────────────
    try:
        ema50  = round(float(ta.trend.EMAIndicator(close, window=50).ema_indicator().iloc[-1]), 2)
        ema200 = round(float(ta.trend.EMAIndicator(close, window=200).ema_indicator().iloc[-1]), 2)

        if ema50 > ema200 and current_price > ema50:
            bullish += 2
            notes.append(f"Golden Cross: EMA50 {ema50} > EMA200 {ema200} → Strong uptrend")
        elif ema50 > ema200:
            bullish += 1
            notes.append(f"EMA50 above EMA200 → Bullish trend")
        elif ema50 < ema200 and current_price < ema50:
            bearish += 2
            notes.append(f"Death Cross: EMA50 {ema50} < EMA200 {ema200} → Strong downtrend")
        else:
            bearish += 1
            notes.append(f"EMA50 below EMA200 → Bearish trend")
    except:
        ema50 = ema200 = None

    # ── 7. OBV — Volume Momentum ──────────────────────────────
    try:
        obv        = ta.volume.OnBalanceVolumeIndicator(close, volume).on_balance_volume()
        obv_recent = obv.iloc[-1]
        obv_prev   = obv.iloc[-10]

        if obv_recent > obv_prev:
            bullish += 1
            notes.append(f"OBV rising → Smart money accumulating (Bullish)")
        else:
            bearish += 1
            notes.append(f"OBV falling → Smart money distributing (Bearish)")
    except:
        pass

    # ── Final Signal ──────────────────────────────────────────
    total = bullish + bearish
    bull_pct = (bullish / total * 100) if total > 0 else 50

    if bull_pct >= 65:
        signal = 'BULLISH'
    elif bull_pct <= 35:
        signal = 'BEARISH'
    else:
        signal = 'NEUTRAL'

    return {
        'rsi':               rsi,
        'macd':              macd_val,
        'macd_signal':       macd_sig,
        'bb_upper':          bb_upper,
        'bb_mid':            bb_mid,
        'bb_lower':          bb_lower,
        'adx':               adx,
        'stoch_k':           stoch_k,
        'ema50':             ema50,
        'ema200':            ema200,
        'current_price':     current_price,
        'technical_signal':  signal,
        'bullish_score':     round(bullish, 1),
        'bearish_score':     round(bearish, 1),
        'bull_pct':          round(bull_pct, 1),
        'technical_notes':   notes,
        'technical_summary': ' | '.join(notes[:4]),
        'signal_score':      round(bull_pct, 1),
    }