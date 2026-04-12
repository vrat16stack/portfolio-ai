"""
technical_analysis.py
7 indicators for strong technical signals:
1. RSI — momentum / overbought / oversold
2. MACD — trend direction and momentum
3. Bollinger Bands — volatility and price extremes
4. ADX — trend strength
5. Stochastic — overbought / oversold confirmation
6. EMA Cross — short vs long term trend (50 EMA vs 200 EMA)
7. OBV — volume momentum

FLAW 1.3 FIX: Each indicator now checks minimum candle count before calculating.
If insufficient data exists, that indicator is skipped and marked in
insufficient_indicators list instead of being calculated on partial data.

Minimum candle requirements:
  RSI:             15 trading days
  MACD:            35 trading days
  Bollinger Bands: 21 trading days
  ADX:             28 trading days
  Stochastic:      17 trading days
  EMA 50:          50 trading days
  EMA 200:        200 trading days
  OBV:             10 trading days
"""

import ta
import pandas as pd
from price_fetcher import get_historical_data

# Minimum candle requirements per indicator
MIN_CANDLES = {
    'rsi':    15,
    'macd':   35,
    'bb':     21,
    'adx':    28,
    'stoch':  17,
    'ema50':  50,
    'ema200': 200,
    'obv':    10,
}


def _has_enough_data(df, indicator_name):
    """Returns True if df has enough rows for the given indicator"""
    required = MIN_CANDLES.get(indicator_name, 30)
    actual   = len(df)
    if actual < required:
        print(f"[technical] {indicator_name.upper()} skipped — only {actual} candles, need {required}")
        return False
    return True


def calculate_indicators(ticker_yf):
    df = get_historical_data(ticker_yf, days=300)

    # Absolute minimum — if fewer than 15 candles, return empty result
    if df is None or len(df) < 15:
        return {
            'rsi':                     None,
            'macd':                    None,
            'macd_signal':             None,
            'bb_upper':                None,
            'bb_lower':                None,
            'bb_mid':                  None,
            'adx':                     None,
            'stoch_k':                 None,
            'ema50':                   None,
            'ema200':                  None,
            'technical_signal':        'NEUTRAL',
            'technical_summary':       'Insufficient data — fewer than 15 candles available',
            'technical_notes':         ['Insufficient historical data for technical analysis'],
            'signal_score':            0,
            'bull_pct':                50,
            'insufficient_indicators': ['RSI', 'MACD', 'Bollinger Bands', 'ADX', 'Stochastic', 'EMA50', 'EMA200', 'OBV'],
            'candles_available':       len(df) if df is not None else 0,
        }

    close  = df['Close'].squeeze()
    high   = df['High'].squeeze()
    low    = df['Low'].squeeze()
    volume = df['Volume'].squeeze()

    current_price = round(float(close.iloc[-1]), 2)

    bullish      = 0
    bearish      = 0
    notes        = []
    insufficient = []   # FLAW 1.3: track which indicators were skipped

    # ── 1. RSI (needs 15 candles) ──────────────────────────────────────────────
    rsi = None
    if _has_enough_data(df, 'rsi'):
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
        except Exception as e:
            print(f"[technical] RSI error: {e}")
            rsi = None
    else:
        insufficient.append('RSI')

    # ── 2. MACD (needs 35 candles) ─────────────────────────────────────────────
    macd_val = macd_sig = None
    if _has_enough_data(df, 'macd'):
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
        except Exception as e:
            print(f"[technical] MACD error: {e}")
            macd_val = macd_sig = None
    else:
        insufficient.append('MACD')

    # ── 3. Bollinger Bands (needs 21 candles) ──────────────────────────────────
    bb_upper = bb_mid = bb_lower = None
    if _has_enough_data(df, 'bb'):
        try:
            bb_obj   = ta.volatility.BollingerBands(close, window=20, window_dev=2)
            bb_upper = round(float(bb_obj.bollinger_hband().iloc[-1]), 2)
            bb_mid   = round(float(bb_obj.bollinger_mavg().iloc[-1]), 2)
            bb_lower = round(float(bb_obj.bollinger_lband().iloc[-1]), 2)
            bb_pct   = round(float(bb_obj.bollinger_pband().iloc[-1]), 4)

            if bb_pct < 0.2:
                bullish += 1.5
                notes.append(f"Price near lower BB → Oversold bounce likely")
            elif bb_pct > 0.8:
                bearish += 1.5
                notes.append(f"Price near upper BB → Overbought pullback likely")
            else:
                notes.append(f"Price within BB bands (neutral zone)")
        except Exception as e:
            print(f"[technical] Bollinger Bands error: {e}")
            bb_upper = bb_mid = bb_lower = None
    else:
        insufficient.append('Bollinger Bands')

    # ── 4. ADX (needs 28 candles) ──────────────────────────────────────────────
    adx = None
    if _has_enough_data(df, 'adx'):
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
        except Exception as e:
            print(f"[technical] ADX error: {e}")
            adx = None
    else:
        insufficient.append('ADX')

    # ── 5. Stochastic (needs 17 candles) ───────────────────────────────────────
    stoch_k = None
    if _has_enough_data(df, 'stoch'):
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
        except Exception as e:
            print(f"[technical] Stochastic error: {e}")
            stoch_k = None
    else:
        insufficient.append('Stochastic')

    # ── 6. EMA Cross (50 vs 200) ───────────────────────────────────────────────
    ema50  = None
    ema200 = None
    ema50_ok  = _has_enough_data(df, 'ema50')
    ema200_ok = _has_enough_data(df, 'ema200')

    if ema50_ok:
        try:
            ema50 = round(float(ta.trend.EMAIndicator(close, window=50).ema_indicator().iloc[-1]), 2)
        except Exception as e:
            print(f"[technical] EMA50 error: {e}")
            ema50 = None

    if ema200_ok:
        try:
            ema200 = round(float(ta.trend.EMAIndicator(close, window=200).ema_indicator().iloc[-1]), 2)
        except Exception as e:
            print(f"[technical] EMA200 error: {e}")
            ema200 = None

    # EMA cross signal — only if both available
    if ema50 is not None and ema200 is not None:
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
    elif ema50 is not None and ema200 is None:
        # Only EMA50 available — partial signal
        if current_price > ema50:
            bullish += 1
            notes.append(f"Price above EMA50 {ema50} → Short-term bullish (EMA200 insufficient data)")
        else:
            bearish += 1
            notes.append(f"Price below EMA50 {ema50} → Short-term bearish (EMA200 insufficient data)")
        insufficient.append('EMA200')
    else:
        insufficient.append('EMA50')
        insufficient.append('EMA200')

    # ── 7. OBV (needs 10 candles) ──────────────────────────────────────────────
    if _has_enough_data(df, 'obv'):
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
        except Exception as e:
            print(f"[technical] OBV error: {e}")
    else:
        insufficient.append('OBV')

    # ── Final Signal ───────────────────────────────────────────────────────────
    total    = bullish + bearish
    bull_pct = (bullish / total * 100) if total > 0 else 50

    if bull_pct >= 65:
        signal = 'BULLISH'
    elif bull_pct <= 35:
        signal = 'BEARISH'
    else:
        signal = 'NEUTRAL'

    # FLAW 1.3: Add skipped indicators note to technical_notes
    if insufficient:
        notes.append(f"Skipped (insufficient data): {', '.join(insufficient)}")

    return {
        'rsi':                     rsi,
        'macd':                    macd_val,
        'macd_signal':             macd_sig,
        'bb_upper':                bb_upper,
        'bb_mid':                  bb_mid,
        'bb_lower':                bb_lower,
        'adx':                     adx,
        'stoch_k':                 stoch_k,
        'ema50':                   ema50,
        'ema200':                  ema200,
        'current_price':           current_price,
        'technical_signal':        signal,
        'bullish_score':           round(bullish, 1),
        'bearish_score':           round(bearish, 1),
        'bull_pct':                round(bull_pct, 1),
        'technical_notes':         notes,
        'technical_summary':       ' | '.join(notes[:4]),
        'signal_score':            round(bull_pct, 1),
        'insufficient_indicators': insufficient,   # FLAW 1.3: passed to Groq + email
        'candles_available':       len(df),
    }