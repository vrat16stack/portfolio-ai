"""
price_fetcher.py
Fetches live/EOD prices and technical indicators for NSE stocks using yfinance.
"""

import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta


def get_live_price(ticker_yf):
    """Get current/last traded price for a stock"""
    try:
        stock = yf.Ticker(ticker_yf)
        info = stock.info
        price = info.get('currentPrice') or info.get('regularMarketPrice') or info.get('previousClose')
        return round(float(price), 2) if price else None
    except Exception as e:
        print(f"[price_fetcher] Error fetching price for {ticker_yf}: {e}")
        return None


def get_historical_data(ticker_yf, days=200):
    """Get historical OHLCV data for technical analysis"""
    try:
        end = datetime.now()
        start = end - timedelta(days=days)
        df = yf.download(ticker_yf, start=start, end=end, progress=False, auto_adjust=True)
        if df.empty:
            return None
        df.columns = [col[0] if isinstance(col, tuple) else col for col in df.columns]
        return df
    except Exception as e:
        print(f"[price_fetcher] Error fetching history for {ticker_yf}: {e}")
        return None


def get_stock_info(ticker_yf):
    """Get fundamental data from yfinance"""
    try:
        stock = yf.Ticker(ticker_yf)
        info = stock.info
        return {
            'pe_ratio': info.get('trailingPE'),
            'pb_ratio': info.get('priceToBook'),
            'market_cap': info.get('marketCap'),
            'revenue_growth': info.get('revenueGrowth'),
            'earnings_growth': info.get('earningsGrowth'),
            'debt_to_equity': info.get('debtToEquity'),
            'roe': info.get('returnOnEquity'),
            'sector': info.get('sector'),
            '52w_high': info.get('fiftyTwoWeekHigh'),
            '52w_low': info.get('fiftyTwoWeekLow'),
            'avg_volume': info.get('averageVolume'),
            'dividend_yield': info.get('dividendYield'),
        }
    except Exception as e:
        print(f"[price_fetcher] Error fetching info for {ticker_yf}: {e}")
        return {}


def enrich_holdings_with_prices(holdings):
    """Add live prices and calculated P&L to holdings list"""
    enriched = []
    for stock in holdings:
        ticker_yf = stock['ticker_yf']
        live_price = get_live_price(ticker_yf)

        if live_price is None:
            live_price = stock.get('current_price')  # fallback to Excel value

        buying_price = stock['buying_price']
        qty = stock['qty']

        if live_price and buying_price:
            profit_per_share = round(live_price - buying_price, 2)
            total_profit = round(profit_per_share * qty, 2)
            growth_pct = round(((live_price - buying_price) / buying_price) * 100, 2)
            investment_amt = round(buying_price * qty, 2)
            current_value = round(live_price * qty, 2)
        else:
            profit_per_share = total_profit = growth_pct = investment_amt = current_value = None

        enriched.append({
            **stock,
            'live_price': live_price,
            'profit_per_share': profit_per_share,
            'total_profit': total_profit,
            'growth_pct': growth_pct,
            'investment_amt': investment_amt,
            'current_value': current_value,
        })
        print(f"[price_fetcher] {stock['ticker']}: ₹{live_price} | Growth: {growth_pct}%")

    return enriched


if __name__ == "__main__":
    from excel_reader import read_holdings
    holdings = read_holdings()
    enriched = enrich_holdings_with_prices(holdings[:3])  # Test with first 3
    for s in enriched:
        print(s)
