"""
excel_reader.py
Reads ALL stocks dynamically from Sheet1 — no config.py dependency.
On cloud (USE_GOOGLE_SHEETS=True), this file is not used.
"""

import pandas as pd
from config import NSE_SUFFIX

try:
    from config import EXCEL_FILE_PATH, HOLDINGS_SHEET
except ImportError:
    EXCEL_FILE_PATH = ""
    HOLDINGS_SHEET = "Sheet1"


def read_holdings():
    if not EXCEL_FILE_PATH:
        print("[excel_reader] No Excel path configured — use Google Sheets instead.")
        return []
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=HOLDINGS_SHEET, header=0, dtype=str)
        holdings = []
        for i, row in df.iterrows():
            sno = str(row.iloc[0]).strip()
            if not sno or sno == 'nan' or not sno.replace('.','').isdigit():
                continue
            ticker = str(row.iloc[3]).strip().upper()
            if not ticker or ticker == 'NAN':
                continue
            industry = str(row.iloc[1]).strip()
            if industry in ('nan', 'NAN'):
                industry = 'N/A'
            try:
                buying_price = float(str(row.iloc[4]).replace('₹','').replace(',','').strip())
            except:
                continue
            try:
                buying_date = str(row.iloc[5]).strip()[:10]
                if buying_date == 'nan':
                    buying_date = '2024-01-01'
            except:
                buying_date = '2024-01-01'
            try:
                qty = int(float(str(row.iloc[6]).replace(',','').strip()))
            except:
                continue
            holdings.append({
                'ticker': ticker,
                'ticker_yf': ticker + NSE_SUFFIX,
                'stock_name': ticker,
                'industry': industry,
                'buying_price': buying_price,
                'buying_date': buying_date,
                'qty': qty,
                'current_price': None,
                'growth_pct': None,
            })
        print(f"[excel_reader] Loaded {len(holdings)} holdings from Excel.")
        return holdings
    except Exception as e:
        print(f"[excel_reader] ERROR: {e}")
        return []


if __name__ == "__main__":
    holdings = read_holdings()
    for h in holdings:
        print(f"{h['ticker']} | Buy: {h['buying_price']} | Qty: {h['qty']}")
