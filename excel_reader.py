"""
excel_reader.py
Reads ALL stocks dynamically from Sheet1 — no config.py dependency.

Sheet1 columns:
  A=S.no, B=Industry, C=Ticker(DataType-skip), D=Stock(plain text=ticker),
  E=Buying Price, F=Buying Date, G=Qty, H onwards=formulas(skip)
"""

import pandas as pd
from config import EXCEL_FILE_PATH, HOLDINGS_SHEET, NSE_SUFFIX


def read_holdings():
    try:
        df = pd.read_excel(
            EXCEL_FILE_PATH,
            sheet_name=HOLDINGS_SHEET,
            header=0,
            dtype=str
        )

        holdings = []
        for i, row in df.iterrows():
            # Column A - S.no must be valid number
            sno = str(row.iloc[0]).strip()
            if not sno or sno == 'nan' or not sno.replace('.','').isdigit():
                continue

            # Column D - Stock (plain text ticker like RELIANCE, SBIN, LT)
            ticker = str(row.iloc[3]).strip().upper()
            if not ticker or ticker == 'NAN':
                continue

            # Column B - Industry
            industry = str(row.iloc[1]).strip()
            if industry == 'nan' or industry == 'NAN':
                industry = 'N/A'

            # Column E - Buying Price
            try:
                buying_price = float(str(row.iloc[4]).replace('₹','').replace(',','').strip())
            except:
                continue

            # Column F - Buying Date
            try:
                buying_date = str(row.iloc[5]).strip()[:10]
                if buying_date == 'nan':
                    buying_date = '2024-01-01'
            except:
                buying_date = '2024-01-01'

            # Column G - Qty
            try:
                qty = int(float(str(row.iloc[6]).replace(',','').strip()))
            except:
                continue

            # Stock name — use ticker as name if not available
            stock_name = ticker

            holdings.append({
                'ticker': ticker,
                'ticker_yf': ticker + NSE_SUFFIX,
                'stock_name': stock_name,
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
        print(f"{h['ticker']} | Buy: {h['buying_price']} | Qty: {h['qty']} | Date: {h['buying_date']}")