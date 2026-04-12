"""
pnl_updater.py
When a stock is sold:
  - On cloud (USE_GOOGLE_SHEETS=True): Updates Google Sheets + sends P&L email
  - On local (USE_GOOGLE_SHEETS=False): Updates Excel + sends P&L email

FLAW 3.4 FIX: target_hit is now read from stock dict (where it is actually stored),
not from pnl dict (where it was never set — existing bug fixed).
"""

import os
import time
from datetime import datetime
import yfinance as yf
from config import NSE_SUFFIX
from email_handler import send_report_email

try:
    from config import EXCEL_FILE_PATH
except ImportError:
    EXCEL_FILE_PATH = ""


def get_current_price(ticker):
    try:
        stock = yf.Ticker(ticker + NSE_SUFFIX)
        info  = stock.info
        price = info.get('currentPrice') or info.get('regularMarketPrice')
        return round(float(price), 2) if price else None
    except:
        return None


def process_sell(stock, selling_price, selling_date=None):
    from config import USE_GOOGLE_SHEETS
    if USE_GOOGLE_SHEETS:
        return _process_sell_sheets(stock, selling_price, selling_date)
    return _process_sell_excel(stock, selling_price, selling_date)


def _process_sell_sheets(stock, selling_price, selling_date=None):
    from sheets_handler import add_to_pnl, remove_stock_from_holdings
    if selling_date is None:
        selling_date = datetime.now().strftime('%Y-%m-%d')

    buying_price  = float(stock['buying_price'])
    qty           = int(stock['qty'])
    buying_date   = str(stock['buying_date'])[:10]
    selling_price = float(selling_price)

    investment_amt   = round(buying_price * qty, 2)
    profit_per_share = round(selling_price - buying_price, 2)
    total_profit     = round(profit_per_share * qty, 2)
    return_pct       = round(((selling_price - buying_price) / buying_price) * 100, 2)

    try:
        bd = datetime.strptime(buying_date, "%Y-%m-%d")
        sd = datetime.strptime(selling_date, "%Y-%m-%d")
        investment_days = (sd - bd).days
        time_months     = round(investment_days / 30.44, 1)
    except:
        investment_days = 0
        time_months     = 0

    current_price = get_current_price(stock["ticker"])
    current_return = round(((current_price - selling_price) / selling_price) * 100, 2) if current_price else 0.0
    if not current_price:
        current_price = selling_price

    add_to_pnl(stock, selling_price, selling_date)
    remove_stock_from_holdings(stock["ticker"], buying_price, buying_date)

    pnl = {
        "investment_amt":   investment_amt,
        "profit_per_share": profit_per_share,
        "total_profit":     total_profit,
        "return_pct":       return_pct,
        "investment_days":  investment_days,
        "time_months":      time_months,
        "current_price":    current_price,
        "current_return":   current_return,
        "selling_price":    selling_price,
        "buying_price":     buying_price,
        "qty":              qty,
        "buying_date":      buying_date,
        "selling_date":     selling_date,
        "sno":              1,
    }

    # FLAW 3.4 FIX: target_hit comes from stock dict, not pnl dict
    target_hit = stock.get('target_hit', False)
    send_pnl_email(stock, pnl, target_hit=target_hit)
    return pnl


def _process_sell_excel(stock, selling_price, selling_date=None):
    if selling_date is None:
        selling_date = datetime.now().strftime('%Y-%m-%d')

    buying_price  = float(stock['buying_price'])
    qty           = int(stock['qty'])
    buying_date   = str(stock['buying_date'])[:10]
    selling_price = float(selling_price)

    investment_amt   = round(buying_price * qty, 2)
    profit_per_share = round(selling_price - buying_price, 2)
    total_profit     = round(profit_per_share * qty, 2)
    return_pct       = round(((selling_price - buying_price) / buying_price) * 100, 2)

    try:
        bd = datetime.strptime(buying_date, '%Y-%m-%d')
        sd = datetime.strptime(selling_date, '%Y-%m-%d')
        investment_days = (sd - bd).days
        time_months     = round(investment_days / 30.44, 1)
    except:
        investment_days = 0
        time_months     = 0

    current_price = get_current_price(stock['ticker'])
    if current_price and selling_price:
        current_return = round(((current_price - selling_price) / selling_price) * 100, 2)
    else:
        current_price  = selling_price
        current_return = 0.0

    pnl = {
        'investment_amt':   investment_amt,
        'profit_per_share': profit_per_share,
        'total_profit':     total_profit,
        'return_pct':       return_pct,
        'investment_days':  investment_days,
        'time_months':      time_months,
        'current_price':    current_price,
        'current_return':   current_return,
        'selling_price':    selling_price,
        'buying_price':     buying_price,
        'qty':              qty,
        'buying_date':      buying_date,
        'selling_date':     selling_date,
    }

    for attempt in range(3):
        try:
            sno = _excel_update(stock, buying_price, qty, buying_date, selling_date, selling_price)
            pnl['sno'] = sno
            # FLAW 3.4 FIX: target_hit from stock dict
            target_hit = stock.get('target_hit', False)
            send_pnl_email(stock, pnl, target_hit=target_hit)
            return pnl
        except Exception as e:
            if attempt < 2:
                print(f"[pnl_updater] Attempt {attempt+1} failed: {e} — retrying in 5s...")
                time.sleep(5)
            else:
                print(f"[pnl_updater] ERROR: {e}")
                return None


def _excel_update(stock, buying_price, qty, buying_date, selling_date, selling_price):
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise Exception("win32com not available — use Google Sheets mode")

    pythoncom.CoInitialize()
    time.sleep(2)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb     = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE_PATH))
    sheet3 = wb.Sheets("Sheet3")
    sheet1 = wb.Sheets("Sheet1")

    RIGHT = -4152
    LEFT  = -4131

    next_row = 2
    while sheet3.Cells(next_row, 1).Value is not None:
        next_row += 1
    sno = next_row - 1

    def write_cell(col, val, align=RIGHT):
        cell = sheet3.Cells(next_row, col)
        cell.Value = val
        cell.Font.Name = "Calibri"
        cell.Font.Size = 11
        cell.Font.Color = 0x000000
        cell.HorizontalAlignment = align

    def write_formula(col, formula):
        cell = sheet3.Cells(next_row, col)
        cell.Formula = formula
        cell.Font.Name = "Calibri"
        cell.Font.Size = 11
        cell.Font.Color = 0x000000
        cell.HorizontalAlignment = RIGHT

    write_cell(1, int(sno),                         RIGHT)
    write_cell(2, str(stock.get('industry', '')),   LEFT)
    write_cell(3, str(stock.get('ticker', '')),     LEFT)
    write_cell(4, str(stock.get('stock_name', '')), LEFT)
    write_cell(5, buying_date,                      RIGHT)
    write_cell(6, selling_date,                     RIGHT)
    write_cell(7, buying_price,                     RIGHT)
    write_cell(8, selling_price,                    RIGHT)
    write_cell(9, int(qty),                         RIGHT)
    sheet3.Cells(next_row, 7).NumberFormat = "#,##0.00"
    sheet3.Cells(next_row, 8).NumberFormat = "#,##0.00"
    sheet3.Cells(next_row, 5).NumberFormat = "DD-MM-YYYY"
    sheet3.Cells(next_row, 6).NumberFormat = "DD-MM-YYYY"

    write_formula(10, f"=G{next_row}*I{next_row}")
    write_formula(11, f"=H{next_row}-G{next_row}")
    write_formula(12, f"=K{next_row}*I{next_row}")
    write_formula(13, f"=(K{next_row}/G{next_row})*100")
    write_formula(14, f'=DATEDIF(E{next_row},F{next_row},"d")')
    write_formula(15, f"=C{next_row}.Price")
    write_formula(16, f"=((O{next_row}-H{next_row})/H{next_row})*100")
    write_formula(17, f'=DATEDIF(F{next_row},TODAY(),"m")')

    print(f"[pnl_updater] Added {stock['ticker']} to Sheet3 at row {next_row}")

    last_row    = sheet1.UsedRange.Rows.Count
    deleted_row = None
    for row in range(2, last_row + 1):
        price_val = sheet1.Cells(row, 5).Value
        date_val  = str(sheet1.Cells(row, 6).Value or '')
        try:
            price_match = price_val and abs(float(price_val) - buying_price) < 0.01
        except:
            price_match = False
        if price_match and buying_date[:7] in date_val:
            deleted_row = row
            sheet1.Rows(row).Delete()
            print(f"[pnl_updater] Removed {stock['ticker']} from Sheet1")
            break

    if deleted_row:
        last_row = sheet1.UsedRange.Rows.Count
        for row in range(2, last_row + 1):
            if sheet1.Cells(row, 1).Value is not None:
                sheet1.Cells(row, 1).Value = row - 1
        print(f"[pnl_updater] Sheet1 S.no renumbered")

    wb.Save()
    try:
        wb.Close(False)
    except:
        pass
    try:
        excel.Quit()
    except:
        pass
    return sno


def send_pnl_email(stock, pnl, target_hit=False):
    is_profit    = pnl['total_profit'] >= 0
    profit_color = '#2E7D32' if is_profit else '#C62828'
    profit_bg    = '#E8F5E9' if is_profit else '#FFEBEE'
    profit_icon  = '📈' if is_profit else '📉'
    result_word  = 'PROFIT' if is_profit else 'LOSS'

    days = pnl['investment_days']
    if days < 30:
        holding_str = f"{days} days"
    elif days < 365:
        holding_str = f"{pnl['time_months']} months ({days} days)"
    else:
        holding_str = f"{round(days/365.25,1)} years ({days} days)"

    cr_color = '#2E7D32' if pnl['current_return'] >= 0 else '#C62828'

    subject = (f"{'📈' if is_profit else '📉'} P&L: {stock.get('ticker','')} | "
               f"{'+'if is_profit else ''}Rs.{pnl['total_profit']:,.0f} "
               f"({'+'if is_profit else ''}{pnl['return_pct']}%) | "
               f"Held {days} days")

    wohoo_banner = ''
    if target_hit:
        wohoo_banner = '<div style="background:linear-gradient(135deg,#FFD700,#FFA500);border-radius:10px;padding:20px;text-align:center;margin-bottom:20px;"><div style="font-size:48px;">🎯🎉</div><div style="font-size:28px;font-weight:bold;color:#1a1a2e;">Wohoooo! My analysis was right like always!</div><div style="font-size:14px;color:#1a1a2e;margin-top:8px;">Target price achieved exactly as predicted!</div></div>'

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:700px;margin:auto;background:#f5f5f5;padding:20px;">
  <div style="background:linear-gradient(135deg,#1a1a2e,#16213e);color:white;padding:24px;border-radius:12px;margin-bottom:20px;">
    <h1 style="margin:0;font-size:22px;">📋 Trade P&L Statement</h1>
    <p style="margin:6px 0 0;opacity:0.8;">{stock.get('stock_name','')} ({stock.get('ticker','')}) — Sold on {pnl['selling_date']}</p>
  </div>
  {wohoo_banner}
  <div style="background:{profit_bg};border:2px solid {profit_color};border-radius:10px;padding:20px;text-align:center;margin-bottom:20px;">
    <div style="font-size:36px;">{profit_icon}</div>
    <div style="font-size:13px;color:{profit_color};font-weight:bold;">{result_word}</div>
    <div style="font-size:38px;font-weight:bold;color:{profit_color};">{'+'if is_profit else ''}Rs.{pnl['total_profit']:,.2f}</div>
    <div style="font-size:18px;color:{profit_color};">{'+'if is_profit else ''}{pnl['return_pct']}% over {holding_str}</div>
  </div>
  <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 2px 6px rgba(0,0,0,0.1);">
    <h2 style="margin:0 0 16px;font-size:16px;border-bottom:2px solid #eee;padding-bottom:8px;">📊 Transaction Record</h2>
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;width:45%;">Industry</td><td style="padding:10px 14px;">{stock.get('industry','')}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Stock</td><td style="padding:10px 14px;">{stock.get('stock_name','')}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Buying Date</td><td style="padding:10px 14px;">{pnl['buying_date']}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Selling Date</td><td style="padding:10px 14px;">{pnl['selling_date']}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Buying Price</td><td style="padding:10px 14px;">Rs.{pnl['buying_price']:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Selling Price</td><td style="padding:10px 14px;">Rs.{pnl['selling_price']:,.2f}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Quantity</td><td style="padding:10px 14px;">{pnl['qty']} shares</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Investment Amount</td><td style="padding:10px 14px;">Rs.{pnl['investment_amt']:,.2f}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Profit per Share</td><td style="padding:10px 14px;color:{profit_color};font-weight:bold;">{'+'if is_profit else ''}Rs.{pnl['profit_per_share']:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Total Profit/Loss</td><td style="padding:10px 14px;color:{profit_color};font-weight:bold;font-size:16px;">{'+'if is_profit else ''}Rs.{pnl['total_profit']:,.2f}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Return %</td><td style="padding:10px 14px;color:{profit_color};font-weight:bold;">{'+'if is_profit else ''}{pnl['return_pct']}%</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Investment Days</td><td style="padding:10px 14px;">{holding_str}</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Current Share Price</td><td style="padding:10px 14px;">Rs.{pnl['current_price']:,.2f}</td></tr>
      <tr><td style="padding:10px 14px;font-weight:bold;">Current Return</td><td style="padding:10px 14px;color:{cr_color};font-weight:bold;">{'+'if pnl['current_return']>=0 else ''}{pnl['current_return']}%</td></tr>
      <tr style="background:#f8f9fa;"><td style="padding:10px 14px;font-weight:bold;">Time (months)</td><td style="padding:10px 14px;">{pnl['time_months']} months</td></tr>
    </table>
  </div>
  <div style="text-align:center;color:#aaa;font-size:11px;padding:12px;">Generated by Portfolio AI | Not financial advice</div>
</body></html>"""

    send_report_email(subject, html)
    print(f"[pnl_updater] P&L email sent for {stock.get('ticker','')}")


if __name__ == "__main__":
    test_stock = {
        'ticker':       'LT',
        'stock_name':   'Larsen & Toubro',
        'industry':     'Construction & Engineering',
        'buying_price': 2296.80,
        'buying_date':  '2023-05-11',
        'qty':          10,
        'target_hit':   False,
    }
    process_sell(test_stock, selling_price=4402.70)