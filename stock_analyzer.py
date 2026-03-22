import os
import time
import requests
import urllib3
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from emailer import Emailer
from dotenv import load_dotenv

# 抑制 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 載入環境變數
load_dotenv()

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}

YAHOO_SESSION = None
YAHOO_CRUMB = None

def get_yahoo_auth():
    global YAHOO_SESSION, YAHOO_CRUMB
    if YAHOO_SESSION is None:
        try:
            YAHOO_SESSION = requests.Session()
            YAHOO_SESSION.headers.update(HEADERS)
            YAHOO_SESSION.get('https://fc.yahoo.com', verify=False)
            YAHOO_CRUMB = YAHOO_SESSION.get('https://query1.finance.yahoo.com/v1/test/getcrumb', verify=False).text.strip()
        except:
            YAHOO_CRUMB = ''
    return YAHOO_SESSION, YAHOO_CRUMB

def is_trading_day():
    # 簡易判斷：六日不執行
    today = datetime.now()
    if today.weekday() >= 5:
        return False
    return True

def fetch_twse_tpex_list():
    """獲取上市與上櫃的個股名單與今日價格，並過濾 50~300 元"""
    tickers = []
    prices_info = {}
    
    # 上市 (TWSE) 採用 MI_INDEX
    found_twse = False
    for i in range(7):
        target_date = (datetime.now() - timedelta(days=i)).strftime('%Y%m%d')
        url_twse = f'https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&type=ALLBUT0999&date={target_date}'
        try:
            res = requests.get(url_twse, timeout=10, verify=False, headers=HEADERS)
            data = res.json()
            if data.get('stat') == 'OK' and data.get('tables'):
                for tbl in data['tables']:
                    title = tbl.get('title', '')
                    fields = tbl.get('fields', [])
                    if '每日收盤行情' in title and '證券代號' in fields:
                        print(f"[OK] 取得 TWSE {target_date} 收盤行情 (MI_INDEX)")
                        idx_code = fields.index('證券代號')
                        idx_name = fields.index('證券名稱')
                        idx_close = fields.index('收盤價')
                        idx_sign = fields.index('漲跌(+/-)')
                        idx_diff = fields.index('漲跌價差')
                        
                        for row in tbl.get('data', []):
                            code = row[idx_code]
                            if len(str(code)) == 4:
                                try:
                                    close_val = float(str(row[idx_close]).replace(',', ''))
                                    
                                    # 計算漲跌價差的正負號
                                    sign_str = str(row[idx_sign])
                                    diff_val = float(str(row[idx_diff]).replace(',', ''))
                                    if 'green' in sign_str or '-' in sign_str:
                                        change_val = -diff_val
                                    elif 'red' in sign_str or '+' in sign_str:
                                        change_val = diff_val
                                    else:
                                        change_val = 0.0
                                        
                                    if 0 <= close_val <= 200:
                                        symbol = f"{code}.TW"
                                        tickers.append(symbol)
                                        prices_info[symbol] = {
                                            'Name': row[idx_name].strip(),
                                            'Change': change_val,
                                            'Market': '上市'
                                        }
                                except ValueError:
                                    pass
                        found_twse = True
                        break
                if found_twse:
                    break
            else:
                print(f"[Wait] TWSE 收盤行情 {target_date} 無資料，嘗試往前一天...")
        except Exception as e:
            print(f"Error fetching TWSE MI_INDEX {target_date}: {e}")

    # 上櫃 (TPEx) - 改用 FinMind 與 Yahoo Finance 替代
    try:
        print("Fetching TPEx latest prices from FinMind & Yahoo Finance...")
        res_info = requests.get("https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo", timeout=10)
        tpex_symbols = []
        tpex_names = {}
        if res_info.status_code == 200:
            for item in res_info.json().get('data', []):
                if item.get('type') == 'tpex' and len(str(item.get('stock_id'))) == 4:
                    sym = f"{item['stock_id']}.TWO"
                    tpex_symbols.append(sym)
                    tpex_names[sym] = item.get('stock_name', '')
                    
        session, crumb = get_yahoo_auth()
        chunk_size = 300
        for i in range(0, len(tpex_symbols), chunk_size):
            chunk = tpex_symbols[i:i+chunk_size]
            syms_str = ",".join(chunk)
            url_quote = f"https://query2.finance.yahoo.com/v7/finance/quote?symbols={syms_str}&crumb={crumb}"
            res_q = session.get(url_quote, timeout=10, verify=False)
            data_q = res_q.json()
            if 'quoteResponse' in data_q and 'result' in data_q['quoteResponse']:
                for q in data_q['quoteResponse']['result']:
                    sym = q.get('symbol')
                    close = q.get('regularMarketPrice', 0)
                    change = q.get('regularMarketChange', 0)
                    if 0 <= close <= 200:
                        tickers.append(sym)
                        prices_info[sym] = {
                            'Name': tpex_names.get(sym, ''),
                            'Change': change,
                            'Market': '上櫃'
                        }
            time.sleep(0.5)
    except Exception as e:
        print(f"Error fetching TPEx: {e}")
        
    return tickers, prices_info

def fetch_day_trading():
    """獲取全部標的當沖資訊 (TWSE 使用主網站 API 回推最近交易日)"""
    day_trade_map = {}
    
    # 1. 抓取 TWSE (有日期回推機制)
    found_twse = False
    for i in range(7):
        target_date = (datetime.now() - timedelta(days=i)).strftime('%Y%m%d')
        url = f'https://www.twse.com.tw/exchangeReport/TWTB4U?response=json&date={target_date}'
        try:
            res = requests.get(url, timeout=10, verify=False, headers=HEADERS)
            data = res.json()
            if data.get('stat') == 'OK' and data.get('tables'):
                for tbl in data['tables']:
                    fields = tbl.get('fields', [])
                    if '證券代號' in fields and '當日沖銷交易成交股數' in fields:
                        print(f"[OK] 取得 TWSE {target_date} 當沖資料")
                        code_idx = fields.index('證券代號')
                        vol_idx = fields.index('當日沖銷交易成交股數')
                        
                        for row in tbl.get('data', []):
                            code = f"{row[code_idx]}.TW"
                            vol_str = str(row[vol_idx]).replace(',', '')
                            if vol_str.isdigit():
                                day_trade_map[code] = {
                                    'shares': int(vol_str),
                                    'lots': int(vol_str) // 1000
                                }
                        found_twse = True
                        break
                if found_twse:
                    break
            else:
                print(f"[Wait] TWSE {target_date} 無資料，嘗試往前一天...")
        except Exception as e:
            print(f"TWSE {target_date} 請求失敗: {e}")
            
    # 2. 抓取 TPEx - 改用 FinMind 與 Yahoo Finance 總成交量替代 (因為當沖 API 被撤除)
    try:
        res_info = requests.get("https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo", timeout=10)
        if res_info.status_code == 200:
            session, crumb = get_yahoo_auth()
            tpex_symbols = [f"{item['stock_id']}.TWO" for item in res_info.json().get('data', []) if item.get('type') == 'tpex' and len(str(item.get('stock_id'))) == 4]
            chunk_size = 300
            for i in range(0, len(tpex_symbols), chunk_size):
                chunk = tpex_symbols[i:i+chunk_size]
                syms_str = ",".join(chunk)
                url_quote = f"https://query2.finance.yahoo.com/v7/finance/quote?symbols={syms_str}&crumb={crumb}"
                res_q = session.get(url_quote, timeout=10, verify=False)
                data_q = res_q.json()
                if 'quoteResponse' in data_q and 'result' in data_q['quoteResponse']:
                    for q in data_q['quoteResponse']['result']:
                        sym = q.get('symbol')
                        vol = q.get('regularMarketVolume', 0)
                        if vol > 0:
                            day_trade_map[sym] = {
                                'shares': vol,
                                'lots': vol // 1000
                            }
                time.sleep(0.5)
    except Exception as e:
        print(f"Error fetching TPEx volumes: {e}")
        
    return day_trade_map

def analyze_stocks():
    print("Starting stock analysis...")
    tickers, prices_info = fetch_twse_tpex_list()
    if not tickers:
        print("No tickers found 0~200.")
        return pd.DataFrame()

    print(f"Found {len(tickers)} stocks in price range 0~200. Processing limits to top 300 by volume to save time...")
    
    dt_map = fetch_day_trading()
    
    # 將 tickers 依照當沖量或是成交量排序，只截取最熱門的 300 檔來抓歷史資料，避免請求過多
    tickers_sorted = sorted(tickers, key=lambda x: dt_map.get(x, {}).get('lots', 0), reverse=True)[:300]
    
    results = []
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    import time
    
    for count, sym in enumerate(tickers_sorted):
        try:
            url = f"https://query2.finance.yahoo.com/v8/finance/chart/{sym}?interval=1d&range=15d"
            res = requests.get(url, headers=headers, timeout=5, verify=False)
            data = res.json()
            
            if 'chart' not in data or not data['chart']['result']:
                continue
                
            result = data['chart']['result'][0]
            if 'indicators' not in result or 'quote' not in result['indicators']:
                continue
                
            quote = result['indicators']['quote'][0]
            closes = pd.Series(quote['close']).dropna()
            highs = pd.Series(quote['high']).dropna()
            lows = pd.Series(quote['low']).dropna()
            opens = pd.Series(quote['open']).dropna()
            
            if len(closes) < 5:
                continue
                
            # 計算 5MA
            ma5 = closes.rolling(window=5).mean()
            last_ma5 = ma5.iloc[-1]
            prev_ma5 = ma5.iloc[-2]
            
            last_close = closes.iloc[-1]
            last_high = highs.iloc[-1]
            last_low = lows.iloc[-1]
            last_open = opens.iloc[-1]
            
            tr = highs - lows
            atr = tr.rolling(window=5).mean().iloc[-1]
            
            trend_up = last_ma5 > prev_ma5
            distance_pct = (last_close - last_ma5) / last_ma5 * 100
            
            if last_close > last_ma5 and trend_up:
                signal = '作多 (Long)'
            elif last_close < last_ma5 and not trend_up:
                signal = '放空 (Short)'
            else:
                signal = '觀望 (Hold)'

            if abs(distance_pct) > 7 or (atr/last_close) > 0.05:
                risk = '高風險 (High)'
            elif abs(distance_pct) > 4:
                risk = '中風險 (Medium)'
            else:
                risk = '低風險 (Low)'

            upper_bound = last_ma5 + atr
            lower_bound = last_ma5 - atr
            invest_range = f"{lower_bound:.1f} ~ {upper_bound:.1f}"
            
            info = prices_info.get(sym, {})
            name = info.get('Name', '')
            change_val = info.get('Change', 0)
            dt_info = dt_map.get(sym, {'shares': 0, 'lots': 0})
            dt_vol = dt_info['lots']
            dt_shares = dt_info['shares']
            
            prev_close = closes.iloc[-2]
            day_change_pct = (last_close - prev_close) / prev_close * 100
            
            if last_open > 0:
                intraday_change_pct = (last_close - last_open) / last_open * 100
            else:
                intraday_change_pct = 0
                
            score = 0
            if signal == '作多 (Long)': score += 10
            if signal == '放空 (Short)': score += 5
            if risk == '低風險 (Low)': score += 5
            
            results.append({
                '代號': sym.replace('.TW', '').replace('.TWO', ''),
                '名稱': name,
                '開盤價': round(last_open, 2),
                '最高價': round(last_high, 2),
                '最低價': round(last_low, 2),
                '收盤價': round(last_close, 2),
                '漲跌價差': round(change_val, 1),
                '開盤後漲跌(%)': round(intraday_change_pct, 2),
                '當沖成交股數(K)': f"{dt_shares / 1000:,.0f}",
                '5MA': round(last_ma5, 2),
                '最高與最低價差(元)': round(last_high - last_low, 2),
                '投資建議': signal,
                '推估投資區間': invest_range,
                '風險評估': risk
            })
            time.sleep(0.05) # 防阻擋
        except Exception as e:
            continue
            
    df_res = pd.DataFrame(results)
    if not df_res.empty:
        # 標註價格區間
        def get_bucket(price):
            if 0 <= price <= 50:
                return '0~50'
            elif 50 < price <= 100:
                return '51~100'
            elif 100 < price <= 200:
                return '100~200'
            else:
                return 'Other'
        df_res['價格區間'] = df_res['收盤價'].apply(get_bucket)
        df_res = df_res[df_res['價格區間'] != 'Other']
        
        # 投資建議轉為數值以便排序 (作多 > 觀望 > 放空)
        signal_map = {'作多 (Long)': 3, '觀望 (Hold)': 2, '放空 (Short)': 1}
        df_res['SignalScore'] = df_res['投資建議'].map(signal_map).fillna(0)
        
        # 整體排序: 價格區間, 開盤後漲跌(%) (降冪), SignalScore (降冪)
        df_res = df_res.sort_values(
            by=['價格區間', '開盤後漲跌(%)', 'SignalScore'], 
            ascending=[True, False, False]
        )
        
    return df_res

def generate_html(df):
    html = """
    <html>
    <head>
    <style>
        table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even){background-color: #f9f9f9;}
        tr:hover {background-color: #f1f1f1;}
        .up { color: red; font-weight: bold; }
        .down { color: green; font-weight: bold; }
    </style>
    </head>
    <body>
    <h2>台灣股市每日分析報告 (5MA 策略)</h2>
    <p>過濾條件：股價 0~200 元。排序方式：各區間與分類內依據開盤後漲跌(%)降冪排序。</p>
    """
    if df.empty:
        html += "<p>今日無符合條件之股票。</p>"
    else:
        # 分別產生三個價格區間的表格
        buckets = ['0~50', '51~100', '100~200']
        signals = ['作多 (Long)', '放空 (Short)', '觀望 (Hold)']
        
        for bucket in buckets:
            html += f"<h3>收盤價區間：{bucket}</h3>"
            sub_df_bucket = df[df['價格區間'] == bucket].drop(columns=['價格區間', 'SignalScore'], errors='ignore')
            
            if sub_df_bucket.empty:
                html += "<p>此區間無符合股票。</p>"
                continue
                
            for sig in signals:
                html += f"<h4>投資建議：{sig}</h4>"
                sub_df = sub_df_bucket[sub_df_bucket['投資建議'] == sig].drop(columns=['投資建議'], errors='ignore')
                
                if sub_df.empty:
                    html += "<p style='color:grey;'>此分類無股票。</p>"
                else:
                    # HTML 著色處理
                    tbl_html = sub_df.to_html(index=False, border=0, classes='table')
                    tbl_html = tbl_html.replace('高風險 (High)', '<span style="color:orange;">高風險 (High)</span>')
                    html += tbl_html

    # 加入資料來源引註
    html += """
    <hr>
    <h3>參考資料來源：</h3>
    <ul>
        <li>TWSE 上市當沖交易標的及成交量值：<a href="https://www.twse.com.tw/zh/trading/day-trading.html">https://www.twse.com.tw/zh/trading/day-trading.html</a></li>
        <li>TWSE 每日收盤行情 (漲跌價差依據)：<a href="https://www.twse.com.tw/zh/trading/historical/mi-index.html">https://www.twse.com.tw/zh/trading/historical/mi-index.html</a></li>
        <li>TPEx 上櫃當沖交易標的及成交量值：<a href="https://www.tpex.org.tw/openapi/v1/t187ap14_L">https://www.tpex.org.tw/openapi/v1/t187ap14_L</a></li>
        <li>TPEx 上櫃個股日成交資訊：<a href="https://www.tpex.org.tw/openapi/v1/t187ap03_L">https://www.tpex.org.tw/openapi/v1/t187ap03_L</a></li>
        <li>個股技術分析與歷史價格：Yahoo Finance (query2.finance.yahoo.com)</li>
    </ul>
    """
    html += "</body></html>"
    return html

def main():
    import sys
    force_run = '--force' in sys.argv
    if not is_trading_day() and not force_run:
        print("Today is weekend. Skipping analysis.")
        return

    df = analyze_stocks()
    html_content = generate_html(df)
    
    sender = os.getenv("GMAIL_USER")
    pwd = os.getenv("GMAIL_APP_PASSWORD")
    recpts = os.getenv("MAIL_RECIPIENT", "rorcl26@gmail.com")
    
    emailer = Emailer(sender, pwd)
    today_str = datetime.now().strftime("%Y-%m-%d")
    subject = f"[{today_str}] 台灣股市 5MA 策略分析報告"
    
    if sender and pwd:
        emailer.send_email(recpts, subject, html_content)
    else:
        # 輸出到檔案做測試
        with open('report.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        print("Email credentials not set. Report saved to report.html")

if __name__ == "__main__":
    main()
