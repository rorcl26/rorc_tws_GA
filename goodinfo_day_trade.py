import os
import time
import pandas as pd
from datetime import datetime
import io
from playwright.sync_api import sync_playwright, TimeoutError

class GoodinfoPlaywrightScraper:
    """
    使用 Playwright 隱藏瀏覽器機制，繞過 Goodinfo 的 Cloudflare 防爬蟲，
    直接抓取「現股當沖率(當日)」的總表 (共 1971 筆資料)。
    """
    def __init__(self, output_dir="data"):
        self.output_dir = output_dir
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def fetch_master_list(self):
        print("啟動 Playwright 瀏覽器...")
        # 設定 Goodinfo 的現股當沖總表網址
        url = "https://goodinfo.tw/tw/StockList.asp?MARKET_CAT=全部&INDUSTRY_CAT=現股當沖率+(當日)%40%40現股當沖%40%40現股當沖率+(當日)"
        
        with sync_playwright() as p:
            # 啟動 Chromium，headless=False 會彈出視窗，能大幅提高繞過 Cloudflare 驗證的機率
            browser = p.chromium.launch(headless=False)
            
            # 建立一個擬真的瀏覽器上下文，增加繞過機率
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                viewport={"width": 1920, "height": 1080}
            )
            
            page = context.new_page()
            
            # 注入簡單的防偵測腳本
            page.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                });
            """)
            
            print("正在導向 Goodinfo... (若遇 Cloudflare 檢查可能需要幾秒鐘)")
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=30000)
                
                # 等待特定區塊載入，或直接等 10 秒
                page.wait_for_timeout(10000)
                # 嘗試等待新版 Goodinfo 資料表
                try:
                    page.wait_for_selector('#txtStockListData table', timeout=5000)
                except:
                    pass
                
                print("成功載入網頁與資料表！")
                
                # 取得整頁 HTML（使用 JS 抓取真正的 outerHTML，避免某些 JS 隱藏內容）
                html_content = page.evaluate("() => document.documentElement.outerHTML")
                
            except TimeoutError:
                print("載入逾時或被阻擋。嘗試直接擷取目前畫面並儲存截圖...")
                page.screenshot(path=os.path.join(self.output_dir, 'error_screenshot.png'), full_page=True)
                html_content = page.evaluate("() => document.documentElement.outerHTML")
            finally:
                browser.close()
                
        return html_content

    def analyze_html(self, html_content):
        print("\n開始解析 HTML 資料表...")
        try:
            # Pandas 讀取 HTML 內的 table
            tables = pd.read_html(io.StringIO(html_content))
            
            target_df = None
            for idx, df in enumerate(tables):
                if df.shape[0] < 50:
                    continue  # 略過太小的表格
                    
                if isinstance(df.columns, pd.MultiIndex):
                    cols = df.columns.get_level_values(-1)
                else:
                    cols = df.columns
                    
                # 只要有代號與名稱欄位，且列數 > 50，通常就是主資料表
                if '代號' in cols and '名稱' in cols:
                    df.columns = cols
                    target_df = df
                    print(f"找到目標資料表！大小 {df.shape}")
                    break
                    
            if target_df is None:
                print("在網頁中找不到符合的當沖資料表。可能因 Goodinfo 結構改變或遭到驗證碼攔截。")
                return None
                
            # 整理與過濾資料
            df = target_df.copy()
            
            # 清除廣告列或無效資料 (通常代號非數字)
            df = df.dropna(subset=['代號'])
            df = df[df['代號'].astype(str).str.isnumeric()]
            
            # 使用者需求欄位對應：需顯示股價日期、成交張數、現股當沖張數、現股當沖率(%)、成交漲跌價、漲跌幅
            # Goodinfo 清單上「股價日期」通常統一為今日，但沒有專屬欄位，我們手動補上
            today_str = datetime.now().strftime('%Y-%m-%d')
            df['股價日期'] = today_str
            
            # 找出對應的欄位名稱
            col_map = {}
            for c in df.columns:
                c_str = str(c)
                if '成交' in c_str and '張數' in c_str:
                    col_map[c] = '成交張數'
                elif '當沖' in c_str and '張數' in c_str:
                    col_map[c] = '現股當沖張數'
                elif '當沖率' in c_str:
                    col_map[c] = '現股當沖率(%)'
                elif '漲跌價' in c_str or c_str == '漲跌':
                    col_map[c] = '成交漲跌價'
                elif '漲跌幅' in c_str:
                    col_map[c] = '漲跌幅(%)'
                    
            # 重新命名欄位
            df = df.rename(columns=col_map)
            
            # 確保數字格式正確，移除可能的 % 與逗號
            convert_cols = ['成交張數', '現股當沖張數', '現股當沖率(%)', '成交漲跌價', '漲跌幅(%)']
            for col in convert_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '').str.replace('%', ''), errors='coerce')
                    
            # 依照當沖率降冪排序
            if '現股當沖率(%)' in df.columns:
                df = df.sort_values(by='現股當沖率(%)', ascending=False).reset_index(drop=True)
                
            print(f"成功解析 {len(df)} 筆個股當沖資料！")
            
            # 只要顯示必要欄位
            display_cols = ['股價日期', '代號', '名稱'] + [c for c in convert_cols if c in df.columns]
            final_df = df[[c for c in display_cols if c in df.columns]]
            
            print("\n=== 當沖率最高前 20 名股 ===")
            print(final_df.head(20).to_string(index=False))
            
            # 匯出資料
            output_file = os.path.join(self.output_dir, f"goodinfo_playwright_day_trade_{datetime.now().strftime('%Y%m%d')}.csv")
            final_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            print(f"\n[Success] 完整資料已順利儲存至：{output_file}")
            
            return final_df
            
        except Exception as e:
            err_str = str(e).encode('cp950', 'replace').decode('cp950')
            print(f"解析 HTML 時發生錯誤：{err_str[:200]}...")
            with open(os.path.join(self.output_dir, 'error_page.html'), 'w', encoding='utf-8') as f:
                f.write(html_content)
            return None

if __name__ == "__main__":
    scraper = GoodinfoPlaywrightScraper()
    html = scraper.fetch_master_list()
    if html:
        df = scraper.analyze_html(html)
