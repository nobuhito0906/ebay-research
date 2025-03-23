import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup
import time
import re
import pandas as pd
from datetime import datetime

# Google APIの認証設定
def setup_google_api():
    scope = ['https://www.googleapis.com/auth/spreadsheets',
             'https://www.googleapis.com/auth/drive']
    
    # サービスアカウントのJSONファイルパスを設定してください
    credentials = ServiceAccountCredentials.from_json_keyfile_name('./config/google-credentials.json', scope)
    client = gspread.authorize(credentials)
    return client

# eBayで検索を実行し結果を取得する関数
def search_ebay(keyword):
    try:
        # キーワードをURLエンコード
        encoded_keyword = requests.utils.quote(keyword)
        url = f"https://www.ebay.com/sch/i.html?_nkw={encoded_keyword}&_sacat=0&LH_ItemCondition=3000"
        
        # ユーザーエージェントを設定
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': 'https://www.ebay.com/',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Cache-Control': 'max-age=0'
        }
        
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # HTTPエラーがあれば例外を発生
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 検索結果の総数を取得
        total_results_elem = soup.select_one('.srp-controls__count-heading')
        total_results = 0
        if total_results_elem:
            total_text = total_results_elem.text.strip()
            numbers = re.findall(r'\d+,?\d*', total_text)
            if numbers:
                total_results = int(numbers[0].replace(',', ''))
        
        # 商品リストを取得
        items = soup.select('li.s-item')
        
        # 最初の5つの商品情報を取得
        top_items = []
        for i, item in enumerate(items[:5]):
            if i >= 5:
                break
                
            title_elem = item.select_one('.s-item__title')
            price_elem = item.select_one('.s-item__price')
            shipping_elem = item.select_one('.s-item__shipping')
            
            title = title_elem.text.strip() if title_elem else "N/A"
            price = price_elem.text.strip() if price_elem else "N/A"
            shipping = shipping_elem.text.strip() if shipping_elem else "N/A"
            
            item_url_elem = item.select_one('a.s-item__link')
            item_url = item_url_elem['href'] if item_url_elem else "N/A"
            
            top_items.append({
                'title': title,
                'price': price,
                'shipping': shipping,
                'url': item_url
            })
        
        return {
            'total_results': total_results,
            'top_items': top_items,
            'search_url': url
        }
        
    except Exception as e:
        print(f"Error searching for '{keyword}': {str(e)}")
        return {
            'total_results': 0,
            'top_items': [],
            'search_url': url if 'url' in locals() else f"https://www.ebay.com/sch/i.html?_nkw={requests.utils.quote(keyword)}&_sacat=0",
            'error': str(e)
        }

# メイン関数
def main():
    # Google APIのセットアップ
    client = setup_google_api()
    
    # スプレッドシートを開く（スプレッドシート名を設定してください）
    spreadsheet = client.open('ebay_searchword')
    
    # ワークシートを選択（ワークシート名を設定してください）
    worksheet = spreadsheet.worksheet('Keywords')
    
    # A列からキーワードを取得
    keywords = worksheet.col_values(1)[1:]  # ヘッダーを除外
    
    # 結果を格納するためのデータフレームを作成
    results_df = pd.DataFrame(columns=[
        'Keyword', 'Total Results', 'Search URL',
        'Item 1 Title', 'Item 1 Price', 'Item 1 Shipping', 'Item 1 URL',
        'Item 2 Title', 'Item 2 Price', 'Item 2 Shipping', 'Item 2 URL',
        'Item 3 Title', 'Item 3 Price', 'Item 3 Shipping', 'Item 3 URL',
        'Item 4 Title', 'Item 4 Price', 'Item 4 Shipping', 'Item 4 URL',
        'Item 5 Title', 'Item 5 Price', 'Item 5 Shipping', 'Item 5 URL',
        'Search Date'
    ])
    
    # 各キーワードで検索を実行
    for keyword in keywords:
        print(f"Searching for: {keyword}")
        
        # eBayで検索
        search_results = search_ebay(keyword)
        
        # 結果を行として準備
        row_data = {
            'Keyword': keyword,
            'Total Results': search_results['total_results'],
            'Search URL': search_results['search_url'],
            'Search Date': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # 最大5つの商品情報を追加
        for i, item in enumerate(search_results['top_items']):
            idx = i + 1
            row_data[f'Item {idx} Title'] = item['title']
            row_data[f'Item {idx} Price'] = item['price']
            row_data[f'Item {idx} Shipping'] = item['shipping']
            row_data[f'Item {idx} URL'] = item['url']
        
        # データフレームに行を追加
        results_df = pd.concat([results_df, pd.DataFrame([row_data])], ignore_index=True)
        
        # APIレート制限を避けるために少し待機
        time.sleep(5)
    
    # 結果のワークシートを作成または更新
    try:
        results_worksheet = spreadsheet.worksheet('Results')
        # 既存のワークシートをクリア
        results_worksheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        # ワークシートが存在しない場合は新規作成
        results_worksheet = spreadsheet.add_worksheet(title='Results', rows=len(results_df)+1, cols=len(results_df.columns))
    
    # ヘッダーとデータを書き込み
    results_worksheet.update([results_df.columns.tolist()] + results_df.values.tolist())
    
    print(f"Research completed. Processed {len(keywords)} keywords.")

if __name__ == "__main__":
    main()