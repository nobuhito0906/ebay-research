import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import json
import time
import pandas as pd
from datetime import datetime

# Google APIの認証設定
def setup_google_api():
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    # サービスアカウントのJSONファイルパスを設定してください
    credentials = ServiceAccountCredentials.from_json_keyfile_name('./config/google-credentials.json', scope)
    client = gspread.authorize(credentials)
    return client

# eBay APIを使用して検索を実行する関数
def search_ebay_api(keyword):
    try:
        # eBay APIのエンドポイントとAPIキー
        endpoint = "https://api.ebay.com/buy/browse/v1/item_summary/search"
        
        # ここにあなたのeBay APIキーを設定
        api_key = "YOUR_EBAY_API_KEY"
        
        # ヘッダー設定
        headers = {
            'Authorization': f'Bearer {api_key}',
            'X-EBAY-C-MARKETPLACE-ID': 'EBAY_US',  # 米国のeBayを検索
            'Content-Type': 'application/json'
        }
        
        # クエリパラメータ（中古品のみのフィルターなど設定可能）
        params = {
            'q': keyword,
            'limit': 5,  # 最大5件の結果を取得
            'filter': 'conditions:{USED}',  # 中古品のみ
        }
        
        # API呼び出し
        response = requests.get(endpoint, headers=headers, params=params)
        response.raise_for_status()
        
        data = response.json()
        
        # 検索結果を処理
        total_results = data.get('total', 0)
        items = data.get('itemSummaries', [])
        
        # 商品情報を抽出
        top_items = []
        for item in items[:5]:
            top_items.append({
                'title': item.get('title', 'N/A'),
                'price': f"{item.get('price', {}).get('value', 'N/A')} {item.get('price', {}).get('currency', '')}",
                'condition': item.get('condition', 'N/A'),
                'url': item.get('itemWebUrl', 'N/A'),
                'shipping': f"{item.get('shippingOptions', [{}])[0].get('shippingCost', {}).get('value', 'N/A')} {item.get('shippingOptions', [{}])[0].get('shippingCost', {}).get('currency', '')}" if item.get('shippingOptions') else 'N/A'
            })
        
        return {
            'total_results': total_results,
            'top_items': top_items,
            'search_url': f"https://www.ebay.com/sch/i.html?_nkw={requests.utils.quote(keyword)}&_sacat=0&LH_ItemCondition=3000"
        }
        
    except Exception as e:
        print(f"Error searching for '{keyword}': {str(e)}")
        return {
            'total_results': 0,
            'top_items': [],
            'search_url': f"https://www.ebay.com/sch/i.html?_nkw={requests.utils.quote(keyword)}&_sacat=0",
            'error': str(e)
        }

# メイン関数
def main():
    # Google APIのセットアップ
    client = setup_google_api()
    
    try:
        # スプレッドシートを開く（スプレッドシート名またはIDを設定してください）
        # IDを使用する場合:
        # spreadsheet = client.open_by_key('あなたのスプレッドシートID')
        # 名前を使用する場合:
        spreadsheet = client.open('ebay_searchword')  # スプレッドシート名を適宜変更
        
        # ワークシートを選択
        worksheet = spreadsheet.worksheet('Keywords')  # ワークシート名を適宜変更
        
        # A列からキーワードを取得
        keywords = worksheet.col_values(1)[1:]  # ヘッダーを除外
        
        # 結果を格納するためのデータフレームを作成
        results_df = pd.DataFrame(columns=[
            'Keyword', 'Total Results', 'Search URL',
            'Item 1 Title', 'Item 1 Price', 'Item 1 Condition', 'Item 1 Shipping', 'Item 1 URL',
            'Item 2 Title', 'Item 2 Price', 'Item 2 Condition', 'Item 2 Shipping', 'Item 2 URL',
            'Item 3 Title', 'Item 3 Price', 'Item 3 Condition', 'Item 3 Shipping', 'Item 3 URL',
            'Item 4 Title', 'Item 4 Price', 'Item 4 Condition', 'Item 4 Shipping', 'Item 4 URL',
            'Item 5 Title', 'Item 5 Price', 'Item 5 Condition', 'Item 5 Shipping', 'Item 5 URL',
            'Search Date'
        ])
        
        # 各キーワードで検索を実行
        for keyword in keywords:
            print(f"Searching for: {keyword}")
            
            # eBay APIで検索
            search_results = search_ebay_api(keyword)
            
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
                row_data[f'Item {idx} Condition'] = item['condition']
                row_data[f'Item {idx} Shipping'] = item['shipping']
                row_data[f'Item {idx} URL'] = item['url']
            
            # データフレームに行を追加
            results_df = pd.concat([results_df, pd.DataFrame([row_data])], ignore_index=True)
            
            # APIレート制限を避けるために待機
            time.sleep(1)  # eBay APIのレート制限に応じて調整
        
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
        
    except gspread.exceptions.SpreadsheetNotFound:
        print("Error: Spreadsheet not found. Please check the spreadsheet name or ID.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()