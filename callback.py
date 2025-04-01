from flask import Flask, request

app = Flask(__name__)

@app.route('/callback')
def callback():
    # OAuth 認証コードを取得して処理する
    code = request.args.get('code')  # eBay から送信される認証コード
    return f"Received code: {code}"

if __name__ == '__main__':
    app.run(port=8000)