
# PaperOn Sales Demo (Streamlit)

CSVをアップロードすると、**注文（Orders）**と**明細（OrderItems）**に正規化し、
注文登録の進捗をアニメーション表示するデモ用Webアプリです。

## 使い方

1. Python 3.10+ をインストール
2. 必要ライブラリをインストール
   ```bash
   pip install -r requirements.txt
   ```
3. アプリを起動
   ```bash
   streamlit run app.py
   ```
4. 画面左の **Upload & Register** からPaperOnのCSVをアップロード → **注文登録** ボタン

## メモ
- CSVのエンコーディングは `utf-8-sig / cp932(Shift_JIS) / utf-8` を自動判定します。
- 注文IDは、得意先・金額などのヘッダ情報から**安定的なハッシュ**で生成します（デモ用）。
- 将来、PaperOnの出力に**注文番号**が含まれる場合は、その列を使うように容易に変更可能です。
