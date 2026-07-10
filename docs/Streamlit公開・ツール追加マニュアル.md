# Streamlit初回公開・ツール追加マニュアル

## 目的

このマニュアルは、Streamlitで作成したPythonツールを、GitHub経由でWebアプリとして公開し、後から新しいツールを追加できるようにするための記録です。

一度作ったStreamlitアプリを忘れても、以下を再現できる状態にすることを目的とします。

* GitHubリポジトリの作成
* ローカルPCとの接続
* Streamlitアプリの初回公開
* ツール追加
* GitHubへのpush
* Streamlit側の反映確認
* エラー時の確認

---

## 全体の流れ

```text
ローカルPCでStreamlitアプリを作る
↓
GitHubリポジトリを作る
↓
ローカルフォルダとGitHubを接続する
↓
GitHubにpushする
↓
Streamlit Community CloudでGitHubリポジトリを指定して公開する
↓
URLが発行される
↓
以後はコード修正 → GitHubにpush → Streamlitに反映
```

---

# 1. 使用するサービス・ツール

## 使用サービス

* GitHub
* Streamlit Community Cloud

## 使用ツール

* Python
* Streamlit
* VS Code または Claude Code
* ターミナル / PowerShell
* Git

---

# 2. 事前準備

## 必要なアカウント

* GitHubアカウント
* Streamlit Community Cloudアカウント

StreamlitはGitHubと連携して使うため、GitHubアカウントが必要です。

---

## PCに必要なもの

* Python
* pip
* Git
* Streamlit

確認コマンド：

```bash
python --version
pip --version
git --version
```

Streamlitが入っていない場合：

```bash
pip install streamlit
```

---

# 3. ローカルでStreamlitアプリを作る

## 基本フォルダ構成

最初はこの形で作る。

```text
streamlit-tools/
  app.py
  requirements.txt
  README.md
  .gitignore
```

複数ツールを管理する場合は、後で `pages` フォルダを使う。

```text
streamlit-tools/
  app.py
  requirements.txt
  README.md
  .gitignore
  pages/
    01_顧客リスト整理.py
    02_画像リサイズ.py
    03_勤怠集計.py
```

---

## app.py の最小構成

```python
import streamlit as st

st.set_page_config(
    page_title="業務効率化ツール管理",
    page_icon="🛠️",
    layout="wide"
)

st.title("業務効率化ツール管理")
st.write("ここから各ツールを選んで使います。")
```

---

## requirements.txt の例

```txt
streamlit
pandas
openpyxl
pillow
```

使うライブラリに応じて追加する。

Excel処理なら：

```txt
pandas
openpyxl
```

画像処理なら：

```txt
pillow
```

---

## .gitignore の例

```txt
__pycache__/
*.pyc
.env
.env.local
.streamlit/secrets.toml
.DS_Store
```

注意：

```text
APIキー、パスワード、SupabaseキーなどはGitHubに上げない。
.env、.env.local、secrets.toml は必ず除外する。
```

---

# 4. ローカルで動作確認する

作業フォルダに移動する。

```bash
cd streamlit-tools
```

Streamlitを起動する。

```bash
streamlit run app.py
```

ブラウザで画面が開けばOK。

確認すること：

* 画面が開くか
* エラーが出ていないか
* ファイルアップロードが動くか
* ダウンロードが動くか
* ExcelやCSVが正しく処理されるか

---

# 5. GitHubリポジトリを作成する

## GitHubで新規リポジトリ作成

GitHubにログインする。

右上の `+` から `New repository` を選ぶ。

リポジトリ名を入力する。

例：

```text
streamlit-tools
```

公開設定を選ぶ。

```text
Public または Private
```

注意：

最初にローカルからpushする場合、GitHub側ではREADMEや.gitignoreを作らない方がエラーが少ない。

---

# 6. ローカルフォルダをGit管理する

作業フォルダで以下を実行する。

```bash
git init
git add .
git commit -m "初回Streamlitアプリ作成"
```

---

# 7. ローカルフォルダとGitHubを接続する

GitHubで作成したリポジトリURLをコピーする。

例：

```text
https://github.com/ユーザー名/streamlit-tools.git
```

ローカルで以下を実行する。

```bash
git remote add origin GitHubのリポジトリURL
```

接続確認：

```bash
git remote -v
```

問題なければGitHubへpushする。

```bash
git branch -M main
git push -u origin main
```

---

# 8. Streamlit Community Cloudで初回公開する

Streamlit Community Cloudにログインする。

`Create app` を選ぶ。

GitHub連携を求められたら許可する。

以下を指定する。

```text
Repository：GitHubのリポジトリ
Branch：main
Main file path：app.py
App URL：任意
```

例：

```text
Repository：futurewithai24/streamlit-tools
Branch：main
Main file path：app.py
```

`Deploy` を押す。

---

# 9. 公開後に確認すること

公開URLが発行される。

確認すること：

* URLを開けるか
* app.pyの画面が表示されるか
* ファイルアップロードができるか
* ダウンロードができるか
* 日本語表示が崩れていないか
* エラーが出ていないか

---

# 10. ツールを追加する方法

## pages フォルダを使う場合

フォルダ構成：

```text
streamlit-tools/
  app.py
  pages/
    01_顧客リスト整理.py
    02_画像リサイズ.py
```

新しいツールを追加する場合：

```text
pages/
  03_新しいツール.py
```

ファイル名の例：

```text
03_勤怠集計.py
04_顧客リスト重複チェック.py
05_画像一括リサイズ.py
```

---

## 新しいツールファイルの例

```python
import streamlit as st
import pandas as pd

st.title("新しいツール")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("読み込み結果")
    st.dataframe(df)
```

---

# 11. ツール追加後にローカル確認する

```bash
streamlit run app.py
```

確認すること：

* サイドバーに新しいページが表示されるか
* ツール画面が開くか
* アップロードが動くか
* 処理結果が正しいか
* エラーが出ないか

---

# 12. GitHubへ更新を反映する

変更状況を確認する。

```bash
git status
```

変更を追加する。

```bash
git add .
```

コミットする。

```bash
git commit -m "新しいツールを追加"
```

GitHubへpushする。

```bash
git push
```

---

# 13. Streamlit側の反映確認

GitHubにpushしたあと、Streamlitの公開URLを開く。

確認すること：

* 新しいツールが追加されているか
* 古い画面のままになっていないか
* エラーが出ていないか

反映されない場合：

* ブラウザを更新する
* StreamlitのManage appを開く
* Logsを確認する
* 必要に応じてReboot appする

---

# 14. Secretsを使う場合

APIキー、Supabase URL、Supabase anon keyなどを使う場合、コードに直接書かない。

ローカルでは以下のファイルを使う。

```text
.streamlit/secrets.toml
```

例：

```toml
SUPABASE_URL = "https://xxxxx.supabase.co"
SUPABASE_ANON_KEY = "xxxxx"
```

Pythonコードでは以下のように読む。

```python
import streamlit as st

supabase_url = st.secrets["SUPABASE_URL"]
supabase_key = st.secrets["SUPABASE_ANON_KEY"]
```

注意：

```text
.streamlit/secrets.toml はGitHubにpushしない。
.gitignore に必ず追加する。
```

Streamlit本番環境では、アプリ設定のSecrets欄に同じ内容を登録する。

---

# 15. よくあるエラーと確認場所

## ModuleNotFoundError

例：

```text
ModuleNotFoundError: No module named 'openpyxl'
```

原因：

```text
requirements.txt に必要なライブラリが書かれていない。
```

対応：

```text
requirements.txt に openpyxl を追加して git push する。
```

---

## File not found

原因：

```text
Streamlitで指定した起動ファイルのパスが違う。
```

確認：

```text
Main file path が app.py になっているか確認する。
```

---

## Secretsエラー

原因：

```text
st.secrets で呼び出しているキーがStreamlit側に登録されていない。
```

確認：

```text
StreamlitのManage app → Settings → Secrets を確認する。
```

---

## GitHubに反映されない

確認：

```bash
git status
git log --oneline
git remote -v
```

対応：

```bash
git add .
git commit -m "修正"
git push
```

---

# 16. 毎回の作業チェックリスト

## 作業前

* [ ] どのGitHubリポジトリを編集するか確認した
* [ ] ローカルの作業フォルダを開いた
* [ ] 現在のファイル構成を確認した
* [ ] requirements.txtを確認した
* [ ] APIキーをコードに直書きしていない

## 作業中

* [ ] 新しいツールファイルを作成した
* [ ] app.py または pages フォルダの構成を確認した
* [ ] ローカルで `streamlit run app.py` を実行した
* [ ] エラーがないことを確認した
* [ ] サンプルデータで動作確認した

## 作業後

* [ ] git status を確認した
* [ ] git add . を実行した
* [ ] git commit した
* [ ] git push した
* [ ] Streamlitの公開URLで反映確認した
* [ ] エラーがあればLogsを確認した
* [ ] 作業内容を記録管理部にメモした

---

# 17. 作業記録テンプレート

```text
日付：
作業名：
対象リポジトリ：
対象Streamlit URL：
追加・修正したファイル：
使用したライブラリ：
実行したコマンド：
発生したエラー：
解決方法：
次回注意すること：
```

---

# 18. 販売・代行作成時の注意

購入者向けに作成する場合は、以下を最初に決める。

```text
1. mai側の環境でデモURLを作るのか
2. 購入者のGitHub・Streamlitに設置するのか
3. ソースコードも渡すのか
4. 修正対応は何日までか
5. 個人情報や機密情報を扱うのか
```

おすすめは、購入者自身のGitHub・Streamlit環境に設置する方法。

理由：

```text
購入者のデータをmai側で管理し続けなくて済む。
納品後の管理責任を分けやすい。
販売後のトラブルを減らしやすい。
```

---

# 19. 次に作る関連マニュアル

このマニュアルの次に作るもの：

```text
02_GitHub基本操作マニュアル.md
03_Streamlitエラー対応マニュアル.md
04_Supabase初回作成マニュアル.md
05_Vercel公開マニュアル.md
06_購入者向け納品マニュアル.md
```
