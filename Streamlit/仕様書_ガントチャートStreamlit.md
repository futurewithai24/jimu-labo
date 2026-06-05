# 業務効率化ツール集 Streamlitアプリ 仕様書

作成日：2026-06-05  
ステータス：ガントチャート完成・ローカル動作確認済み  
目標：複数の業務ツールをまとめたWebサイト（ポートフォリオ）を作成し、社内活用・将来的な社外販売を目指す

---

## 概要

Python製のガントチャート生成スクリプト（gantt_chart.py）をStreamlitでWebアプリ化したもの。  
ブラウザ上でCSVをアップロードするだけでガントチャートExcelを生成できる。

---

## 目的・背景

- 実践塾コミュニティでガントチャート記事が好評 →「Webアプリ化して社内で使えるようにすれば？」とアドバイスをもらったことがきっかけ
- **フェーズ1：** 社内WebサイトにiframeでURL埋め込み → 社内業務ツールとして活用
- **フェーズ2：** 将来的に業務改善ツールとして社外販売も視野

---

## 元スクリプトとの違い

| 項目 | 元スクリプト（gantt_chart.py） | Streamlitアプリ（app.py） |
|---|---|---|
| CSVの入力 | ローカルファイルパス指定 | ブラウザでアップロード |
| Excel出力 | ローカルに .xlsx 保存 | ダウンロードボタンで取得 |
| 設定変更 | スクリプト内の定数を直接編集 | 画面上のUIで変更 |
| 実行方法 | `python gantt_chart.py` | ブラウザでアクセスするだけ |

---

## ファイル構成

```
E:\Antigravity\Streamlit\
├── 仕様書_ガントチャートStreamlit.md   ← 本ファイル
└── app.py                               ← Streamlitアプリ本体（作成済み）
```

---

## アプリの画面構成

### サイドバー（左側・設定エリア）
- **表示モード**：daily（3ヶ月以内向け）/ weekly（年間向け）をセレクトボックスで選択
- **ガントバーの色**：HEXコードで入力（例：4472C4 = 青）
- **今日の列の色**：HEXコードで入力（例：FFE699 = 黄）
- **土日の色**：HEXコードで入力（例：F2F2F2 = 薄グレー）

### メイン画面（右側）
1. タイトル・説明文
2. サンプルCSVダウンロードボタン（初回利用者向け）
3. CSVアップロードエリア
4. アップロード後：タスク一覧のプレビュー表示
5. 「ガントチャートを生成する」ボタン
6. 生成後：Excelダウンロードボタン

---

## 入力仕様（CSVフォーマット）

| 列名 | 必須 | 内容 | 例 |
|---|---|---|---|
| タスク名 | ✅ | タスクの名前 | 企画書作成 |
| 開始日 | ✅ | YYYY/MM/DD または YYYY-MM-DD | 2026/06/01 |
| 終了日 | ✅ | YYYY/MM/DD または YYYY-MM-DD | 2026/06/15 |
| 担当者 | 任意 | 担当者名（なくてもOK） | 田中 |

---

## 出力仕様

- Excelファイル（.xlsx）
- ガントチャートシート1枚
- 元スクリプトと同じデザイン・フォント（Meiryo UI）

---

## 技術構成

| 項目 | 内容 |
|---|---|
| 言語 | Python 3.x |
| フレームワーク | Streamlit |
| Excel生成 | openpyxl |
| デプロイ先（予定） | Streamlit Cloud（無料） |

---

---

# セットアップ手順（ゼロから作る場合）

> この章は、自分でゼロから同じアプリを作りたいときのための手順です。  
> すでにapp.pyが完成している場合は「アプリの起動方法」だけ参照してください。

---

## STEP 1：Pythonの確認

Streamlitを使うには、まずPythonがインストールされている必要があります。

### Windowsの場合
1. スタートメニューから「PowerShell」を開く（または「コマンドプロンプト」でもOK）
2. 以下のコマンドを入力してEnter

```
python --version
```

3. `Python 3.x.x` と表示されればOK
4. 何も表示されない・エラーになる場合はPythonのインストールが必要（python.org からインストール）

### Macの場合
1. 「ターミナル」を開く（Spotlight検索で「ターミナル」と入力）
2. 以下のコマンドを入力してEnter

```
python3 --version
```

---

## STEP 2：Streamlitのインストール

### ⚠️ インストール前の注意点

- **インターネット接続が必要**です（会社のWi-Fi環境によってはプロキシ設定が必要な場合あり）
- インストールには1〜3分かかります（PCの環境によって異なる）
- 会社PCの場合、セキュリティソフトが警告を出すことがありますが、Streamlitは公式の安全なライブラリです
- **一度インストールすれば次回以降は不要**です

### Windowsの場合

PowerShellを開いて以下を実行：

```
pip install streamlit
```

### Macの場合

ターミナルを開いて以下を実行：

```
pip3 install streamlit
```

### インストール確認

完了後、以下のコマンドでバージョンが表示されればインストール成功です：

```
streamlit --version
```

---

## STEP 3：作業フォルダの作成

アプリファイルを置くフォルダを作成します。

### Windowsの場合（PowerShellで実行）

```
mkdir E:\Antigravity\Streamlit
```

### Macの場合（ターミナルで実行）

```
mkdir ~/Documents/Streamlit
```

---

## STEP 4：app.py の作成

アプリの本体ファイル（app.py）を作成します。  
以下の内容をコピーして、メモ帳や VS Code などで `app.py` という名前で保存してください。

> **保存場所：** `E:\Antigravity\Streamlit\app.py`（Windowsの場合）

```python
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import csv
import io
from datetime import datetime, timedelta
from calendar import monthrange

FONT_NAME = "Meiryo UI"

st.set_page_config(page_title="ガントチャート自動生成", layout="wide")
st.title("ガントチャート自動生成ツール")
st.caption("CSVをアップロードしてExcel形式のガントチャートを作成します")

# サイドバー（設定）
with st.sidebar:
    st.header("設定")
    mode = st.selectbox("表示モード", ["daily（3ヶ月以内向け）", "weekly（年間向け）"])
    mode_key = "daily" if mode.startswith("daily") else "weekly"
    bar_color     = st.text_input("ガントバーの色（HEX）", value="4472C4")
    today_color   = st.text_input("今日の列の色（HEX）",   value="FFE699")
    weekend_color = st.text_input("土日の色（HEX・dailyのみ）", value="F2F2F2")

# サンプルCSVダウンロード
sample_csv = """タスク名,開始日,終了日,担当者
企画書作成,2026/06/01,2026/06/10,田中
デザイン確認,2026/06/08,2026/06/20,佐藤
システム開発,2026/06/15,2026/07/10,山田
テスト・修正,2026/07/08,2026/07/20,田中
リリース,2026/07/21,2026/07/25,
"""
st.download_button(
    label="サンプルCSVをダウンロード",
    data=sample_csv.encode("utf-8-sig"),
    file_name="sample_tasks.csv",
    mime="text/csv",
)

st.divider()

# CSVアップロード
uploaded_file = st.file_uploader("タスクCSVをアップロード（必須列：タスク名・開始日・終了日）", type="csv")

# ── ガントチャート生成ロジック（元スクリプトと同じ） ──

def parse_date(s):
    for fmt in ('%Y/%m/%d', '%Y-%m-%d'):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            continue
    raise ValueError(f"日付フォーマットが認識できません: {s}")

def load_tasks_from_upload(file):
    content = file.read().decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(content))
    tasks = []
    for row in reader:
        tasks.append({
            'task':  row['タスク名'],
            'start': parse_date(row['開始日']),
            'end':   parse_date(row['終了日']),
            'owner': row.get('担当者', ''),
        })
    return tasks

def get_monday(d):
    return d - timedelta(days=d.weekday())

def build_gantt_daily(ws, tasks, today, bar_color, today_color, weekend_color):
    min_date = min(t['start'] for t in tasks).replace(day=1)
    raw_max  = max(t['end']   for t in tasks)
    last_day = monthrange(raw_max.year, raw_max.month)[1]
    max_date = raw_max.replace(day=last_day)
    total_days = (max_date - min_date).days + 1
    COL_TASK  = 1
    COL_OWNER = 2
    COL_DATE0 = 3
    thin   = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.row_dimensions[1].height = 16
    for col, label in [(COL_TASK, "タスク名"), (COL_OWNER, "担当者")]:
        cell = ws.cell(row=1, column=col, value=label)
        cell.font      = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
        cell.fill      = PatternFill("solid", fgColor="2F5496")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = border
    month_start_col = COL_DATE0
    current_month   = None
    for i in range(total_days):
        d   = min_date + timedelta(days=i)
        col = COL_DATE0 + i
        ws.column_dimensions[get_column_letter(col)].width = 2.8
        if d.month != current_month:
            if current_month is not None:
                end_col = col - 1
                if month_start_col < end_col:
                    ws.merge_cells(start_row=1, start_column=month_start_col, end_row=1, end_column=end_col)
                prev_d = min_date + timedelta(days=i - 1)
                cell = ws.cell(row=1, column=month_start_col, value=f"{prev_d.year}年{prev_d.month}月")
                cell.font      = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=9)
                cell.fill      = PatternFill("solid", fgColor="2F5496")
                cell.alignment = Alignment(horizontal='center', vertical='center')
            current_month   = d.month
            month_start_col = col
    end_col = COL_DATE0 + total_days - 1
    if month_start_col < end_col:
        ws.merge_cells(start_row=1, start_column=month_start_col, end_row=1, end_column=end_col)
    last_d = min_date + timedelta(days=total_days - 1)
    cell = ws.cell(row=1, column=month_start_col, value=f"{last_d.year}年{last_d.month}月")
    cell.font      = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=9)
    cell.fill      = PatternFill("solid", fgColor="2F5496")
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 16
    for col in [COL_TASK, COL_OWNER]:
        cell = ws.cell(row=2, column=col)
        cell.fill   = PatternFill("solid", fgColor="2F5496")
        cell.border = border
    for i in range(total_days):
        d   = min_date + timedelta(days=i)
        col = COL_DATE0 + i
        cell = ws.cell(row=2, column=col, value=d.day)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = border
        is_weekend = d.weekday() >= 5
        is_today   = (d == today)
        if is_today:
            cell.fill = PatternFill("solid", fgColor=today_color)
            cell.font = Font(name=FONT_NAME, size=8, bold=True, color="333333")
        elif is_weekend:
            cell.fill = PatternFill("solid", fgColor="9DC3E6")
            cell.font = Font(name=FONT_NAME, size=8, color="FFFFFF")
        else:
            cell.fill = PatternFill("solid", fgColor="2F5496")
            cell.font = Font(name=FONT_NAME, size=8, color="FFFFFF")
    for row_idx, task in enumerate(tasks):
        row = 3 + row_idx
        ws.row_dimensions[row].height = 18
        cell = ws.cell(row=row, column=COL_TASK, value=task['task'])
        cell.font      = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(vertical='center')
        cell.border    = border
        cell = ws.cell(row=row, column=COL_OWNER, value=task['owner'])
        cell.font      = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = border
        for i in range(total_days):
            d   = min_date + timedelta(days=i)
            col = COL_DATE0 + i
            cell = ws.cell(row=row, column=col)
            cell.border = border
            in_range   = task['start'] <= d <= task['end']
            is_weekend = d.weekday() >= 5
            is_today   = (d == today)
            if in_range:
                cell.fill = PatternFill("solid", fgColor=bar_color)
            elif is_today:
                cell.fill = PatternFill("solid", fgColor=today_color)
            elif is_weekend:
                cell.fill = PatternFill("solid", fgColor=weekend_color)
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 10
    ws.freeze_panes = ws.cell(row=3, column=COL_DATE0)

def build_gantt_weekly(ws, tasks, today, bar_color, today_color):
    min_date = min(t['start'] for t in tasks).replace(day=1)
    raw_max  = max(t['end']   for t in tasks)
    last_day = monthrange(raw_max.year, raw_max.month)[1]
    max_date = raw_max.replace(day=last_day)
    week_start = get_monday(min_date)
    weeks = []
    while week_start <= max_date:
        weeks.append(week_start)
        week_start += timedelta(weeks=1)
    today_week = get_monday(today)
    COL_TASK  = 1
    COL_OWNER = 2
    COL_DATE0 = 3
    thin   = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.row_dimensions[1].height = 18
    for col, label in [(COL_TASK, "タスク名"), (COL_OWNER, "担当者")]:
        cell = ws.cell(row=1, column=col, value=label)
        cell.font      = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
        cell.fill      = PatternFill("solid", fgColor="2F5496")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = border
    month_groups = {}
    for i, w in enumerate(weeks):
        key = (w.year, w.month)
        month_groups.setdefault(key, []).append(COL_DATE0 + i)
    for (year, month), cols in month_groups.items():
        start_col = min(cols)
        end_col   = max(cols)
        if start_col < end_col:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        cell = ws.cell(row=1, column=start_col, value=f"{year}年{month}月")
        cell.font      = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
        cell.fill      = PatternFill("solid", fgColor="2F5496")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = border
    ws.row_dimensions[2].height = 16
    for col in [COL_TASK, COL_OWNER]:
        cell = ws.cell(row=2, column=col)
        cell.fill   = PatternFill("solid", fgColor="2F5496")
        cell.border = border
    for i, w in enumerate(weeks):
        col = COL_DATE0 + i
        ws.column_dimensions[get_column_letter(col)].width = 5.5
        cell = ws.cell(row=2, column=col, value=f"{w.month}/{w.day}")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = border
        if w == today_week:
            cell.fill = PatternFill("solid", fgColor=today_color)
            cell.font = Font(name=FONT_NAME, size=8, bold=True, color="333333")
        else:
            cell.fill = PatternFill("solid", fgColor="2F5496")
            cell.font = Font(name=FONT_NAME, size=8, color="FFFFFF")
    for row_idx, task in enumerate(tasks):
        row = 3 + row_idx
        ws.row_dimensions[row].height = 18
        cell = ws.cell(row=row, column=COL_TASK, value=task['task'])
        cell.font      = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(vertical='center')
        cell.border    = border
        cell = ws.cell(row=row, column=COL_OWNER, value=task['owner'])
        cell.font      = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = border
        for i, w in enumerate(weeks):
            col   = COL_DATE0 + i
            w_end = w + timedelta(days=6)
            cell  = ws.cell(row=row, column=col)
            cell.border = border
            if task['start'] <= w_end and task['end'] >= w:
                cell.fill = PatternFill("solid", fgColor=bar_color)
            elif w == today_week:
                cell.fill = PatternFill("solid", fgColor=today_color)
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 10
    ws.freeze_panes = ws.cell(row=3, column=COL_DATE0)

def generate_excel(tasks, mode_key, bar_color, today_color, weekend_color):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ガントチャート"
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    if mode_key == "weekly":
        build_gantt_weekly(ws, tasks, today, bar_color, today_color)
    else:
        build_gantt_daily(ws, tasks, today, bar_color, today_color, weekend_color)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# メイン処理
if uploaded_file is not None:
    try:
        tasks = load_tasks_from_upload(uploaded_file)
        st.success(f"{len(tasks)} 件のタスクを読み込みました")
        import pandas as pd
        df = pd.DataFrame([{
            'タスク名': t['task'],
            '開始日': t['start'].strftime('%Y/%m/%d'),
            '終了日': t['end'].strftime('%Y/%m/%d'),
            '担当者': t['owner'],
        } for t in tasks])
        st.dataframe(df, use_container_width=True)
        if st.button("ガントチャートを生成する", type="primary"):
            with st.spinner("生成中..."):
                excel_data = generate_excel(tasks, mode_key, bar_color, today_color, weekend_color)
            st.success("生成完了！")
            st.download_button(
                label="Excelをダウンロード",
                data=excel_data,
                file_name="gantt_chart.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
        st.info("CSVの列名（タスク名・開始日・終了日）と日付フォーマット（YYYY/MM/DD）を確認してください")
```

---

## STEP 5：アプリの起動方法

### Windowsの場合
1. PowerShellを開く
2. 以下のコマンドを実行

```
cd E:\Antigravity\Streamlit
streamlit run app.py
```

### Macの場合
1. ターミナルを開く
2. 以下のコマンドを実行

```
cd ~/Documents/Streamlit
streamlit run app.py
```

### 起動後の動き
- 自動的にブラウザが開き、アプリが表示される
- 表示されない場合はブラウザで `http://localhost:8501` を開く
- **PowerShell/ターミナルは閉じないこと**（閉じるとアプリが止まる）
- 終了したいときは PowerShell/ターミナルで `Ctrl + C` を押す

---

## STEP 6：アプリの使い方

1. 左サイドバーで表示モード・色を設定（デフォルトのままでもOK）
2. 「サンプルCSVをダウンロード」で雛形を取得（初回のみ）
3. CSVにタスク情報を入力して保存
4. 「タスクCSVをアップロード」エリアにCSVをドラッグ＆ドロップ
5. タスク一覧がプレビュー表示される
6. 「ガントチャートを生成する」ボタンをクリック
7. 「Excelをダウンロード」ボタンでファイルを取得

---

## STEP 7：Streamlit Cloudでデプロイ（社外公開・社内URL共有）

> ここまでの手順はすべてローカル（自分のPC）での動作。  
> 他の人がブラウザからアクセスできるようにするにはデプロイが必要。

### 必要なもの
- GitHubアカウント（無料）
- Streamlit Cloudアカウント（無料・GitHubアカウントでログイン）

### 手順

**① GitHubにリポジトリを作成**
1. github.com にログイン
2. 「New repository」でリポジトリを作成（名前例：`gantt-chart-app`）
3. 公開設定：Public（無料プランで公開デプロイする場合）

**② ファイルをGitHubにアップロード**
- `app.py` をリポジトリにアップロード
- `requirements.txt` も一緒に作成してアップロード（必須）

**requirements.txtの内容：**
```
streamlit
openpyxl
pandas
```

**③ Streamlit Cloudでデプロイ**
1. share.streamlit.io にアクセス
2. GitHubアカウントでログイン
3. 「New app」→ 該当リポジトリ・ブランチ・app.py を選択
4. 「Deploy!」ボタンをクリック
5. 数分でデプロイ完了 → URLが発行される

**④ 社内WebサイトへiframeでURL埋め込み**

発行されたURLを以下のように社内Webページに貼り付ける：

```html
<iframe 
  src="https://[あなたのアプリURL].streamlit.app" 
  width="100%" 
  height="800px" 
  frameborder="0">
</iframe>
```

---

## STEP 8：アプリを他の人と共有する方法

> 「共有する」には2つのレベルがあります。目的に合わせて選んでください。

---

### 方法A：同じWi-Fi（社内LAN）内だけで共有する

**特徴：** GitHubなど不要。社内の同じネットワークにいる人だけがアクセスできる。

**手順：**

1. 自分のPCのIPアドレスを調べる

   Windowsの場合、PowerShellで以下を実行：
   ```
   ipconfig
   ```
   「IPv4アドレス」に書いてある数字（例：`192.168.1.10`）がIPアドレス

2. 以下のコマンドでStreamlitを起動する（通常の起動コマンドとは少し違う）
   ```
   streamlit run app.py --server.address 0.0.0.0
   ```

3. 同じWi-Fiにつながっている人に以下のURLを教える
   ```
   http://192.168.1.10:8501
   ```
   （数字の部分は自分のIPアドレスに置き換える）

**注意点：**
- 自分のPCがオン・かつStreamlitが起動中でないとアクセスできない
- 社外からはアクセスできない
- 会社のネットワーク設定によっては繋がらない場合もある

---

### 方法B：インターネット経由でどこからでも共有する（Streamlit Cloud）

**特徴：** 専用URLが発行され、誰でもどこからでもアクセス可能。無料。

**必要なもの：**
- GitHubアカウント（無料）→ github.com で登録
- Streamlit Cloudアカウント（無料）→ GitHubアカウントでそのままログインできる

**⚠️ 注意点：**
- GitHubにアップするファイルは**Public（公開）リポジトリ**にする必要がある（無料プランの場合）
- コードが公開されるため、**個人情報・社内機密データは含めない**こと
- アプリ自体のURL（ツールを使う画面）は誰でも開けるが、コードを読まれるリスクは低い

**手順：**

**① GitHubにリポジトリを作成**
1. github.com にログイン（アカウントがなければ無料登録）
2. 右上の「＋」→「New repository」をクリック
3. Repository name に名前を入力（例：`jimu-tools`）
4. 「Public」を選択
5. 「Create repository」をクリック

**② ファイルをGitHubにアップロード**
1. 作成したリポジトリのページを開く
2. 「Add file」→「Upload files」をクリック
3. 以下のファイルをアップロード：
   - `app.py`
   - `requirements.txt`（下記の内容で新規作成してアップ）

**requirements.txt の内容（メモ帳で作成してOK）：**
```
streamlit
openpyxl
pandas
```

**③ Streamlit Cloudでデプロイ**
1. share.streamlit.io にアクセス
2. 「Continue with GitHub」でログイン
3. 「Create app」→「Deploy a public app from GitHub」
4. リポジトリ・ブランチ（main）・ファイル（app.py）を選択
5. App URL を好きな名前に設定（例：`jimu-gantt-chart`）
6. 「Deploy!」をクリック
7. 数分待つと専用URLが発行される（例：`https://jimu-gantt-chart.streamlit.app`）

**④ URLを共有するだけで誰でも使える**
- そのURLをSlackやメールで送るだけでOK
- 社内Webサイトにiframeで埋め込むこともできる：
  ```html
  <iframe 
    src="https://jimu-gantt-chart.streamlit.app" 
    width="100%" 
    height="800px" 
    frameborder="0">
  </iframe>
  ```

---

## STEP 9：複数アプリをまとめたポートフォリオサイトを作る

> ガントチャートだけでなく、PDF結合・グラフ生成・勤怠集計なども  
> 1つのWebサイトにまとめて「業務効率化ツール集」として公開する方法。

---

### 仕組み：Streamlitの「多ページ機能」を使う

Streamlitには**1つのアプリの中に複数のページを作れる機能**があります。  
トップページ（メニュー画面）からツールを選べる、まるでWebサービスのような構成にできます。

---

### 完成イメージ

```
【トップページ】
  業務効率化ツール集
  ━━━━━━━━━━━━━━━━━━
  📊 ガントチャート自動生成
  📄 PDF結合・ページ抽出
  📈 グラフ自動生成
  📅 勤怠データ集計
  ━━━━━━━━━━━━━━━━━━
  ← 左のメニューから選んでください
```

左サイドバーに全ツールがリストアップされ、クリックするだけで切り替えられる。

---

### フォルダ構成（多ページ対応版）

```
E:\Antigravity\Streamlit\
├── app.py                      ← トップページ（メニュー・説明画面）
├── requirements.txt            ← 必要ライブラリ一覧
├── pages\                      ← この中にツールを1ファイル1ページで追加していく
│   ├── 1_ガントチャート.py     ← 現在のapp.pyの内容をここに移す
│   ├── 2_PDF結合.py            ← 次に追加予定
│   ├── 3_グラフ自動生成.py     ← その次
│   └── 4_勤怠集計.py           ← さらにその次
└── 仕様書_ガントチャートStreamlit.md
```

**ポイント：**
- `pages\` フォルダの中に置いたPythonファイルが自動でメニューに追加される
- ファイル名の先頭の数字が並び順になる（`1_〇〇.py` → `2_〇〇.py` の順）

---

### 多ページ対応への変更手順

**① `pages` フォルダを作成する**

PowerShellで実行：
```
mkdir E:\Antigravity\Streamlit\pages
```

**② 現在のapp.pyをガントチャートページとしてコピーする**

PowerShellで実行：
```
copy E:\Antigravity\Streamlit\app.py E:\Antigravity\Streamlit\pages\1_ガントチャート.py
```

**③ app.py をトップページ（メニュー画面）に書き換える**

app.py の中身を以下に差し替える：

```python
import streamlit as st

st.set_page_config(page_title="業務効率化ツール集", layout="wide")

st.title("業務効率化ツール集")
st.subheader("Pythonで作った業務自動化ツールをまとめています")
st.write("左のメニューからツールを選んでください。")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 ガントチャート自動生成")
    st.write("CSVをアップロードするだけでExcel形式のガントチャートを自動作成。dailyモード・weeklyモード対応。")

    st.markdown("### 📄 PDF結合・ページ抽出")
    st.write("複数のPDFを1つに結合、または必要なページだけを抽出できます。（準備中）")

with col2:
    st.markdown("### 📈 グラフ自動生成")
    st.write("Excelデータを読み込んで棒グラフ・折れ線グラフを自動生成。（準備中）")

    st.markdown("### 📅 勤怠データ集計")
    st.write("勤怠CSVを読み込んで月次レポートを自動作成。（準備中）")

st.divider()
st.caption("© 2026 まるのまい｜事務×AI効率化ツール")
```

**④ 起動して確認する**

```
streamlit run app.py
```

左サイドバーに「ガントチャート」が表示され、クリックで切り替えられれば成功！

---

### ツール追加のたびにやること

新しいツール（例：PDF結合）を追加するときは：

1. `pages\` フォルダに `2_PDF結合.py` を作成してコードを書く
2. `app.py` のトップページにそのツールの説明を追記する
3. `requirements.txt` に必要なライブラリを追記する（PDF結合なら `pypdf` など）
4. GitHubにアップ → Streamlit Cloudが自動で更新される

---

### 今後のツール追加予定

| ページ番号 | ツール名 | 元スクリプト | 状態 |
|---|---|---|---|
| 1 | ガントチャート自動生成 | gantt_chart.py | ✅ 完成 |
| 2 | PDF結合・ページ抽出 | merge_pdf.py / extract_pages.py | 未着手 |
| 3 | グラフ自動生成 | graph_report.py | 未着手 |
| 4 | 勤怠データ集計 | kintai_report.py | 未着手 |

---

## よくあるトラブルと対処法

| 症状 | 原因 | 対処法 |
|---|---|---|
| `streamlit: command not found` | インストールが完了していない | `pip install streamlit` を再実行 |
| ブラウザが自動で開かない | 環境によっては開かないことがある | ブラウザで `http://localhost:8501` を開く |
| CSVアップロードでエラーになる | 列名が違う・日付フォーマットが違う | 列名「タスク名・開始日・終了日」、日付「2026/06/01」形式で確認 |
| 文字化けする | CSVの文字コードが違う | ExcelでCSV保存するときは「UTF-8（BOM付き）」を選択 |
| Excelを開いたらフォントが変 | Meiryo UIが入っていないPC | Windowsでは標準搭載。Macでは別フォントに見える場合あり |

---

## 開発スケジュール（実績・今後）

### フェーズ1：ガントチャート単体アプリ

| ステップ | 内容 | 状態 |
|---|---|---|
| 1 | Streamlitインストール | ✅ 完了（2026-06-05） |
| 2 | app.py作成 | ✅ 完了（2026-06-05） |
| 3 | ローカルで動作テスト | ✅ 完了（2026-06-05） |
| 4 | GitHubにリポジトリ作成・アップロード | 未着手 |
| 5 | Streamlit Cloudでデプロイ・URL取得 | 未着手 |
| 6 | 社内WebサイトへURL埋め込み | 未着手 |

### フェーズ2：ポートフォリオサイト化（まとめサイト）

| ステップ | 内容 | 状態 |
|---|---|---|
| 7 | pagesフォルダ作成・多ページ対応に構成変更 | 未着手 |
| 8 | トップページ（メニュー画面）作成 | 未着手 |
| 9 | PDF結合ツールをページとして追加 | 未着手 |
| 10 | グラフ自動生成ツールをページとして追加 | 未着手 |
| 11 | 勤怠集計ツールをページとして追加 | 未着手 |
| 12 | GitHub更新 → Streamlit Cloud自動反映確認 | 未着手 |
