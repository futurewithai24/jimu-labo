import streamlit as st
import base64
from pathlib import Path

# ── ヒーローセクション ──────────────────────────
_img = base64.b64encode(Path("assets/unloop-logo-wide.png").read_bytes()).decode()
st.markdown(
    f'<div style="text-align:center;padding:8px 0 0">'
    f'<div style="display:inline-block;background:#eaf1fb;border-radius:20px;padding:16px 28px;">'
    f'<img src="data:image/png;base64,{_img}" width="280" style="display:block;">'
    f'</div>'
    f'</div>',
    unsafe_allow_html=True
)
st.markdown("<p style='text-align:center; font-size:1.2rem; font-weight:bold; margin-top:32px;'>ファイル作業、もっとラクにしよう。</p>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;'>CSVをアップして、ボタンを押すだけ。面倒な事務作業を自動化するツール集です。</p>", unsafe_allow_html=True)

st.divider()

# ── 使い方（3ステップ） ─────────────────────────
st.subheader("使い方は3ステップだけ")
col1, col2, col3 = st.columns(3)

with col1:
    with st.container(border=True):
        st.markdown("#### 1️⃣ ツールを選ぶ")
        st.write("左メニューから使いたいツールをクリック")

with col2:
    with st.container(border=True):
        st.markdown("#### 2️⃣ ファイルをアップ")
        st.write("CSVやExcel・PDFをドラッグ＆ドロップ")

with col3:
    with st.container(border=True):
        st.markdown("#### 3️⃣ ダウンロードして完成")
        st.write("ボタンを押すだけでファイルが完成")

st.divider()

# ── ツール一覧 ──────────────────────────────────
st.subheader("ツール一覧")
col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.markdown("### 📊 ガントチャート自動生成")
        st.write("タスク一覧のCSVをアップするだけで、Excel形式のガントチャートを自動作成。日単位・週単位に対応。")
        st.page_link("pages/1_ガントチャート.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 📅 勤怠データ集計")
        st.write("勤怠Excelを読み込んで月次レポートを自動作成。集計ミス・転記ミスをなくします。")
        st.page_link("pages/3_勤怠集計.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 🖼️ 画像一括リサイズ")
        st.write("複数の画像をまとめてリサイズ。縦横比を保ったまま幅を指定するだけでOK。")
        st.page_link("pages/4_画像リサイズ.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 💴 経費集計")
        st.write("日々の経費を入力してカテゴリ別に自動集計。Excelでダウンロードできます。")
        st.page_link("pages/8_経費集計.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 📦 在庫チェック自動化")
        st.write("入出庫記録をExcelに反映して、基準値以下の品目を赤くハイライト。二重計算防止つき。")
        st.page_link("pages/12_在庫チェック.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 📁 フォルダ構成ZIP作成")
        st.write("プロジェクト名とフォルダ構成を入力するだけで、フォルダ構造入りのZIPをダウンロード。")
        st.page_link("pages/13_フォルダ構成ZIP作成.py", label="使ってみる →")

with col2:
    with st.container(border=True):
        st.markdown("### 📄 PDF結合・ページ抽出")
        st.write("複数のPDFを1つにまとめる・必要なページだけ取り出す。手作業でのコピペから解放されます。")
        st.page_link("pages/2_PDF結合・ページ抽出.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 📝 案内状一括生成")
        st.write("送付先リストのExcelをアップするだけで、人数分のWord案内状を自動作成してZIPでダウンロード。")
        st.page_link("pages/5_案内状一括生成.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 📋 Excel一括集計")
        st.write("列名がバラバラな複数のExcelをまとめて1ファイルに集約。列名の統一も自動対応。")
        st.page_link("pages/6_Excel一括集計.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 🔏 画像透かし一括追加")
        st.write("複数の画像にテキストやロゴをまとめて追加してZIPでダウンロード。背景透過PNG対応。")
        st.page_link("pages/9_画像透かし追加.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 🔒 パスワード付きZIP作成")
        st.write("複数ファイルをまとめてAES暗号化ZIPに圧縮。パスワードを設定してダウンロードできます。")
        st.page_link("pages/10_パスワード付きZIP作成.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 👥 顧客リスト重複チェック")
        st.write("会社名などのキー列を基に重複行を検出。全角・半角ゆれも自動吸収して黄色ハイライト。")
        st.page_link("pages/11_顧客リスト重複チェック.py", label="使ってみる →")

    with st.container(border=True):
        st.markdown("### 📰 キーワードニュース収集")
        st.write("キーワードを入れるだけでGoogle ニュースから最新記事を自動収集してCSVでダウンロード。")
        st.page_link("pages/14_キーワードニュース収集.py", label="使ってみる →")

st.divider()

# ── こんな人に使ってほしい ──────────────────────
st.subheader("こんな人に使ってほしい")
col1, col2, col3 = st.columns(3)
with col1:
    st.info("📅　Excelの作業に毎回時間がかかる")
with col2:
    st.info("📄　PDFの結合・分割が手作業")
with col3:
    st.info("📝　同じ書類を何十枚も手で作成")

st.divider()
st.caption("© 2026 UNLOOP｜業務効率化ツール集")
