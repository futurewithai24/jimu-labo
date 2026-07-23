import streamlit as st
import zipfile
import io
from datetime import date

from utils import show_logo
show_logo()
st.page_link("home.py", label="← ホームに戻る")
st.title("📁 フォルダ構成ZIP作成")
st.caption("プロジェクト名とフォルダ構成を入力するだけで、フォルダ構造入りのZIPをダウンロードできます")

with st.sidebar:
    st.header("設定")
    use_date = st.checkbox("プロジェクト名に今日の日付を付ける", value=True)

today = date.today().strftime("%Y%m%d")

project_name = st.text_input(
    "プロジェクト名",
    value="新規案件_サンプル",
    placeholder="例：2026_新規案件"
)
if use_date:
    full_name = f"{today}_{project_name}"
else:
    full_name = project_name

st.caption(f"作成されるフォルダ名：`{full_name}`")

default_structure = """01_受信
02_作業中
03_完了
04_提出済み
05_参考資料"""

folder_input = st.text_area(
    "フォルダ構成（1行に1フォルダ。サブフォルダは / で区切る）",
    value=default_structure,
    height=200,
    placeholder="01_受信\n02_作業中\n02_作業中/メール\n03_完了"
)

st.caption("例：`02_作業中/メール` と書くとサブフォルダも作れます")

if st.button("フォルダ構成ZIPを作成する", type="primary"):
    folders = [line.strip() for line in folder_input.strip().splitlines() if line.strip()]

    if not folders:
        st.warning("フォルダ構成を入力してください")
    else:
        with st.spinner("処理中です…"):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for folder in folders:
                    path = f"{full_name}/{folder}/.gitkeep"
                    zf.writestr(path, "")
            zip_buf.seek(0)

        st.success(f"完成しました！{len(folders)} 個のフォルダを作成しました")
        st.write("**作成されるフォルダ一覧：**")
        for f in folders:
            st.write(f"・{full_name}/{f}/")

        st.download_button(
            label="ZIPをダウンロード",
            data=zip_buf,
            file_name=f"{full_name}.zip",
            mime="application/zip",
        )

with st.expander("📖 使い方"):
    st.write("1. プロジェクト名を入力する")
    st.write("2. フォルダ構成を1行1フォルダで入力する")
    st.write("3. 「フォルダ構成ZIPを作成する」ボタンを押す")
    st.write("4. ZIPをダウンロードして展開すると、フォルダ構造がそのまま作られます")
    st.write("- サブフォルダは `02_作業中/メール` のように / で区切って書く")
    st.write("- 各フォルダに `.gitkeep` ファイルが入ります（空フォルダを保持するため）")
