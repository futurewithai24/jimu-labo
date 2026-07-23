import streamlit as st
import pyzipper
import io

from utils import show_logo
show_logo()
st.page_link("home.py", label="← ホームに戻る")
st.title("🔒 パスワード付きZIP作成")
st.caption("複数ファイルをまとめてパスワード付きZIPに圧縮してダウンロードできます")

with st.sidebar:
    st.header("設定")
    password = st.text_input("パスワード", type="password", placeholder="半角英数字推奨")
    zip_name = st.text_input("ZIPファイル名", value="protected.zip")

uploaded_files = st.file_uploader(
    "圧縮したいファイルをアップロード（複数選択OK）",
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"**{len(uploaded_files)} 件のファイルを読み込みました**")
    for f in uploaded_files:
        st.caption(f"・{f.name}")

    if not password:
        st.warning("サイドバーでパスワードを設定してください")
    else:
        if st.button("パスワード付きZIPを作成する", type="primary"):
            with st.spinner("処理中です…"):
                zip_buf = io.BytesIO()
                with pyzipper.AESZipFile(
                    zip_buf, "w",
                    compression=pyzipper.ZIP_LZMA,
                    encryption=pyzipper.WZ_AES,
                ) as zf:
                    zf.setpassword(password.encode("utf-8"))
                    for f in uploaded_files:
                        f.seek(0)
                        zf.writestr(f.name, f.read())
                zip_buf.seek(0)

            st.success(f"完成しました！{len(uploaded_files)} 件を圧縮しました")
            if not zip_name.endswith(".zip"):
                zip_name += ".zip"
            st.download_button(
                label="ZIPをダウンロード",
                data=zip_buf,
                file_name=zip_name,
                mime="application/zip",
            )

with st.expander("📖 使い方"):
    st.write("1. サイドバーでパスワードとZIPファイル名を設定する")
    st.write("2. 圧縮したいファイルを複数まとめてアップロードする")
    st.write("3. 「パスワード付きZIPを作成する」ボタンを押してダウンロード")
    st.write("- AES暗号化（WinZip・7-Zip・macOS対応）")
    st.write("- 元ファイルは変更されません")
