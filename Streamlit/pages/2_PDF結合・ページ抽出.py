import streamlit as st
import io
from pypdf import PdfWriter, PdfReader

st.page_link("home.py", label="← ホームに戻る")
st.title("📄 PDF結合・ページ抽出")
st.caption("複数のPDFを1つにまとめたり、必要なページだけ取り出せます")

tab1, tab2 = st.tabs(["📎 PDF結合", "✂️ ページ抽出"])

# ── Tab1: PDF結合 ──────────────────────────────
with tab1:
    st.subheader("複数のPDFを1つに結合する")

    uploaded_files = st.file_uploader(
        "PDFファイルをアップロード（複数選択OK・アップした順に結合されます）",
        type="pdf",
        accept_multiple_files=True,
        key="merge_upload"
    )

    if uploaded_files:
        st.write(f"**{len(uploaded_files)} 件のPDFを読み込みました（結合順）**")
        for i, f in enumerate(uploaded_files, 1):
            st.text(f"  {i}. {f.name}")

        if st.button("PDFを結合する", type="primary", key="merge_btn"):
            with st.spinner("作成中です。しばらくお待ちください…"):
                writer = PdfWriter()
                for f in uploaded_files:
                    writer.append(PdfReader(f))
                output = io.BytesIO()
                writer.write(output)
                output.seek(0)

            st.success(f"完成しました！{len(uploaded_files)} 件のPDFを結合しました")
            st.download_button(
                label="結合PDFをダウンロード",
                data=output,
                file_name="merged.pdf",
                mime="application/pdf",
                key="merge_download"
            )

    with st.expander("📖 使い方"):
        st.write("1. 「Browse files」またはドラッグ＆ドロップで複数のPDFをアップロード")
        st.write("2. アップロードした順番に結合されます（順番を変えたい場合は一度クリアして入れ直してください）")
        st.write("3. 「PDFを結合する」ボタンを押してダウンロード")

# ── Tab2: ページ抽出 ───────────────────────────
with tab2:
    st.subheader("PDFから必要なページだけ取り出す")

    uploaded_file = st.file_uploader(
        "PDFファイルをアップロード",
        type="pdf",
        key="extract_upload"
    )

    if uploaded_file:
        reader = PdfReader(uploaded_file)
        total = len(reader.pages)
        st.info(f"総ページ数：{total} ページ")

        page_input = st.text_input(
            "抽出するページ番号をカンマ区切りで入力",
            placeholder=f"例：1, 3, 5　（最大 {total} ページ）"
        )

        if page_input:
            try:
                pages = [int(p.strip()) for p in page_input.split(",") if p.strip()]
                valid_pages = [p for p in pages if 1 <= p <= total]
                invalid_pages = [p for p in pages if p < 1 or p > total]

                if invalid_pages:
                    st.warning(f"範囲外のページをスキップします：{invalid_pages}")

                if valid_pages:
                    st.write(f"抽出対象：{valid_pages}（{len(valid_pages)} ページ）")

                    if st.button("ページを抽出する", type="primary", key="extract_btn"):
                        with st.spinner("作成中です。しばらくお待ちください…"):
                            uploaded_file.seek(0)
                            reader2 = PdfReader(uploaded_file)
                            writer = PdfWriter()
                            for p in valid_pages:
                                writer.add_page(reader2.pages[p - 1])
                            output = io.BytesIO()
                            writer.write(output)
                            output.seek(0)

                        st.success(f"完成しました！{len(valid_pages)} ページを抽出しました")
                        st.download_button(
                            label="抽出PDFをダウンロード",
                            data=output,
                            file_name="extracted_pages.pdf",
                            mime="application/pdf",
                            key="extract_download"
                        )
                else:
                    st.error(f"有効なページ番号がありません。1〜{total} の範囲で入力してください。")

            except ValueError:
                st.error("ページ番号は数字をカンマ区切りで入力してください（例：1, 2, 3）")

    with st.expander("📖 使い方"):
        st.write("1. PDFをアップロードすると総ページ数が表示されます")
        st.write("2. 抽出したいページ番号をカンマ区切りで入力（例：1, 3, 5）")
        st.write("3. 「ページを抽出する」ボタンを押してダウンロード")
