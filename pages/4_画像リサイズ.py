import streamlit as st
from PIL import Image
import io
import zipfile

from utils import show_logo
show_logo()
st.page_link("home.py", label="← ホームに戻る")
st.title("🖼️ 画像一括リサイズ")
st.caption("複数の画像をまとめてリサイズしてZIPでダウンロードできます")

# ── サイドバー（設定） ──────────────────────────
with st.sidebar:
    st.header("設定")
    target_width = st.number_input("リサイズ後の幅（px）", min_value=100, max_value=5000, value=800, step=100)
    quality      = st.slider("JPEG画質（1〜95）", min_value=1, max_value=95, value=85)
    st.caption("PNG画像は画質設定の影響を受けません")

# ── サンプル画像ダウンロード ──────────────────
def _make_sample_zip():
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as zf:
        for name, color in [("sample_blue.png",(70,130,180)), ("sample_green.png",(60,179,113)), ("sample_orange.png",(255,140,0))]:
            img = Image.new("RGB", (1600, 900), color=color)
            buf = io.BytesIO()
            img.save(buf, "PNG")
            buf.seek(0)
            zf.writestr(name, buf.read())
    zip_buf.seek(0)
    return zip_buf

st.download_button(
    label="サンプル画像をダウンロード（3枚入りZIP）",
    data=_make_sample_zip(),
    file_name="サンプル画像.zip",
    mime="application/zip"
)
st.caption("※ 1600×900pxのカラー画像が3枚入っています。リサイズのお試しにどうぞ。")

st.divider()

# ── ファイルアップロード ────────────────────────
uploaded_files = st.file_uploader(
    "画像をアップロード（JPG・PNG・WEBP対応・複数選択OK）",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"**{len(uploaded_files)} 件の画像を読み込みました**")

    # プレビュー表示（最大4件）
    preview_files = uploaded_files[:4]
    cols = st.columns(len(preview_files))
    for col, f in zip(cols, preview_files):
        img = Image.open(f)
        col.image(img, caption=f"{f.name}　{img.width}×{img.height}px", use_container_width=True)
    if len(uploaded_files) > 4:
        st.caption(f"… 他 {len(uploaded_files) - 4} 件")

    if st.button("一括リサイズする", type="primary"):
        with st.spinner("作成中です。しばらくお待ちください…"):
            zip_buffer = io.BytesIO()
            results = []

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in uploaded_files:
                    try:
                        f.seek(0)
                        img = Image.open(f)
                        orig_w, orig_h = img.width, img.height

                        ratio      = target_width / orig_w
                        new_height = int(orig_h * ratio)
                        resized    = img.resize((target_width, new_height), Image.LANCZOS)

                        img_buffer = io.BytesIO()
                        suffix = f.name.rsplit(".", 1)[-1].lower()
                        if suffix == "png":
                            resized.save(img_buffer, "PNG")
                        else:
                            if resized.mode in ("RGBA", "P"):
                                resized = resized.convert("RGB")
                            resized.save(img_buffer, "JPEG", quality=quality)

                        img_buffer.seek(0)
                        zf.writestr(f.name, img_buffer.read())
                        results.append({"ファイル名": f.name, "元サイズ": f"{orig_w}×{orig_h}px", "変換後": f"{target_width}×{new_height}px", "結果": "✅ 完了"})

                    except Exception as e:
                        results.append({"ファイル名": f.name, "元サイズ": "-", "変換後": "-", "結果": f"❌ エラー：{e}"})

            zip_buffer.seek(0)

        st.success(f"完成しました！{len(uploaded_files)} 件をリサイズしました")

        import pandas as pd
        st.dataframe(pd.DataFrame(results), use_container_width=True)

        st.download_button(
            label="ZIPをダウンロード",
            data=zip_buffer,
            file_name="resized_images.zip",
            mime="application/zip"
        )

with st.expander("📖 使い方"):
    st.write("1. 画像をアップロード（JPG・PNG・WEBPに対応、複数まとめてOK）")
    st.write("2. 左のサイドバーでリサイズ後の幅・JPEG画質を設定")
    st.write("3. 「一括リサイズする」ボタンを押してZIPをダウンロード")
    st.write("- 縦横比は自動で維持されます（幅を指定すると高さが自動計算）")
    st.write("- 元の画像ファイルは変更されません")
