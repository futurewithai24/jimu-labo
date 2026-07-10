import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import os

st.page_link("home.py", label="← ホームに戻る")
st.title("🔏 画像透かし一括追加")
st.caption("複数の画像にテキストやロゴをまとめて追加してZIPでダウンロードできます")


def get_font(size):
    candidates = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttf",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        "C:/Windows/Fonts/meiryo.ttc",
        "C:/Windows/Fonts/msgothic.ttc",
        "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return ImageFont.load_default()


def calc_position(img_size, elem_size, position_key, margin=20):
    w, h = img_size
    ew, eh = elem_size
    return {
        "bottom_right": (w - ew - margin, h - eh - margin),
        "bottom_left":  (margin,           h - eh - margin),
        "top_right":    (w - ew - margin,  margin),
        "top_left":     (margin,           margin),
        "center":       ((w - ew) // 2,    (h - eh) // 2),
    }.get(position_key, (w - ew - margin, h - eh - margin))


def add_text_watermark(img, text, position_key, opacity, font_size):
    img = img.convert("RGBA")
    overlay = Image.new("RGBA", img.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)
    font = get_font(font_size)
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    except AttributeError:
        tw, th = draw.textsize(text, font=font)
    x, y = calc_position(img.size, (tw, th), position_key)
    draw.text((x + 2, y + 2), text, font=font, fill=(0, 0, 0, opacity))
    draw.text((x,     y),     text, font=font, fill=(255, 255, 255, opacity))
    return Image.alpha_composite(img, overlay).convert("RGB")


def add_logo_watermark(img, logo, position_key, opacity, scale):
    img  = img.convert("RGBA")
    logo = logo.convert("RGBA")
    lw   = int(img.width * scale)
    lh   = int(logo.height * (lw / logo.width))
    logo = logo.resize((lw, lh), Image.LANCZOS)
    r, g, b, a = logo.split()
    a = a.point(lambda v: int(v * opacity / 255))
    logo.putalpha(a)
    x, y   = calc_position(img.size, (lw, lh), position_key)
    result = img.copy()
    result.paste(logo, (x, y), logo)
    return result.convert("RGB")


# ── サイドバー ──────────────────────────────────
with st.sidebar:
    st.header("設定")
    mode = st.radio("透かしの種類", ["テキスト", "ロゴ画像"])
    if mode == "テキスト":
        wm_text    = st.text_input("透かしテキスト", value="© まるのまい")
        font_size  = st.slider("文字サイズ（px）", 16, 120, 36)
    position_options = {
        "右下": "bottom_right",
        "左下": "bottom_left",
        "右上": "top_right",
        "左上": "top_left",
        "中央": "center",
    }
    pos_label  = st.selectbox("表示位置", list(position_options.keys()))
    pos_value  = position_options[pos_label]
    opacity    = st.slider("透明度（低いほど薄い）", 50, 255, 180)
    if mode == "ロゴ画像":
        logo_pct = st.slider("ロゴのサイズ（画像幅に対する%）", 5, 40, 15)

# ── メインエリア ────────────────────────────────
logo_img = None
if mode == "ロゴ画像":
    logo_file = st.file_uploader(
        "ロゴ画像（背景透過PNG推奨）", type=["png", "jpg", "jpeg"]
    )
    if logo_file:
        logo_img = Image.open(logo_file)
    else:
        st.info("ロゴ画像をアップロードしてください")

uploaded_files = st.file_uploader(
    "画像をアップロード（JPG・PNG・WEBP対応・複数選択OK）",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"**{len(uploaded_files)} 件の画像を読み込みました**")
    preview = uploaded_files[:4]
    cols = st.columns(len(preview))
    for col, f in zip(cols, preview):
        col.image(Image.open(f), caption=f.name, use_container_width=True)
    if len(uploaded_files) > 4:
        st.caption(f"… 他 {len(uploaded_files) - 4} 件")

    ready = mode == "テキスト" or (mode == "ロゴ画像" and logo_img is not None)

    if ready and st.button("透かしを一括追加する", type="primary"):
        with st.spinner("処理中です…"):
            zip_buf = io.BytesIO()
            results = []
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in uploaded_files:
                    try:
                        f.seek(0)
                        img = Image.open(f)
                        if mode == "テキスト":
                            out = add_text_watermark(img, wm_text, pos_value, opacity, font_size)
                        else:
                            out = add_logo_watermark(img, logo_img, pos_value, opacity, logo_pct / 100)
                        buf  = io.BytesIO()
                        ext  = f.name.rsplit(".", 1)[-1].lower()
                        if ext == "png":
                            out.save(buf, "PNG")
                        else:
                            out.save(buf, "JPEG", quality=90)
                        buf.seek(0)
                        zf.writestr(f.name, buf.read())
                        results.append({"ファイル名": f.name, "結果": "✅ 完了"})
                    except Exception as e:
                        results.append({"ファイル名": f.name, "結果": f"❌ {e}"})
            zip_buf.seek(0)

        st.success(f"完成しました！{len(uploaded_files)} 件を処理しました")
        import pandas as pd
        st.dataframe(pd.DataFrame(results), use_container_width=True)
        st.download_button(
            label="ZIPをダウンロード",
            data=zip_buf,
            file_name="watermarked_images.zip",
            mime="application/zip",
        )

with st.expander("📖 使い方"):
    st.write("1. サイドバーで透かしの設定をする（テキスト or ロゴ画像）")
    st.write("2. ロゴ画像モードの場合はロゴPNGをアップロード")
    st.write("3. 透かしを入れたい画像を複数まとめてアップロード")
    st.write("4. 「透かしを一括追加する」ボタンを押してZIPをダウンロード")
    st.write("- 元画像は変更されません")
    st.write("- テキスト透かしは白文字＋黒い影で視認しやすく表示されます")
