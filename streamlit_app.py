import streamlit as st

st.set_page_config(
    page_title="UNLOOP｜業務効率化ツール集",
    page_icon="🗂️",
    layout="wide"
)

st.markdown("""
<style>
    /* ── フォントサイズ ── */
    h1 { font-size: 28px !important; }        /* アプリ名（UNLOOP） */
    h3 { font-size: 16px !important; }        /* ツール一覧の名前 */
    h4 { font-size: 14px !important; }        /* 使い方3ステップの見出し */
    p, .stMarkdown p { font-size: 14px !important; }  /* 本文テキスト */

    /* ── 見出しのリンクアイコンを非表示 ── */
    h4 a, h3 a, h2 a { display: none !important; }

    /* ── 「使ってみる→」リンクをBold ── */
    [data-testid="stPageLink"] a,
    [data-testid="stPageLink"] p,
    [data-testid="stPageLink"] span,
    [data-testid="stPageLink"] * { font-weight: bold !important; }

    /* ── 背景：宇宙系ブルー（ダークモード対応） ── */
    .stApp {
        background:
            radial-gradient(ellipse at top left,      rgba(0, 100, 255, 0.45) 0%, transparent 50%),
            radial-gradient(ellipse at top right,     rgba(0, 200, 255, 0.25) 0%, transparent 45%),
            radial-gradient(ellipse at bottom left,   rgba(60, 0, 200, 0.35) 0%, transparent 50%),
            radial-gradient(ellipse at bottom right,  rgba(0, 140, 255, 0.40) 0%, transparent 50%) !important;
    }
</style>
""", unsafe_allow_html=True)

st.logo("assets/unloop-logo.png", link="https://unloop-tools.streamlit.app")

pg = st.navigation({
    "ホーム": [
        st.Page("home.py",                            title="HOME",               icon="🏠"),
    ],
    "Excel・データ": [
        st.Page("pages/1_ガントチャート.py",           title="ガントチャート",      icon="📊"),
        st.Page("pages/3_勤怠集計.py",                title="勤怠集計",            icon="📅"),
        st.Page("pages/6_Excel一括集計.py",           title="Excel一括集計",       icon="📋"),
        st.Page("pages/11_顧客リスト重複チェック.py", title="顧客リスト重複チェック", icon="👥"),
        st.Page("pages/12_在庫チェック.py",           title="在庫チェック",        icon="📦"),
    ],
    "PDF・ファイル": [
        st.Page("pages/2_PDF結合・ページ抽出.py",     title="PDF結合・ページ抽出", icon="📄"),
        st.Page("pages/10_パスワード付きZIP作成.py",  title="パスワード付きZIP",   icon="🔒"),
        st.Page("pages/13_フォルダ構成ZIP作成.py",    title="フォルダ構成ZIP作成", icon="📁"),
    ],
    "画像": [
        st.Page("pages/4_画像リサイズ.py",            title="画像リサイズ",        icon="🖼️"),
        st.Page("pages/9_画像透かし追加.py",          title="画像透かし追加",      icon="🔏"),
    ],
    "情報収集": [
        st.Page("pages/14_キーワードニュース収集.py", title="キーワードニュース収集", icon="📰"),
    ],
    "書類作成": [
        st.Page("pages/5_案内状一括生成.py",          title="案内状一括生成",      icon="📝"),
        st.Page("pages/8_経費集計.py",                title="経費集計",            icon="💴"),
    ],
})
pg.run()
