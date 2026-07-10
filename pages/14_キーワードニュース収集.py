import streamlit as st
import feedparser
import pandas as pd
import io
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
import time

st.page_link("home.py", label="← ホームに戻る")
st.title("📰 キーワードニュース収集")
st.caption("キーワードを入力するだけで、Google ニュースから最新記事を自動収集してCSVでダウンロードできます")

with st.sidebar:
    st.header("設定")
    days = st.slider("直近何日分を取得するか", 1, 30, 7)

st.write("**収集したいキーワードを1行に1つ入力してください**")
st.caption("複数キーワードのAND検索は `+` でつなぐ（例：`副業+Python`）")

keyword_input = st.text_area(
    "キーワード",
    value="副業+Python\nAI+自動化",
    height=150,
    placeholder="副業+Python\nExcel+自動化\nStreamlit",
)

if st.button("ニュースを収集する", type="primary"):
    keywords = [k.strip() for k in keyword_input.strip().splitlines() if k.strip()]

    if not keywords:
        st.warning("キーワードを入力してください")
    else:
        cutoff = datetime.now(timezone.utc) - timedelta(days=days)
        results = []

        progress = st.progress(0, text="収集中です…")
        for i, keyword in enumerate(keywords):
            url = f"https://news.google.com/rss/search?q={keyword}&hl=ja&gl=JP&ceid=JP:ja"
            feed = feedparser.parse(url)
            for entry in feed.entries:
                published_str = entry.get("published", "")
                try:
                    published_dt = parsedate_to_datetime(published_str)
                    if published_dt < cutoff:
                        continue
                except Exception:
                    pass
                results.append({
                    "キーワード": keyword,
                    "タイトル": entry.get("title", ""),
                    "URL": entry.get("link", ""),
                    "更新日時": published_str,
                })
            time.sleep(1)
            progress.progress((i + 1) / len(keywords), text=f"収集中… {i+1}/{len(keywords)}")

        progress.empty()

        if not results:
            st.warning(f"直近 {days} 日以内の記事が見つかりませんでした")
        else:
            df = pd.DataFrame(results)
            st.success(f"完成しました！{len(df)} 件の記事を収集しました")
            st.dataframe(df, use_container_width=True)

            buf = io.BytesIO()
            df.to_csv(buf, index=False, encoding="utf-8-sig")
            buf.seek(0)
            today = datetime.now().strftime("%Y%m%d")
            st.download_button(
                label="CSVをダウンロード",
                data=buf,
                file_name=f"news_{today}.csv",
                mime="text/csv",
            )

with st.expander("📖 使い方"):
    st.write("1. 収集したいキーワードを1行1つで入力する")
    st.write("2. サイドバーで取得期間（日数）を設定する")
    st.write("3. 「ニュースを収集する」ボタンを押す")
    st.write("4. 結果をCSVでダウンロードする")
    st.write("- AND検索は `副業+Python` のように + でつなぐ")
    st.write("- Google ニュースのRSSを使っているため、登録不要・無料で使えます")
