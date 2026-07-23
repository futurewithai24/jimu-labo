import streamlit as st
import pandas as pd
import io
import zipfile

from utils import show_logo
show_logo()
st.page_link("home.py", label="← ホームに戻る")
st.title("📋 Excel一括集計")
st.caption("列名がバラバラな複数のExcelをアップするだけで、1つのファイルにまとめます")

# ── 使い方説明 ─────────────────────────────────
st.info("複数の拠点・担当者からExcelを集めたとき、列名が微妙に違っていても自動で統一してまとめます。")

# ── サンプルZIPダウンロード ────────────────────
def _make_sample_zip():
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as zf:
        for name, rows in [
            ("東京支店.xlsx",  [["受付日","顧客名","商品名","売上金額","担当者"],["2026/06/01","山田商事","A製品",50000,"田中"],["2026/06/02","佐藤物産","B製品",30000,"鈴木"]]),
            ("大阪支店.xlsx",  [["登録日","得意先名","品名","税込売上","担当営業"],["2026/06/01","中村商店","A製品",48000,"山本"],["2026/06/03","高橋工業","C製品",75000,"伊藤"]]),
        ]:
            wb = pd.DataFrame(rows[1:], columns=rows[0])
            buf = io.BytesIO()
            wb.to_excel(buf, index=False)
            buf.seek(0)
            zf.writestr(name, buf.read())
    zip_buf.seek(0)
    return zip_buf

st.download_button(
    label="サンプルExcelをダウンロード（2ファイル入りZIP）",
    data=_make_sample_zip(),
    file_name="サンプル_Excel一括集計.zip",
    mime="application/zip"
)
st.caption("※ 東京支店・大阪支店で列名が異なるサンプルです。列名統一の練習に使えます。")

st.divider()

# ── ファイルアップロード ────────────────────────
uploaded_files = st.file_uploader(
    "Excelファイルをアップロード（複数選択OK）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    # 全ファイルの列名を収集・表示
    all_columns = set()
    file_info   = []

    for f in uploaded_files:
        try:
            df_tmp = pd.read_excel(f)
            df_tmp.columns = df_tmp.columns.str.strip()
            cols = list(df_tmp.columns)
            all_columns.update(cols)
            file_info.append({"ファイル名": f.name, "行数": len(df_tmp), "列名": "、".join(cols)})
        except Exception as e:
            file_info.append({"ファイル名": f.name, "行数": "エラー", "列名": str(e)})

    st.write(f"**{len(uploaded_files)} 件のExcelを読み込みました**")
    st.dataframe(pd.DataFrame(file_info), use_container_width=True)

    st.divider()

    # ── 列名の統一設定 ──────────────────────────
    st.subheader("列名の統一設定（任意）")
    st.write("列名が複数ファイルで異なる場合、まとめたい列名を「統一後の列名」に設定してください。")
    st.caption("例：あるファイルは「売上金額」、別ファイルは「税込売上」→ どちらも「売上金額」に統一")

    sorted_cols = sorted(all_columns)
    mapping_input = st.text_area(
        "列名の対応表（形式：元の列名 → 統一後の列名　※1行に1つ）",
        placeholder="例：\n税込売上 → 売上金額\n得意先名 → 顧客名\n担当営業 → 担当者",
        height=150
    )

    # ── 出力列の選択 ────────────────────────────
    st.subheader("出力する列を選択")
    selected_cols = st.multiselect(
        "集計結果に含める列を選んでください（未選択の場合は全列を出力）",
        options=sorted_cols
    )

    if st.button("Excelをまとめる", type="primary"):
        with st.spinner("作成中です。しばらくお待ちください…"):

            # 列名の対応表を解析
            col_map = {}
            if mapping_input.strip():
                for line in mapping_input.strip().splitlines():
                    if "→" in line:
                        parts = line.split("→", 1)
                        src = parts[0].strip()
                        dst = parts[1].strip()
                        if src and dst:
                            col_map[src] = dst

            all_data = []
            for f in uploaded_files:
                try:
                    f.seek(0)
                    df = pd.read_excel(f)
                    df.columns = df.columns.str.strip()

                    # 列名を統一
                    if col_map:
                        df = df.rename(columns=col_map)

                    # 拠点名をファイル名から追加
                    df["ファイル名"] = f.name.rsplit(".", 1)[0]

                    all_data.append(df)
                except Exception as e:
                    st.warning(f"{f.name} の処理中にエラーが発生しました：{e}")

            if all_data:
                result = pd.concat(all_data, ignore_index=True)

                # 出力列の絞り込み
                if selected_cols:
                    exist = [c for c in selected_cols if c in result.columns]
                    if "ファイル名" not in exist:
                        exist = ["ファイル名"] + exist
                    result = result[exist]

                # Excel出力
                output = io.BytesIO()
                result.to_excel(output, index=False)
                output.seek(0)

                st.success(f"完成しました！合計 {len(result)} 行のデータをまとめました")
                st.dataframe(result.head(20), use_container_width=True)
                if len(result) > 20:
                    st.caption(f"※ 先頭20行を表示しています（全 {len(result)} 行）")

                st.download_button(
                    label="Excelをダウンロード",
                    data=output,
                    file_name="結果一覧.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("処理できるデータがありませんでした。")

with st.expander("📖 使い方"):
    st.write("1. まとめたいExcelファイルを複数アップロード")
    st.write("2. 各ファイルの列名が一覧表示されます")
    st.write("3. 列名が異なる場合は「列名の対応表」に入力（任意）")
    st.write("   - 形式：元の列名 → 統一後の列名")
    st.write("4. 出力する列をマルチセレクトで選択（任意・未選択で全列出力）")
    st.write("5. 「Excelをまとめる」ボタンを押してダウンロード")
