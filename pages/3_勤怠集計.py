import streamlit as st
import openpyxl
import pandas as pd
import io
import zipfile
from datetime import timedelta, time, datetime, date

from utils import show_logo
show_logo()
st.page_link("home.py", label="← ホームに戻る")
st.title("📅 勤怠集計")
st.caption("社内の勤怠Excelを複数アップするだけで、社員別の月次集計表を自動作成します（1人1ファイル対応）")

# ── サンプルExcelダウンロード ──────────────────
def _make_kintai_sample():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "勤怠2026"
    ws["A1"] = "勤怠表サンプル（田中 花子）"
    ws["A10"] = "日付"
    ws["C10"] = "開始時刻"
    ws["I10"] = "普通残業"
    ws["J10"] = "深夜残業"
    ws["AM10"] = "勤務時間数"
    from datetime import time as t
    for i, row in enumerate(range(11, 41)):
        ws.cell(row=row, column=1).value = f"2026/07/{i+1:02d}"
        if i % 7 in (5, 6):
            continue
        ws.cell(row=row, column=3).value  = t(9, 0)
        ws.cell(row=row, column=39).value = 8.0
        if i % 3 == 0:
            ws.cell(row=row, column=9).value = 1.0
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

st.download_button(
    label="サンプルExcelをダウンロード",
    data=_make_kintai_sample(),
    file_name="勤怠サンプル_田中花子.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.caption("※ このサンプルを使って動作確認できます")
st.divider()

# ── ファイルアップロード ────────────────────────
uploaded_files = st.file_uploader(
    "勤怠Excelをアップロード（複数選択OK・1人1ファイル）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# ── ファイル解析関数 ────────────────────────────
def td_to_hours(val):
    """timedelta・datetime.time・数値を時間数（float）に変換。負の値は0を返す"""
    if val is None:
        return 0.0
    if isinstance(val, timedelta):
        h = val.total_seconds() / 3600
        return round(h, 2) if h > 0 else 0.0
    try:
        h = float(val)
        return round(h, 2) if h > 0 else 0.0
    except Exception:
        return 0.0

KINTAI_SHEET_KEYWORDS = ["オリジナル", "勤怠", "出勤"]

def _get_sheet(wb):
    """「オリジナル2026」等の勤怠シートを名前で探し、なければアクティブシートを返す"""
    for name in wb.sheetnames:
        if any(kw in name for kw in KINTAI_SHEET_KEYWORDS):
            return wb[name]
    return wb.active

def parse_kintai(file):
    file.seek(0)
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = _get_sheet(wb)
    sheet_name = ws.title

    # 氏名：ファイル名（拡張子なし）をそのまま使う
    name = file.name.rsplit(".", 1)[0]

    # データ行（11〜41行）を集計
    attendance = 0
    work_hours = 0.0
    normal_ot  = 0.0
    night_ot   = 0.0

    for row in range(11, 42):
        start = ws.cell(row=row, column=3).value   # 開始時刻
        if start is None:
            continue  # 出勤なし

        attendance += 1
        work_hours += td_to_hours(ws.cell(row=row, column=39).value)  # 勤務時間数(AM列)
        normal_ot  += td_to_hours(ws.cell(row=row, column=9).value)   # 普通残業(I列)
        night_ot   += td_to_hours(ws.cell(row=row, column=10).value)  # 深夜残業(J列)

    return {
        "氏名":          name,
        "出勤日数（日）": attendance,
        "労働時間（h）":  round(work_hours, 1),
        "普通残業（h）":  round(normal_ot, 1),
        "深夜残業（h）":  round(night_ot, 1),
        "_sheet":        sheet_name,
    }

# ── メイン処理 ──────────────────────────────────
if uploaded_files:
    results = []
    errors  = []

    for f in uploaded_files:
        try:
            data = parse_kintai(f)
            results.append(data)
        except Exception as e:
            errors.append(f"{f.name}：{e}")

    if errors:
        for err in errors:
            st.warning(err)

    if results:
        st.success(f"{len(results)} 件のファイルを集計しました")
        df = pd.DataFrame(results)
        with st.expander("🔍 デバッグ情報（問題確認用）"):
            for f2 in uploaded_files:
                f2.seek(0)
                wb2 = openpyxl.load_workbook(f2, data_only=True)
                ws2 = _get_sheet(wb2)
                st.write(f"**ファイル：{f2.name}　シート：{ws2.title}**")
                debug_rows = []
                for r in range(11, 20):
                    c_val = ws2.cell(row=r, column=3).value
                    i_val = ws2.cell(row=r, column=9).value
                    am_val = ws2.cell(row=r, column=39).value
                    debug_rows.append({"行": r, "C列(開始)": c_val, "I列(普残)": i_val, "AM列(勤務h)": am_val})
                st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)
        df = df.drop(columns=["_sheet"])
        st.dataframe(df, use_container_width=True)

        # Excel出力
        def to_excel(df):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "勤怠集計"
            from openpyxl.styles import Font, PatternFill, Alignment
            headers = list(df.columns)
            hfill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            hfont = Font(name="Meiryo UI", bold=True, color="FFFFFF")
            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.fill = hfill
                cell.font = hfont
                cell.alignment = Alignment(horizontal="center")
            dfont = Font(name="Meiryo UI")
            for ri, row in enumerate(df.itertuples(index=False), 2):
                for ci, val in enumerate(row, 1):
                    ws.cell(row=ri, column=ci, value=val).font = dfont
            for col in ws.columns:
                max_len = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 12)
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf

        st.download_button(
            label="集計結果をExcelでダウンロード",
            data=to_excel(df),
            file_name="勤怠集計結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with st.expander("📖 使い方"):
    st.write("1. 社員ごとの勤怠Excelをまとめてアップロード（複数選択OK）")
    st.write("2. 自動で氏名・出勤日数・労働時間・残業時間を読み取って集計")
    st.write("3. 「集計結果をExcelでダウンロード」ボタンで取得")
    st.write("")
    st.write("**対応フォーマット**")
    st.write("- 社内勤怠フォーマット（開始時刻：C列、終了時刻：F列、普通残業：I列、深夜残業：J列）")
    st.write("- 氏名が空の場合はファイル名を氏名として使用します")
