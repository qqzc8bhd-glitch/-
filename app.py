import json
from pathlib import Path
from datetime import date, datetime
from calendar import monthrange

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="給与計算入力フォーム",
    page_icon="💴",
    layout="centered",
    initial_sidebar_state="collapsed",
)

DATA_PATH = Path("payroll_data.json")

EMPLOYEES = ["仁吾", "三澤", "田辺", "須恵"]
PARTNERS = ["村上(固定)", "南本(変動)"]
FREELANCE_PARTS = ["平賀"]


def load_data() -> dict:
    if DATA_PATH.exists():
        try:
            return json.loads(DATA_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_data(data: dict) -> None:
    DATA_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def calc_period(year: int, pay_month: int) -> tuple[date, date]:
    # 例：2026年6月支給分 → 2026/5/16〜2026/6/15
    if pay_month == 1:
        start_year, start_month = year - 1, 12
    else:
        start_year, start_month = year, pay_month - 1
    return date(start_year, start_month, 16), date(year, pay_month, 15)


def month_key(year: int, pay_month: int) -> str:
    return f"{year}-{pay_month:02d}"


def fmt_date(d: date) -> str:
    return f"{d.year}年{d.month}月{d.day}日"


def default_record(year: int, pay_month: int, pay_day: int) -> dict:
    start, end = calc_period(year, pay_month)
    return {
        "year": year,
        "pay_month": pay_month,
        "pay_day": pay_day,
        "pay_date": f"{year}-{pay_month:02d}-{pay_day:02d}",
        "period_start": start.isoformat(),
        "period_end": end.isoformat(),
        "confirmed": False,
        "employees": {name: {"baseup": 0} for name in EMPLOYEES},
        "partners": {name: {"amount": 0, "travel": 0} for name in PARTNERS},
        "freelance_parts": {name: {"hours": 0.0, "baseup": 0, "travel": 0} for name in FREELANCE_PARTS},
        "notes": "",
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }


def yen_input(label: str, value: int, disabled: bool, key: str) -> int:
    return int(st.number_input(label, min_value=0, step=100, value=int(value or 0), disabled=disabled, key=key))


def render_header(record: dict) -> None:
    start = date.fromisoformat(record["period_start"])
    end = date.fromisoformat(record["period_end"])
    st.markdown(
        f"""
        <div class="card">
          <div class="small">給与計算入力フォーム</div>
          <h2>{record['year']}年{record['pay_month']}月分</h2>
          <p><b>支給日：</b>{fmt_date(date.fromisoformat(record['pay_date']))}</p>
          <p><b>計算期間：</b>{start.month}月{start.day}日 ～ {end.month}月{end.day}日</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def make_excel_bytes(record: dict) -> bytes:
    rows = []
    for name, v in record["employees"].items():
        rows.append({"区分": "社員・パート", "氏名": name, "時間": "", "ベースアップ": v.get("baseup", 0), "支給額": "", "交通費": ""})
    for name, v in record["partners"].items():
        rows.append({"区分": "フリーランス", "氏名": name, "時間": "", "ベースアップ": "", "支給額": v.get("amount", 0), "交通費": v.get("travel", 0)})
    for name, v in record["freelance_parts"].items():
        rows.append({"区分": "フリーランスパート", "氏名": name, "時間": v.get("hours", 0), "ベースアップ": v.get("baseup", 0), "支給額": "", "交通費": v.get("travel", 0)})

    df = pd.DataFrame(rows)
    meta = pd.DataFrame([
        ["年", record["year"]],
        ["月分", record["pay_month"]],
        ["支給日", record["pay_date"]],
        ["計算期間", f"{record['period_start']}〜{record['period_end']}"],
        ["確定", "済" if record.get("confirmed") else "未確定"],
        ["その他変更事項", record.get("notes", "")],
    ], columns=["項目", "内容"])

    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        meta.to_excel(writer, sheet_name="基本情報", index=False)
        df.to_excel(writer, sheet_name="入力データ", index=False)
    return output.getvalue()


st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 560px;}
    .card {border: 1px solid #ddd; border-radius: 14px; padding: 14px; margin-bottom: 12px; background: #fff;}
    .small {font-size: 0.85rem; color: #666;}
    h1, h2, h3 {line-height: 1.2;}
    div[data-testid="stExpander"] {border-radius: 12px;}
    .stButton button {width: 100%; border-radius: 10px; height: 3rem; font-weight: 700;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("給与計算入力")

all_data = load_data()

with st.sidebar:
    st.header("管理")
    st.caption("無料試作版：まずはダミーデータでテストしてください。")
    if st.button("全データを再読み込み"):
        st.rerun()

current_year = date.today().year
col1, col2 = st.columns(2)
with col1:
    year = st.selectbox("年", list(range(current_year - 1, current_year + 3)), index=1 if current_year + 1 in list(range(current_year - 1, current_year + 3)) else 1)
with col2:
    pay_month = st.selectbox("何月分の給料？", list(range(1, 13)), index=date.today().month - 1, format_func=lambda m: f"{m}月分")

last_day = monthrange(year, pay_month)[1]
default_day = 30 if last_day >= 30 else last_day
pay_day = st.selectbox("支給日", list(range(1, last_day + 1)), index=default_day - 1, format_func=lambda d: f"{year}年{pay_month}月{d}日")

key = month_key(year, pay_month)
start, end = calc_period(year, pay_month)
st.info(f"計算期間：{start.month}月{start.day}日 ～ {end.month}月{end.day}日")

if key not in all_data:
    st.warning("この月のデータはまだありません。")
    if st.button("この月を新しく作成"):
        all_data[key] = default_record(year, pay_month, pay_day)
        save_data(all_data)
        st.success("作成しました。")
        st.rerun()
    st.stop()

record = all_data[key]
# 支給日が変わった場合は未確定なら反映
if not record.get("confirmed") and record.get("pay_day") != pay_day:
    record["pay_day"] = pay_day
    record["pay_date"] = f"{year}-{pay_month:02d}-{pay_day:02d}"

render_header(record)

confirmed = bool(record.get("confirmed"))
if confirmed:
    st.success("この月は確定済みです。編集するには確定解除が必要です。")
else:
    st.warning("未確定です。入力後に確定してください。")

st.subheader("社員・パート")
for name in EMPLOYEES:
    record["employees"].setdefault(name, {"baseup": 0})
    record["employees"][name]["baseup"] = yen_input(f"{name}　ベースアップ", record["employees"][name].get("baseup", 0), confirmed, f"emp_{name}")

st.subheader("フリーランス")
for name in PARTNERS:
    record["partners"].setdefault(name, {"amount": 0, "travel": 0})
    c1, c2 = st.columns(2)
    with c1:
        record["partners"][name]["amount"] = yen_input(f"{name} 支給額", record["partners"][name].get("amount", 0), confirmed, f"p_amt_{name}")
    with c2:
        record["partners"][name]["travel"] = yen_input(f"{name} 交通費", record["partners"][name].get("travel", 0), confirmed, f"p_tr_{name}")

st.subheader("フリーランスパート")
for name in FREELANCE_PARTS:
    record["freelance_parts"].setdefault(name, {"hours": 0.0, "baseup": 0, "travel": 0})
    c1, c2, c3 = st.columns(3)
    with c1:
        record["freelance_parts"][name]["hours"] = float(st.number_input(f"{name} 時間", min_value=0.0, step=0.5, value=float(record["freelance_parts"][name].get("hours", 0.0)), disabled=confirmed, key=f"fp_h_{name}"))
    with c2:
        record["freelance_parts"][name]["baseup"] = yen_input(f"{name} ベースUP", record["freelance_parts"][name].get("baseup", 0), confirmed, f"fp_b_{name}")
    with c3:
        record["freelance_parts"][name]["travel"] = yen_input(f"{name} 交通費", record["freelance_parts"][name].get("travel", 0), confirmed, f"fp_t_{name}")

st.subheader("その他変更連絡事項")
record["notes"] = st.text_area("住所変更・入退社・扶養変更など", value=record.get("notes", ""), height=140, disabled=confirmed)

record["updated_at"] = datetime.now().isoformat(timespec="seconds")
all_data[key] = record

st.divider()

if not confirmed:
    if st.button("一時保存"):
        save_data(all_data)
        st.success("保存しました。")
    if st.button("確定する"):
        record["confirmed"] = True
        all_data[key] = record
        save_data(all_data)
        st.success("確定しました。以後は編集不可です。")
        st.rerun()
else:
    if st.button("確定解除する"):
        record["confirmed"] = False
        all_data[key] = record
        save_data(all_data)
        st.warning("確定解除しました。編集できます。")
        st.rerun()

st.download_button(
    "Excel出力",
    data=make_excel_bytes(record),
    file_name=f"給与入力_{year}年{pay_month}月分.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("過去月を見る"):
    if not all_data:
        st.write("まだデータがありません。")
    else:
        rows = []
        for k, r in sorted(all_data.items(), reverse=True):
            rows.append({
                "年月": f"{r.get('year')}年{r.get('pay_month')}月分",
                "支給日": r.get("pay_date"),
                "計算期間": f"{r.get('period_start')}〜{r.get('period_end')}",
                "状態": "確定済" if r.get("confirmed") else "未確定",
                "更新日時": r.get("updated_at", ""),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

st.caption("※無料試作版です。本番利用前にログイン・バックアップ・権限管理を追加してください。")
