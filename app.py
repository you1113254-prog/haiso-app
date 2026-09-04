from __future__ import annotations

import io
import hmac
import re
import uuid
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from app_core import (
    NON_RENTAL_SLIP_STATUS,
    RENTAL_SLIP_STATUSES,
    VISIT_STATUSES,
    active_records,
    area_summary,
    as_bool,
    as_float,
    customer_label,
    format_liters,
    month_key,
    month_records,
    refill_cycle_summary,
    refill_timeline,
    slip_already_counted,
    summarize,
    validate_delivery,
)
from repository import GoogleSheetsRepository, LocalCsvRepository


ROOT = Path(__file__).resolve().parent
JST = ZoneInfo("Asia/Tokyo")

st.set_page_config(
    page_title="灯油配送台帳",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1100px;}
      div[data-testid="stMetric"] {border: 1px solid #e5e7eb; border-radius: 12px; padding: 12px;}
      div[data-testid="stForm"] {border: 1px solid #dbeafe; border-radius: 14px; padding: 18px;}
      .customer-card {background:#f8fafc;border:1px solid #dbe3ec;border-radius:14px;padding:14px 16px;margin:8px 0 16px;}
      .customer-name {font-size:1.35rem;font-weight:700;margin-bottom:6px;}
      .privacy-note {font-size:.9rem;color:#475569;background:#f1f5f9;border-radius:10px;padding:10px 12px;}
      @media (max-width: 720px) {.block-container {padding-left: .8rem; padding-right: .8rem;} }
    </style>
    """,
    unsafe_allow_html=True,
)


def require_login() -> None:
    """Keep the existing app login without publishing its password in source."""
    if st.session_state.get("authenticated"):
        return

    try:
        configured_password = str(st.secrets.get("APP_PASSWORD", ""))
    except FileNotFoundError:
        configured_password = ""

    if not configured_password:
        st.error("管理者によるログインパスワード設定が必要です。")
        st.stop()

    st.title("🛢️ 灯油配送台帳")
    st.subheader("ログイン")
    with st.form("login_form"):
        entered = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン", type="primary", width="stretch")

    if submitted:
        valid = hmac.compare_digest(entered, configured_password)
        if valid:
            st.session_state["authenticated"] = True
            st.rerun()
        st.error("パスワードが違います。")
    st.stop()


require_login()


@st.cache_resource
def build_repository():
    try:
        spreadsheet_id = str(st.secrets["spreadsheet_id"])
        service_account = dict(st.secrets["gcp_service_account"])
        if service_account.get("private_key"):
            service_account["private_key"] = str(service_account["private_key"]).replace(
                "\\n", "\n"
            )
        if spreadsheet_id and service_account:
            return GoogleSheetsRepository(spreadsheet_id, service_account)
    except (KeyError, FileNotFoundError):
        pass
    local_customers = ROOT / "data" / "customers.csv"
    if local_customers.exists():
        return LocalCsvRepository(ROOT / "data")
    raise RuntimeError("Google Sheetsの接続設定が見つかりません。")


try:
    repo = build_repository()
except Exception:
    st.error("Google Sheetsへ接続できません。管理者へ接続設定の確認を依頼してください。")
    st.stop()


def refresh_data() -> None:
    st.session_state["customers"] = repo.load_customers()
    st.session_state["monthly_history"] = repo.load_monthly_history()
    st.session_state["deliveries"] = repo.load_deliveries()
    st.session_state["tank_inventory"] = repo.load_tank_inventory()


if "customers" not in st.session_state:
    refresh_data()

customers = [row for row in st.session_state["customers"] if row.get("is_active", True)]
monthly_history = st.session_state["monthly_history"]
deliveries = st.session_state["deliveries"]
tank_inventory = st.session_state["tank_inventory"]


def clear_pending() -> None:
    st.session_state.pop("pending_delivery", None)


def customer_picker(key: str, title: str = "お客様名"):
    ordered = sorted(customers, key=lambda row: (str(row["customer_name"]), str(row["customer_code"])))
    label_to_customer = {customer_label(row): row for row in ordered}
    selected_label = st.selectbox(
        title,
        options=list(label_to_customer),
        index=None,
        placeholder="名前を1文字入力すると候補が絞られます",
        key=key,
        on_change=clear_pending,
    )
    return label_to_customer.get(selected_label)


def show_customer_card(customer: dict) -> None:
    rental_badge = "R・レンタル" if customer["is_rental"] else "買取・一般"
    st.markdown(
        f"""
        <div class="customer-card">
          <div class="customer-name">{customer['customer_name']}</div>
          <div>コード：<b>{customer['customer_code']}</b>　｜　エリア：<b>{customer['area']}</b>　｜　{rental_badge}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def monthly_row_for(code: str):
    return next(
        (row for row in monthly_history if str(row.get("customer_code")) == str(code)),
        None,
    )


def history_for(code: str):
    rows = [
        row
        for row in active_records(deliveries)
        if str(row.get("customer_code")) == str(code)
    ]
    return sorted(
        rows,
        key=lambda row: (str(row.get("delivery_date", "")), str(row.get("delivery_time", ""))),
        reverse=True,
    )


def monthly_unvisited_view(customers: list[dict], records: list[dict]) -> list[dict]:
    rows = active_records(records)
    visited_codes = {
        str(row.get("customer_code", "")).strip()
        for row in rows
        if row.get("visit_status") in {"補給あり", "訪問・補給なし"}
    }
    skipped_dates: dict[str, list[str]] = {}
    for row in rows:
        if row.get("visit_status") == "今回は飛ばした":
            code = str(row.get("customer_code", "")).strip()
            skipped_dates.setdefault(code, []).append(str(row.get("delivery_date", "")))

    result = []
    for customer in customers:
        code = str(customer.get("customer_code", "")).strip()
        if code in visited_codes:
            continue
        item = dict(customer)
        item["last_skipped_date"] = max(skipped_dates.get(code, []), default="")
        result.append(item)
    return sorted(result, key=lambda row: (str(row.get("area", "")), str(row.get("customer_name", ""))))


def monthly_rental_pending_view(customers: list[dict], records: list[dict]) -> list[dict]:
    rows = active_records(records)
    actual_visits: dict[str, list[dict]] = {}
    counted_codes = set()
    for row in rows:
        code = str(row.get("customer_code", "")).strip()
        if row.get("visit_status") in {"補給あり", "訪問・補給なし"}:
            actual_visits.setdefault(code, []).append(row)
        if (
            row.get("visit_status") in {"補給あり", "訪問・補給なし"}
            and row.get("rental_slip_status") in {"計上済み", "今月すでに計上済み"}
        ):
            counted_codes.add(code)

    result = []
    for customer in customers:
        code = str(customer.get("customer_code", "")).strip()
        if not customer.get("is_rental") or code in counted_codes:
            continue
        item = dict(customer)
        visits = actual_visits.get(code, [])
        if visits:
            latest = max(
                visits,
                key=lambda row: (str(row.get("delivery_date", "")), str(row.get("delivery_time", ""))),
            )
            item["last_visit_date"] = str(latest.get("delivery_date", ""))
            item["last_visit_status"] = str(latest.get("visit_status", ""))
        else:
            item["last_visit_date"] = ""
            item["last_visit_status"] = "未訪問"
        result.append(item)
    return sorted(result, key=lambda row: (str(row.get("area", "")), str(row.get("customer_name", ""))))


def show_monthly_metrics(customer: dict) -> None:
    row = monthly_row_for(customer["customer_code"])
    if not row:
        st.info("この顧客の月次実績は登録されていません。")
        return
    month_keys = sorted(
        (str(key) for key in row if re.fullmatch(r"\d{4}-\d{2}", str(key))),
    )[-4:]
    if not month_keys:
        st.info("この顧客の月次実績は登録されていません。")
        return
    cols = st.columns(len(month_keys) + 1)
    for index, key in enumerate(month_keys):
        cols[index].metric(key.replace("-", "年", 1) + "月", format_liters(row.get(key)))
    total = sum(as_float(row.get(key)) or 0 for key in month_keys)
    cols[-1].metric(f"{len(month_keys)}か月計", f"{total:,.1f}L")

    dated_refills = refill_timeline(deliveries, customer["customer_code"])
    dates_by_month: dict[str, list[str]] = {}
    for refill in dated_refills:
        key = refill["delivery_date"].strftime("%Y-%m")
        dates_by_month.setdefault(key, []).append(refill["delivery_date"].strftime("%m/%d"))
    monthly_details = [
        {
            "月": key.replace("-", "年", 1) + "月",
            "補給日": "・".join(dates_by_month.get(key, [])) or "日付データなし",
            "月合計": format_liters(row.get(key)),
        }
        for key in month_keys
    ]
    st.dataframe(pd.DataFrame(monthly_details), hide_index=True, width="stretch")


def show_refill_cycle(customer: dict) -> None:
    timeline = refill_timeline(deliveries, customer["customer_code"])
    summary = refill_cycle_summary(timeline, now.date())
    st.markdown("#### 補給サイクル")
    metrics = st.columns(4)
    last_day = summary["last_refill_date"]
    metrics[0].metric("前回補給日", last_day.strftime("%Y/%m/%d") if last_day else "記録なし")
    elapsed = summary["days_since_last"]
    metrics[1].metric("今日まで", f"{elapsed}日" if elapsed is not None else "—")
    last_cycle = summary["last_cycle_days"]
    metrics[2].metric("直近サイクル", f"{last_cycle}日" if last_cycle is not None else "—")
    average = summary["average_cycle_days"]
    metrics[3].metric("平均サイクル", f"{average:g}日" if average is not None else "—")

    if timeline:
        history_rows = [
            {
                "補給日": row["delivery_date"].strftime("%Y/%m/%d"),
                "補給量": f"{row['liters']:.1f}L",
                "前回から": (
                    f"{row['days_since_previous']}日"
                    if row["days_since_previous"] is not None
                    else "—"
                ),
            }
            for row in reversed(timeline[-12:])
        ]
        with st.expander("日付別の補給履歴を見る"):
            st.dataframe(pd.DataFrame(history_rows), hide_index=True, width="stretch")
    else:
        st.caption("日付入りの補給履歴はまだありません。今後の保存記録から自動計算します。")


def dataframe_for_records(rows: list[dict]) -> pd.DataFrame:
    display = []
    for row in rows:
        display.append(
            {
                "日付": row.get("delivery_date", ""),
                "時刻": row.get("delivery_time", ""),
                "お客様": row.get("customer_name", ""),
                "コード": row.get("customer_code", ""),
                "エリア": row.get("area", ""),
                "結果": row.get("visit_status", ""),
                "灯油(L)": as_float(row.get("liters")) or 0.0,
                "伝票": row.get("rental_slip_status", ""),
                "備考": row.get("note", ""),
            }
        )
    return pd.DataFrame(display)


now = datetime.now(JST)
st.title("🛢️ 灯油配送台帳")

with st.sidebar:
    page = st.radio(
        "メニュー",
        ["配送入力", "顧客検索", "本日の記録", "月間チェック", "接続・使い方"],
    )
    st.caption(repo.mode_label)
    if st.button("ログアウト", width="stretch"):
        st.session_state["authenticated"] = False
        st.rerun()
    if st.button("データを再読込", width="stretch"):
        refresh_data()
        st.rerun()


if page == "配送入力":
    st.subheader("配送を記録")
    st.caption("入力欄をタップし、お客様名を一文字入力してください。名前・コード・エリア・R区分が候補に表示されます。")
    customer = customer_picker("delivery_customer")

    if customer:
        show_customer_card(customer)
        show_monthly_metrics(customer)
        show_refill_cycle(customer)

        delivery_date = st.date_input("配送日", value=now.date(), key="delivery_date")
        delivery_time = st.time_input(
            "時刻", value=now.time().replace(second=0, microsecond=0), key="delivery_time"
        )
        already_counted = customer["is_rental"] and slip_already_counted(
            deliveries, customer["customer_code"], delivery_date
        )

        with st.form("delivery_form", clear_on_submit=False):
            available_visit_statuses = (
                tuple(status for status in VISIT_STATUSES if status != "今回は飛ばした")
                if customer["is_rental"]
                else VISIT_STATUSES
            )
            visit_status = st.radio("訪問結果", available_visit_statuses, horizontal=True)
            if customer["is_rental"]:
                st.caption("R顧客は訪問・伝票計上が必須です。「今回は飛ばした」は選択できません。")
            liters = st.number_input(
                "灯油量（L）", min_value=0.0, max_value=2000.0, value=0.0, step=0.2, format="%.1f"
            )
            if customer["is_rental"]:
                default_index = 1 if already_counted else 0
                slip_status = st.radio(
                    "レンタル伝票",
                    RENTAL_SLIP_STATUSES,
                    index=default_index,
                    horizontal=True,
                )
                if already_counted:
                    st.success("今月のレンタル伝票は計上済みです。今回は灯油のみで大丈夫です。")
            else:
                slip_status = NON_RENTAL_SLIP_STATUS
                st.caption("レンタル伝票：対象外")
            note = st.text_area("備考（任意）", placeholder="ボイラーの状態、次回の注意点など")
            review = st.form_submit_button("入力内容を確認", type="primary", width="stretch")

        if review:
            pending = {
                "record_id": str(uuid.uuid4()),
                "delivery_date": delivery_date.isoformat(),
                "delivery_time": delivery_time.strftime("%H:%M"),
                "customer_code": customer["customer_code"],
                "customer_name": customer["customer_name"],
                "area": customer["area"],
                "is_rental": customer["is_rental"],
                "visit_status": visit_status,
                "liters": round(float(liters), 1),
                "rental_slip_status": slip_status,
                "note": note.strip(),
                "created_at": datetime.now(JST).isoformat(timespec="seconds"),
                "updated_at": datetime.now(JST).isoformat(timespec="seconds"),
                "is_deleted": False,
            }
            errors = validate_delivery(pending)
            if errors:
                for error in errors:
                    st.error(error)
            else:
                st.session_state["pending_delivery"] = pending

        pending = st.session_state.get("pending_delivery")
        if pending:
            st.divider()
            st.subheader("保存前の確認")
            if pending["is_rental"] and pending["rental_slip_status"] == "未計上":
                st.warning("レンタル伝票が未計上です。内容を確認してください。")
            st.write(
                f"**{pending['customer_name']}**（{pending['customer_code']}・{pending['area']}）  "
                f"\n{pending['delivery_date']} {pending['delivery_time']}／{pending['visit_status']}／"
                f"{pending['liters']:.1f}L／伝票：{pending['rental_slip_status']}"
            )
            confirm_col, cancel_col = st.columns(2)
            if confirm_col.button("確定して保存", type="primary", width="stretch"):
                try:
                    repo.append_delivery(pending)
                    clear_pending()
                    refresh_data()
                    st.success("配送記録を保存しました。")
                    st.rerun()
                except Exception as exc:
                    st.error(f"保存できませんでした。接続を確認してください。詳細: {exc}")
            if cancel_col.button("入力へ戻る", width="stretch"):
                clear_pending()
                st.rerun()
    else:
        st.info("まずお客様名を入力して、候補から選んでください。")


elif page == "顧客検索":
    st.subheader("顧客を検索")
    customer = customer_picker("search_customer", "名前・コード・エリアで検索")
    if customer:
        show_customer_card(customer)
        show_monthly_metrics(customer)
        show_refill_cycle(customer)
        rows = history_for(customer["customer_code"])
        st.markdown("#### アプリ登録後の配送履歴")
        if rows:
            st.dataframe(dataframe_for_records(rows), hide_index=True, width="stretch")
        else:
            st.info("アプリで登録した配送履歴はまだありません。")


elif page == "本日の記録":
    st.subheader("日ごとの記録")
    selected_date = st.date_input("表示する日", value=now.date(), key="daily_date")
    rows = [
        row
        for row in active_records(deliveries)
        if str(row.get("delivery_date", "")) == selected_date.isoformat()
    ]
    rows.sort(key=lambda row: str(row.get("delivery_time", "")))
    summary = summarize(rows)
    metrics = st.columns(4)
    metrics[0].metric("訪問", f"{summary['visits']}軒")
    metrics[1].metric("補給", f"{summary['supplied']}軒")
    metrics[2].metric("灯油", f"{summary['liters']:,.1f}L")
    metrics[3].metric("伝票計上", f"{summary['slips']}件")
    if summary["rental_warnings"]:
        st.warning(f"レンタル伝票の未計上が{summary['rental_warnings']}件あります。")
    if rows:
        st.dataframe(dataframe_for_records(rows), hide_index=True, width="stretch")
        text_summary = (
            f"{selected_date.isoformat()}　訪問{summary['visits']}軒（補給{summary['supplied']}軒）／"
            f"灯油{summary['liters']:.1f}L／レンタル伝票{summary['slips']}件"
        )
        st.code(text_summary, language=None)

        options = {
            f"{row.get('delivery_time','')}｜{row.get('customer_name','')}｜{format_liters(row.get('liters'))}": row
            for row in rows
        }
        with st.expander("記録を修正する"):
            selected_label = st.selectbox("修正する記録", list(options), index=None)
            selected = options.get(selected_label)
            if selected:
                with st.form("edit_record"):
                    edited_status = st.selectbox(
                        "訪問結果", VISIT_STATUSES, index=VISIT_STATUSES.index(selected.get("visit_status"))
                    )
                    edited_liters = st.number_input(
                        "灯油量（L）", min_value=0.0, max_value=2000.0,
                        value=as_float(selected.get("liters")) or 0.0, step=0.2, format="%.1f"
                    )
                    if as_bool(selected.get("is_rental")):
                        current_slip = selected.get("rental_slip_status")
                        slip_index = RENTAL_SLIP_STATUSES.index(current_slip) if current_slip in RENTAL_SLIP_STATUSES else 0
                        edited_slip = st.selectbox("伝票状況", RENTAL_SLIP_STATUSES, index=slip_index)
                    else:
                        edited_slip = NON_RENTAL_SLIP_STATUS
                    edited_note = st.text_area("備考", value=str(selected.get("note", "")))
                    save_edit = st.form_submit_button("修正を保存", type="primary")
                if save_edit:
                    updated = {
                        **selected,
                        "visit_status": edited_status,
                        "liters": round(float(edited_liters), 1),
                        "rental_slip_status": edited_slip,
                        "note": edited_note.strip(),
                        "updated_at": datetime.now(JST).isoformat(timespec="seconds"),
                    }
                    errors = validate_delivery(updated)
                    if errors:
                        for error in errors:
                            st.error(error)
                    else:
                        try:
                            repo.update_delivery(str(selected["record_id"]), updated)
                            refresh_data()
                            st.success("記録を修正しました。")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"修正できませんでした。詳細: {exc}")
    else:
        st.info("この日の記録はありません。")


elif page == "月間チェック":
    st.subheader("月間チェック")
    history_months = {
        str(key)
        for row in monthly_history
        for key in row
        if re.fullmatch(r"\d{4}-\d{2}", str(key))
    }
    available_months = sorted(
        history_months | {month_key(now)}, reverse=True
    )
    selected_month = st.selectbox("対象月", available_months)
    app_rows = month_records(deliveries, selected_month)

    if selected_month in history_months:
        month_data = []
        for row in monthly_history:
            value = as_float(row.get(selected_month))
            if value is not None:
                month_data.append(
                    {
                        "お客様": row.get("customer_name", ""),
                        "コード": row.get("customer_code", ""),
                        "エリア": row.get("area", ""),
                        "R": "R" if as_bool(row.get("is_rental")) else "",
                        "灯油(L)": value,
                    }
                )
        total = sum(float(row["灯油(L)"]) for row in month_data)
        st.metric("確定済み月間灯油量", f"{total:,.1f}L")
        st.dataframe(pd.DataFrame(month_data), hide_index=True, width="stretch")
        st.caption("取込済みの月別集計です。日ごとの伝票状況はこの表には含まれません。")
    else:
        summary = summarize(app_rows)
        metrics = st.columns(3)
        metrics[0].metric("訪問", f"{summary['visits']}軒")
        metrics[1].metric("灯油", f"{summary['liters']:,.1f}L")
        metrics[2].metric("伝票計上", f"{summary['slips']}件")
        area_rows = area_summary(app_rows)
        if area_rows:
            st.markdown("#### エリア別")
            st.dataframe(pd.DataFrame(area_rows), hide_index=True, width="stretch")

        unvisited = monthly_unvisited_view(customers, app_rows)
        rental_pending = monthly_rental_pending_view(customers, app_rows)

        st.divider()
        status_metrics = st.columns(2)
        status_metrics[0].metric("今月の未訪問", f"{len(unvisited)}件")
        status_metrics[1].metric("R伝票未計上", f"{len(rental_pending)}件")

        areas = sorted({str(row.get("area", "") or "未設定") for row in customers})
        selected_area = st.selectbox(
            "一覧のエリア絞り込み", ["すべて"] + areas, key="monthly_list_area"
        )

        def in_selected_area(row: dict[str, Any]) -> bool:
            return selected_area == "すべて" or str(row.get("area", "") or "未設定") == selected_area

        unvisited_view = [row for row in unvisited if in_selected_area(row)]
        rental_pending_view = [row for row in rental_pending if in_selected_area(row)]
        unvisited_tab, rental_tab = st.tabs(["未訪問リスト", "R伝票未計上リスト"])

        with unvisited_tab:
            st.caption(
                "今月「補給あり」または「訪問・補給なし」の記録がないお客様です。"
                "「今回は飛ばした」は未訪問として残ります。"
            )
            st.write(f"表示：{len(unvisited_view)}件")
            if unvisited_view:
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "お客様": row["customer_name"],
                                "コード": row["customer_code"],
                                "エリア": row["area"],
                                "R": "R" if row["is_rental"] else "",
                                "今月飛ばした日": row.get("last_skipped_date", ""),
                            }
                            for row in unvisited_view
                        ]
                    ),
                    hide_index=True,
                    width="stretch",
                )
            else:
                st.success("該当する未訪問顧客はいません。")

        with rental_tab:
            st.caption(
                "全R顧客のうち、今月の伝票が「計上済み」または"
                "「今月すでに計上済み」になっていないお客様です。"
                "月初は全R顧客が表示され、計上確認ごとに減って0件で完了です。"
            )
            st.write(f"表示：{len(rental_pending_view)}件")
            if rental_pending_view:
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "お客様": row["customer_name"],
                                "コード": row["customer_code"],
                                "エリア": row["area"],
                                "最終訪問日": row.get("last_visit_date", ""),
                                "今月の状況": row.get("last_visit_status", ""),
                            }
                            for row in rental_pending_view
                        ]
                    ),
                    hide_index=True,
                    width="stretch",
                )
            else:
                st.success("該当するレンタル伝票未計上はありません。")

    if app_rows:
        csv_buffer = io.StringIO()
        dataframe_for_records(app_rows).to_csv(csv_buffer, index=False)
        st.download_button(
            "この月のアプリ記録をCSV保存",
            csv_buffer.getvalue().encode("utf-8-sig"),
            file_name=f"delivery_{selected_month}.csv",
            mime="text/csv",
        )


else:
    st.subheader("接続・使い方")
    st.write(f"現在の動作：**{repo.mode_label}**")
    st.write(f"顧客マスター：**{len(customers)}件**／レンタル顧客：**{sum(row['is_rental'] for row in customers)}件**")
    st.markdown(
        """
        1. 「配送入力」を開きます。
        2. お客様名の入力欄をタップします。
        3. 名前を一文字入力すると、候補が絞り込まれます。
        4. 名前・コード・エリア・R区分を見て選択します。
        5. 数量と伝票状況を入力し、確認後に保存します。
        """
    )
    st.markdown(
        '<div class="privacy-note">このアプリのデータ項目に住所はありません。住所入りの既存シートにも接続しません。</div>',
        unsafe_allow_html=True,
    )
