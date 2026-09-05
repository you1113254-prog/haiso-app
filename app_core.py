from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any, Iterable


VISIT_STATUSES = ("補給あり", "訪問・補給なし", "今回は飛ばした")
RENTAL_SLIP_STATUSES = ("計上済み", "今月すでに計上済み", "未計上")
NON_RENTAL_SLIP_STATUS = "対象外"


def as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"true", "1", "yes", "y", "r", "有効"}


def as_float(value: Any) -> float | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(str(value).replace(",", ""))
    except (ValueError, InvalidOperation):
        return None


def rounded_meter(value: Any) -> int | None:
    """Round a non-empty meter value to an integer using ordinary half-up rounding."""
    number = as_float(value)
    if number is None:
        return None
    return int(Decimal(str(number)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def customer_label(customer: dict[str, Any]) -> str:
    rental = " ｜ R" if as_bool(customer.get("is_rental")) else ""
    return (
        f"{customer.get('customer_name', '').strip()}"
        f" ｜ コード {str(customer.get('customer_code', '')).strip()}"
        f" ｜ {customer.get('area', '').strip()}{rental}"
    )


def normalize_customer(customer: dict[str, Any]) -> dict[str, Any]:
    return {
        "customer_code": str(customer.get("customer_code", "")).strip(),
        "customer_name": str(customer.get("customer_name", "")).strip(),
        "area": str(customer.get("area", "")).strip(),
        "is_rental": as_bool(customer.get("is_rental")),
        "is_active": as_bool(customer.get("is_active", True)),
        "note": str(customer.get("note", "") or "").strip(),
    }


def month_key(value: date | datetime | str) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m")
    if isinstance(value, date):
        return value.strftime("%Y-%m")
    return str(value)[:7]


def validate_delivery(record: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if not str(record.get("customer_code", "")).strip():
        errors.append("お客様を選択してください。")
    if record.get("visit_status") not in VISIT_STATUSES:
        errors.append("訪問結果を選択してください。")

    liters = as_float(record.get("liters")) or 0.0
    if record.get("visit_status") == "補給あり" and liters <= 0:
        errors.append("「補給あり」の場合は灯油量を入力してください。")
    if record.get("visit_status") != "補給あり" and liters != 0:
        errors.append("補給しなかった場合、灯油量は0Lにしてください。")

    is_rental = as_bool(record.get("is_rental"))
    slip = record.get("rental_slip_status")
    if is_rental and record.get("visit_status") == "今回は飛ばした":
        errors.append("R顧客は訪問・伝票計上が必須のため「今回は飛ばした」を選択できません。")
    if is_rental and slip not in RENTAL_SLIP_STATUSES:
        errors.append("レンタル顧客の伝票状況を選択してください。")
    if not is_rental and slip != NON_RENTAL_SLIP_STATUS:
        errors.append("レンタル対象外の顧客は伝票状況を「対象外」にしてください。")
    return errors


def active_records(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    return [record for record in records if not as_bool(record.get("is_deleted", False))]


def slip_already_counted(
    records: Iterable[dict[str, Any]], customer_code: str, delivery_date: date | str
) -> bool:
    target_month = month_key(delivery_date)
    for record in active_records(records):
        if str(record.get("customer_code", "")) != str(customer_code):
            continue
        if month_key(str(record.get("delivery_date", ""))) != target_month:
            continue
        if (
            record.get("visit_status") in {"補給あり", "訪問・補給なし"}
            and record.get("rental_slip_status") == "計上済み"
        ):
            return True
    return False


def summarize(records: Iterable[dict[str, Any]]) -> dict[str, Any]:
    rows = active_records(records)
    total_liters = sum(as_float(row.get("liters")) or 0.0 for row in rows)
    supplied = sum(row.get("visit_status") == "補給あり" for row in rows)
    slips = sum(row.get("rental_slip_status") == "計上済み" for row in rows)
    warnings = sum(
        as_bool(row.get("is_rental")) and row.get("rental_slip_status") == "未計上"
        for row in rows
    )
    visits = sum(
        row.get("visit_status") in {"補給あり", "訪問・補給なし"} for row in rows
    )
    return {
        "visits": visits,
        "supplied": supplied,
        "liters": round(total_liters, 1),
        "slips": slips,
        "rental_warnings": warnings,
    }


def area_summary(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, float | int | str]] = defaultdict(
        lambda: {"area": "", "visits": 0, "liters": 0.0}
    )
    for row in active_records(records):
        area = str(row.get("area", "未設定") or "未設定")
        grouped[area]["area"] = area
        grouped[area]["visits"] = int(grouped[area]["visits"]) + 1
        grouped[area]["liters"] = float(grouped[area]["liters"]) + (
            as_float(row.get("liters")) or 0.0
        )
    return sorted(
        [
            {
                "エリア": item["area"],
                "訪問件数": item["visits"],
                "灯油量(L)": round(float(item["liters"]), 1),
            }
            for item in grouped.values()
        ],
        key=lambda item: (-float(item["灯油量(L)"]), str(item["エリア"])),
    )


def month_records(records: Iterable[dict[str, Any]], selected_month: str) -> list[dict[str, Any]]:
    return [
        row
        for row in active_records(records)
        if month_key(str(row.get("delivery_date", ""))) == selected_month
    ]


def refill_timeline(
    records: Iterable[dict[str, Any]], customer_code: str
) -> list[dict[str, Any]]:
    # One calendar day is one cycle point. If an imported historical row and
    # an app record share a date, prefer the app record because it has the
    # exact amount and time.
    refills_by_day: dict[date, dict[str, Any]] = {}
    for row in active_records(records):
        if str(row.get("customer_code", "")).strip() != str(customer_code).strip():
            continue
        is_historical = bool(str(row.get("refill_date", "")).strip())
        if not is_historical and (
            row.get("visit_status") != "補給あり"
            or (as_float(row.get("liters")) or 0) <= 0
        ):
            continue
        try:
            raw_date = row.get("refill_date") if is_historical else row.get("delivery_date")
            delivery_day = date.fromisoformat(str(raw_date or "")[:10])
        except ValueError:
            continue
        liters = as_float(row.get("liters"))
        candidate = {
            "delivery_date": delivery_day,
            "delivery_time": str(row.get("delivery_time", "")),
            "liters": round(liters, 1) if liters is not None else None,
            "is_historical": is_historical,
        }
        current = refills_by_day.get(delivery_day)
        if current is None or (current["is_historical"] and not is_historical):
            refills_by_day[delivery_day] = candidate

    refills = sorted(
        refills_by_day.values(),
        key=lambda row: (row["delivery_date"], row["delivery_time"]),
    )
    previous_day: date | None = None
    for row in refills:
        current_day = row["delivery_date"]
        row["days_since_previous"] = (
            (current_day - previous_day).days if previous_day is not None else None
        )
        previous_day = current_day
    return refills


def refill_cycle_summary(
    timeline: Iterable[dict[str, Any]], today: date
) -> dict[str, Any]:
    rows = list(timeline)
    if not rows:
        return {
            "last_refill_date": None,
            "days_since_last": None,
            "last_cycle_days": None,
            "average_cycle_days": None,
        }
    intervals = [
        int(row["days_since_previous"])
        for row in rows
        if row.get("days_since_previous") is not None
    ]
    last_day = rows[-1]["delivery_date"]
    return {
        "last_refill_date": last_day,
        "days_since_last": (today - last_day).days,
        "last_cycle_days": intervals[-1] if intervals else None,
        "average_cycle_days": round(sum(intervals) / len(intervals), 1) if intervals else None,
    }


def monthly_unvisited_customers(
    customers: Iterable[dict[str, Any]], records: Iterable[dict[str, Any]]
) -> list[dict[str, Any]]:
    rows = active_records(records)
    visited_codes = {
        str(row.get("customer_code", "")).strip()
        for row in rows
        if row.get("visit_status") in {"補給あり", "訪問・補給なし"}
    }
    skipped_dates: dict[str, list[str]] = defaultdict(list)
    for row in rows:
        if row.get("visit_status") != "今回は飛ばした":
            continue
        code = str(row.get("customer_code", "")).strip()
        skipped_dates[code].append(str(row.get("delivery_date", "")))

    result = []
    for customer in customers:
        code = str(customer.get("customer_code", "")).strip()
        if not as_bool(customer.get("is_active", True)) or code in visited_codes:
            continue
        item = dict(customer)
        item["last_skipped_date"] = max(skipped_dates.get(code, []), default="")
        result.append(item)
    return sorted(
        result,
        key=lambda row: (str(row.get("area", "")), str(row.get("customer_name", ""))),
    )


def monthly_rental_slip_pending(
    customers: Iterable[dict[str, Any]], records: Iterable[dict[str, Any]]
) -> list[dict[str, Any]]:
    rows = active_records(records)
    actual_visits: dict[str, list[dict[str, Any]]] = defaultdict(list)
    counted_codes: set[str] = set()
    for row in rows:
        code = str(row.get("customer_code", "")).strip()
        if row.get("visit_status") in {"補給あり", "訪問・補給なし"}:
            actual_visits[code].append(row)
        if (
            row.get("visit_status") in {"補給あり", "訪問・補給なし"}
            and row.get("rental_slip_status") in {"計上済み", "今月すでに計上済み"}
        ):
            counted_codes.add(code)

    result = []
    for customer in customers:
        code = str(customer.get("customer_code", "")).strip()
        if (
            not as_bool(customer.get("is_active", True))
            or not as_bool(customer.get("is_rental"))
            or code in counted_codes
        ):
            continue
        item = dict(customer)
        visits = actual_visits.get(code, [])
        if visits:
            latest = max(
                visits,
                key=lambda row: (
                    str(row.get("delivery_date", "")),
                    str(row.get("delivery_time", "")),
                ),
            )
            item["last_visit_date"] = str(latest.get("delivery_date", ""))
            item["last_visit_status"] = str(latest.get("visit_status", ""))
        else:
            item["last_visit_date"] = ""
            item["last_visit_status"] = "未訪問"
        result.append(item)
    return sorted(
        result,
        key=lambda row: (str(row.get("area", "")), str(row.get("customer_name", ""))),
    )


def format_liters(value: Any) -> str:
    number = as_float(value)
    if number is None:
        return "—"
    return f"{number:,.1f}L"
