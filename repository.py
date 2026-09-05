from __future__ import annotations

import csv
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Protocol

from app_core import as_bool, normalize_customer


CUSTOMER_HEADERS = [
    "customer_code",
    "customer_name",
    "area",
    "is_rental",
    "is_active",
    "note",
]
DELIVERY_HEADERS = [
    "record_id",
    "delivery_date",
    "delivery_time",
    "customer_code",
    "customer_name",
    "area",
    "is_rental",
    "visit_status",
    "liters",
    "rental_slip_status",
    "note",
    "created_at",
    "updated_at",
    "is_deleted",
]
TANK_HEADERS = ["date", "meter", "stock_liters", "dispensed_liters", "note"]
FORBIDDEN_COLUMNS = {"address", "住所", "郵便番号", "電話番号"}


class Repository(Protocol):
    mode_label: str

    def load_customers(self) -> list[dict[str, Any]]: ...
    def load_monthly_history(self) -> list[dict[str, Any]]: ...
    def load_historical_refills(self) -> list[dict[str, Any]]: ...
    def load_deliveries(self) -> list[dict[str, Any]]: ...
    def load_tank_inventory(self) -> list[dict[str, Any]]: ...
    def upsert_tank_inventory(self, record: dict[str, Any]) -> None: ...
    def append_delivery(self, record: dict[str, Any]) -> None: ...
    def update_delivery(self, record_id: str, record: dict[str, Any]) -> None: ...


def _check_privacy(headers: list[str]) -> None:
    forbidden = FORBIDDEN_COLUMNS.intersection(set(headers))
    if forbidden:
        raise ValueError(f"住所等の禁止列を検出しました: {sorted(forbidden)}")


def _normalize_sheet_row(row: dict[str, Any]) -> dict[str, Any]:
    normalized = {key: value for key, value in row.items() if key}
    if "customer_code" in normalized:
        normalized["customer_code"] = str(normalized["customer_code"]).strip()
    for key in ("is_rental", "is_active", "is_deleted"):
        if key in normalized:
            normalized[key] = as_bool(normalized[key])
    return normalized


def _date_key(value: Any) -> str:
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return text


class LocalCsvRepository:
    mode_label = "ローカル確認モード"

    def __init__(self, data_dir: Path):
        self.data_dir = data_dir

    def _read(self, filename: str) -> list[dict[str, Any]]:
        path = self.data_dir / filename
        if not path.exists():
            return []
        with path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)
            _check_privacy(reader.fieldnames or [])
            return [_normalize_sheet_row(row) for row in reader]

    def load_customers(self) -> list[dict[str, Any]]:
        return [normalize_customer(row) for row in self._read("customers.csv")]

    def load_monthly_history(self) -> list[dict[str, Any]]:
        return self._read("monthly_history.csv")

    def load_historical_refills(self) -> list[dict[str, Any]]:
        return self._read("historical_refills.csv")

    def load_deliveries(self) -> list[dict[str, Any]]:
        return self._read("delivery_records.csv")

    def load_tank_inventory(self) -> list[dict[str, Any]]:
        return self._read("tank_inventory.csv")

    def upsert_tank_inventory(self, record: dict[str, Any]) -> None:
        path = self.data_dir / "tank_inventory.csv"
        rows = self.load_tank_inventory()
        target = _date_key(record.get("date"))
        replaced = False
        for index, row in enumerate(rows):
            if _date_key(row.get("date")) == target:
                rows[index] = {**row, **record}
                replaced = True
                break
        if not replaced:
            rows.append(record)
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(".tmp")
        with temp_path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=TANK_HEADERS)
            writer.writeheader()
            writer.writerows(
                {key: row.get(key, "") for key in TANK_HEADERS} for row in rows
            )
        temp_path.replace(path)

    def append_delivery(self, record: dict[str, Any]) -> None:
        path = self.data_dir / "delivery_records.csv"
        path.parent.mkdir(parents=True, exist_ok=True)
        exists = path.exists() and path.stat().st_size > 0
        with path.open("a", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=DELIVERY_HEADERS)
            if not exists:
                writer.writeheader()
            writer.writerow({key: record.get(key, "") for key in DELIVERY_HEADERS})

    def update_delivery(self, record_id: str, record: dict[str, Any]) -> None:
        path = self.data_dir / "delivery_records.csv"
        rows = self.load_deliveries()
        updated = False
        for index, row in enumerate(rows):
            if str(row.get("record_id")) == str(record_id):
                rows[index] = {**row, **record}
                updated = True
                break
        if not updated:
            raise KeyError(f"Record not found: {record_id}")
        temp_path = path.with_suffix(".tmp")
        with temp_path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=DELIVERY_HEADERS)
            writer.writeheader()
            writer.writerows(
                {key: row.get(key, "") for key in DELIVERY_HEADERS} for row in rows
            )
        temp_path.replace(path)


class GoogleSheetsRepository:
    mode_label = "Google Sheets接続モード"

    def __init__(self, spreadsheet_id: str, service_account: dict[str, Any]):
        import gspread

        client = gspread.service_account_from_dict(service_account)
        self.book = client.open_by_key(spreadsheet_id)
        self._worksheet_not_found = gspread.WorksheetNotFound
        self._api_error = gspread.exceptions.APIError

    def _call_with_retry(self, operation, attempts: int = 4):
        for attempt in range(attempts):
            try:
                return operation()
            except self._api_error as exc:
                status = getattr(getattr(exc, "response", None), "status_code", None)
                retryable = status in {429, 500, 502, 503, 504}
                if not retryable or attempt == attempts - 1:
                    raise
                time.sleep(2**attempt)
        return None

    def _read_values_with_retry(self, worksheet, attempts: int = 4) -> list[list[str]]:
        return self._call_with_retry(worksheet.get_all_values, attempts) or []

    def _records(self, sheet_name: str) -> list[dict[str, Any]]:
        worksheet = self.book.worksheet(sheet_name)
        values = self._read_values_with_retry(worksheet)
        if not values:
            return []
        headers = values[0]
        _check_privacy(headers)
        rows: list[dict[str, Any]] = []
        for values_row in values[1:]:
            padded = values_row + [""] * max(0, len(headers) - len(values_row))
            row = dict(zip(headers, padded[: len(headers)]))
            if any(str(value).strip() for value in row.values()):
                rows.append(_normalize_sheet_row(row))
        return rows

    def load_customers(self) -> list[dict[str, Any]]:
        return [normalize_customer(row) for row in self._records("顧客マスター")]

    def load_monthly_history(self) -> list[dict[str, Any]]:
        return self._records("月次実績")

    def load_historical_refills(self) -> list[dict[str, Any]]:
        try:
            return self._records("過去補給日")
        except self._worksheet_not_found:
            return []

    def load_deliveries(self) -> list[dict[str, Any]]:
        return [
            row
            for row in self._records("配送記録")
            if str(row.get("record_id", "")).strip()
        ]

    def load_tank_inventory(self) -> list[dict[str, Any]]:
        return self._records("タンク在庫")

    def upsert_tank_inventory(self, record: dict[str, Any]) -> None:
        worksheet = self.book.worksheet("タンク在庫")
        values = self._read_values_with_retry(worksheet)
        target = _date_key(record.get("date"))
        target_row = None
        for row_number, values_row in enumerate(values[1:], start=2):
            if values_row and _date_key(values_row[0]) == target:
                target_row = row_number
                break
        row_values = [record.get(key, "") for key in TANK_HEADERS]
        if target_row is None:
            self._call_with_retry(
                lambda: worksheet.append_row(row_values, value_input_option="USER_ENTERED")
            )
        else:
            self._call_with_retry(
                lambda: worksheet.update(
                    values=[row_values],
                    range_name=f"A{target_row}:E{target_row}",
                    value_input_option="USER_ENTERED",
                )
            )

    def append_delivery(self, record: dict[str, Any]) -> None:
        worksheet = self.book.worksheet("配送記録")
        worksheet.append_row(
            [record.get(key, "") for key in DELIVERY_HEADERS],
            value_input_option="USER_ENTERED",
        )

    def update_delivery(self, record_id: str, record: dict[str, Any]) -> None:
        worksheet = self.book.worksheet("配送記録")
        cell = worksheet.find(str(record_id), in_column=1)
        if not cell:
            raise KeyError(f"Record not found: {record_id}")
        existing = dict(zip(DELIVERY_HEADERS, worksheet.row_values(cell.row)))
        merged = {**existing, **record}
        worksheet.update(
            values=[[merged.get(key, "") for key in DELIVERY_HEADERS]],
            range_name=f"A{cell.row}:N{cell.row}",
            value_input_option="USER_ENTERED",
        )
