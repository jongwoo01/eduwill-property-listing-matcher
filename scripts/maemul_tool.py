#!/usr/bin/env python3
"""Deterministic profile, CSV/XLSX search, and safe mutation tool for maemul-matching."""

from __future__ import annotations

import argparse
import csv
import ctypes
import fcntl
import hashlib
import io
import json
import math
import os
import shutil
import stat
import sys
import tempfile
from contextlib import contextmanager
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterable

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover - reported at the command boundary
    Workbook = None
    load_workbook = None
    Font = None
    PatternFill = None
    DataValidation = None


LISTING_COLUMNS = [
    "번호",
    "상태",
    "종류",
    "거래",
    "지역",
    "동네",
    "단지명",
    "동호",
    "매매가(만원)",
    "보증금(만원)",
    "월세(만원)",
    "관리비(만원)",
    "전용(㎡)",
    "방수",
    "욕실",
    "층",
    "총층",
    "향",
    "준공",
    "주차(대)",
    "반려",
    "옵션",
    "입주가능",
    "인근역",
    "역도보(분)",
    "초등학교",
    "초등도보(분)",
    "접수일",
]

DETAIL_COLUMNS = [
    "번호",
    "소유자",
    "연락처",
    "권리관계",
    "임대차현황",
    "용도지역",
    "건폐율·용적률",
    "위반건축물",
    "주차시설유형",
    "관리상세",
    "시설상태",
    "옵션상세",
    "학군상세",
    "비선호시설",
    "공시가격",
    "취득조세",
    "특약·메모",
    "접수경로",
    "계약일",
    "실제계약금액(만원)",
]

NUMERIC_FIELDS = {
    "매매가(만원)",
    "보증금(만원)",
    "월세(만원)",
    "관리비(만원)",
    "전용(㎡)",
    "방수",
    "욕실",
    "층",
    "총층",
    "준공",
    "주차(대)",
    "역도보(분)",
    "초등도보(분)",
    "실제계약금액(만원)",
}

NONNEGATIVE_NUMERIC_FIELDS = NUMERIC_FIELDS - {"층"}
FORMULA_PREFIXES = ("=", "+", "-", "@")

STATUS_VALUES = {"진행", "보류", "완료"}
YN_UNKNOWN_VALUES = {"Y", "N", "?"}
UNKNOWN_VALUES = {"", "?"}
SUPPORTED_OPERATORS = {"eq", "ne", "in", "not-in", "contains", "lte", "gte", "between"}

DETAIL_SOURCE_ALIASES = {
    "공시가격(만원)": "공시가격",
}

LEGACY_ONLY_FIELDS = {
    "학군(분)",
    "접수일",
}


class ToolError(Exception):
    pass


def emit(payload: dict[str, Any]) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def fail(message: str, **details: Any) -> None:
    raise ToolError(json.dumps({"message": message, **details}, ensure_ascii=False))


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def resolve_path(raw: str) -> Path:
    return Path(raw).expanduser().resolve()


def require_openpyxl() -> None:
    if load_workbook is None or Workbook is None:
        fail("로컬 Excel 기능에는 openpyxl이 필요합니다", hint="python3 -m pip install openpyxl")


def require_xlsx_path(path: Path) -> None:
    if path.suffix.casefold() != ".xlsx":
        fail("로컬 매물장은 .xlsx 파일이어야 합니다", path=str(path))


def default_profile_store() -> Path:
    codex_root = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex")).expanduser()
    return (codex_root / "maemul-matching" / "profiles.json").resolve()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fsync_directory(path: Path) -> None:
    descriptor = os.open(path, os.O_RDONLY)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def atomic_exchange(first: Path, second: Path) -> bool:
    """Atomically exchange two paths where the host exposes a native primitive."""
    library = ctypes.CDLL(None, use_errno=True)
    first_raw = os.fsencode(first)
    second_raw = os.fsencode(second)
    if sys.platform == "darwin" and hasattr(library, "renamex_np"):
        function = library.renamex_np
        function.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_uint]
        function.restype = ctypes.c_int
        result = function(first_raw, second_raw, 0x00000002)  # RENAME_SWAP
    elif sys.platform.startswith("linux") and hasattr(library, "renameat2"):
        function = library.renameat2
        function.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int, ctypes.c_char_p, ctypes.c_uint]
        function.restype = ctypes.c_int
        result = function(-100, first_raw, -100, second_raw, 0x00000002)  # AT_FDCWD, RENAME_EXCHANGE
    else:
        return False
    if result != 0:
        error_number = ctypes.get_errno()
        raise OSError(error_number, os.strerror(error_number), str(first), str(second))
    return True


def parse_json(raw: str, label: str) -> Any:
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        fail(f"{label} JSON을 해석할 수 없습니다", error=str(exc))


def is_unknown(value: Any) -> bool:
    return value is None or str(value).strip() in UNKNOWN_VALUES


def parse_number(value: Any, field: str) -> float:
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        fail("숫자 필드 값이 올바르지 않습니다", field=field, value=value)
    if not math.isfinite(number):
        fail("숫자 필드는 유한한 값이어야 합니다", field=field, value=value)
    if field in NONNEGATIVE_NUMERIC_FIELDS and number < 0:
        fail("숫자 필드는 음수일 수 없습니다", field=field, value=value)
    if field == "층" and not -20 <= number <= 200:
        fail("층 값이 허용 범위를 벗어났습니다", field=field, value=value)
    if field == "준공" and not 1800 <= number <= date.today().year + 10:
        fail("준공 연도가 허용 범위를 벗어났습니다", field=field, value=value)
    return number


def normalize_input_value(field: str, value: Any) -> str:
    """Normalize JSON scalars and neutralize spreadsheet formulas in text fields."""
    if value is None:
        return "?"
    if isinstance(value, (dict, list)):
        fail("필드 값은 객체나 배열일 수 없습니다", field=field, value=value)
    raw = str(value)
    if field not in NUMERIC_FIELDS and raw.lstrip().startswith(FORMULA_PREFIXES):
        return "'" + raw
    return raw


def reject_unsafe_formula(value: str, field: str, row_number: int) -> None:
    if field in NUMERIC_FIELDS or is_unknown(value):
        return
    if value.lstrip().startswith(FORMULA_PREFIXES):
        fail(
            "스프레드시트 수식으로 실행될 수 있는 텍스트입니다",
            field=field,
            row=row_number,
            value=value,
            hint="앞에 작은따옴표(')를 붙여 일반 텍스트로 저장하세요",
        )


def validate_iso_date(value: str, field: str, row_number: int) -> None:
    if value in UNKNOWN_VALUES:
        return
    try:
        date.fromisoformat(value)
    except ValueError:
        fail("날짜는 YYYY-MM-DD 또는 ? 형식이어야 합니다", field=field, row=row_number, value=value)


def validate_rows(rows: list[dict[str, str]], kind: str) -> dict[str, Any]:
    columns = LISTING_COLUMNS if kind == "listing" else DETAIL_COLUMNS
    seen: set[str] = set()
    warnings: list[dict[str, Any]] = []

    for index, row in enumerate(rows, start=2):
        if set(row) != set(columns):
            fail("행 컬럼이 스키마와 다릅니다", row=index)
        item_id = row["번호"].strip()
        if not item_id or item_id == "?":
            fail("번호는 비어 있거나 ?일 수 없습니다", row=index)
        if item_id in seen:
            fail("중복 번호가 있습니다", row=index, item_id=item_id)
        seen.add(item_id)

        for field in columns:
            reject_unsafe_formula(row[field], field, index)

        for field in NUMERIC_FIELDS.intersection(columns):
            if not is_unknown(row[field]):
                parse_number(row[field], field)

        if kind == "listing":
            if row["상태"] not in STATUS_VALUES:
                fail("상태 값이 올바르지 않습니다", row=index, value=row["상태"])
            for field in ("반려", "옵션"):
                if row[field] not in YN_UNKNOWN_VALUES:
                    fail("Y/N/? 필드 값이 올바르지 않습니다", row=index, field=field, value=row[field])
            validate_iso_date(row["접수일"].strip(), "접수일", index)
            if row["입주가능"] not in {"", "?", "즉시", "협의"}:
                validate_iso_date(row["입주가능"].strip(), "입주가능", index)
            for required in ("종류", "거래", "지역"):
                if is_unknown(row[required]):
                    warnings.append({"row": index, "field": required, "warning": "검색 핵심값 미확인"})
        else:
            validate_iso_date(row["계약일"].strip(), "계약일", index)

    return {"rows": len(rows), "warnings": warnings}


def read_csv_table(path: Path, kind: str) -> tuple[list[dict[str, str]], dict[str, Any]]:
    rows, summary, _ = read_csv_snapshot(path, kind)
    return rows, summary


def read_csv_snapshot(path: Path, kind: str) -> tuple[list[dict[str, str]], dict[str, Any], str]:
    """Parse and hash one immutable byte snapshot of a CSV file."""
    if not path.is_file():
        fail("CSV 파일을 찾을 수 없습니다", path=str(path))
    expected = LISTING_COLUMNS if kind == "listing" else DETAIL_COLUMNS
    try:
        content = path.read_bytes()
        text = content.decode("utf-8-sig")
        with io.StringIO(text, newline="") as handle:
            reader = csv.DictReader(handle)
            actual = reader.fieldnames or []
            if actual != expected:
                fail("CSV 헤더가 표준 스키마와 다릅니다", path=str(path), expected=expected, actual=actual)
            rows = [dict(row) for row in reader]
    except UnicodeDecodeError as exc:
        fail("CSV가 UTF-8로 읽히지 않습니다", path=str(path), error=str(exc))
    summary = validate_rows(rows, kind)
    return rows, summary, hashlib.sha256(content).hexdigest()


def workbook_cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat() if value.time().isoformat() == "00:00:00" else value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_workbook_snapshot(
    path: Path,
    kind: str,
    sheet_name: str,
) -> tuple[list[dict[str, str]], dict[str, Any], str]:
    require_openpyxl()
    require_xlsx_path(path)
    if not path.is_file():
        fail("Excel 매물장을 찾을 수 없습니다", path=str(path))
    snapshot_sha = sha256_file(path)
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            fail("Excel 매물장에 필요한 탭이 없습니다", path=str(path), sheet=sheet_name)
        sheet = workbook[sheet_name]
        expected = LISTING_COLUMNS if kind == "listing" else DETAIL_COLUMNS
        actual = [workbook_cell_text(cell.value) for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
        while actual and actual[-1] == "":
            actual.pop()
        if actual != expected:
            fail("Excel 탭 헤더가 표준 스키마와 다릅니다", path=str(path), sheet=sheet_name, expected=expected, actual=actual)
        rows: list[dict[str, str]] = []
        for values in sheet.iter_rows(min_row=2, max_col=len(expected), values_only=True):
            if all(value is None or str(value).strip() == "" for value in values):
                continue
            rows.append({field: workbook_cell_text(value) for field, value in zip(expected, values)})
    finally:
        workbook.close()
    summary = validate_rows(rows, kind)
    return rows, summary, snapshot_sha


def style_workbook_sheet(sheet: Any, columns: list[str]) -> None:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(1, len(columns)).coordinate}"
    fill = PatternFill("solid", fgColor="1F4E78")
    for index, column in enumerate(columns, start=1):
        cell = sheet.cell(1, index, column)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = fill
        sheet.column_dimensions[cell.column_letter].width = min(max(len(column) + 4, 12), 24)


def command_init_workbook(args: argparse.Namespace) -> dict[str, Any]:
    require_openpyxl()
    path = resolve_path(args.workbook)
    require_xlsx_path(path)
    if args.listing_sheet == args.detail_sheet:
        fail("검색용 탭과 상세 탭 이름은 달라야 합니다")
    if path.exists():
        fail("기존 Excel 파일을 덮어쓰지 않습니다", path=str(path))
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    listing = workbook.active
    listing.title = args.listing_sheet
    detail = workbook.create_sheet(args.detail_sheet)
    style_workbook_sheet(listing, LISTING_COLUMNS)
    style_workbook_sheet(detail, DETAIL_COLUMNS)
    status_validation = DataValidation(type="list", formula1='"진행,보류,완료"', allow_blank=False)
    yn_validation = DataValidation(type="list", formula1='"Y,N,?"', allow_blank=False)
    listing.add_data_validation(status_validation)
    listing.add_data_validation(yn_validation)
    status_validation.add("B2:B1048576")
    yn_validation.add("T2:U1048576")

    descriptor, raw_temporary = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=".xlsx", dir=path.parent)
    os.close(descriptor)
    temporary = Path(raw_temporary)
    try:
        workbook.save(temporary)
        workbook.close()
        os.chmod(temporary, 0o600)
        read_workbook_snapshot(temporary, "listing", args.listing_sheet)
        read_workbook_snapshot(temporary, "detail", args.detail_sheet)
        os.link(temporary, path)
        temporary.unlink()
        fsync_directory(path.parent)
    except Exception:
        workbook.close()
        if temporary.exists():
            temporary.unlink()
        raise
    return {
        "ok": True,
        "workbook": str(path),
        "listing_sheet": args.listing_sheet,
        "detail_sheet": args.detail_sheet,
        "sha256": sha256_file(path),
    }


def read_source_csv(path: Path) -> tuple[list[str], list[dict[str, str]], str]:
    """Read a UTF-8 CSV without requiring the current standard schema."""
    if not path.is_file():
        fail("원본 CSV 파일을 찾을 수 없습니다", path=str(path))
    content = path.read_bytes()
    try:
        text = content.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        fail("원본 CSV가 UTF-8로 읽히지 않습니다", path=str(path), error=str(exc))
    with io.StringIO(text, newline="") as handle:
        reader = csv.DictReader(handle)
        headers = reader.fieldnames or []
        if not headers or any(not header for header in headers):
            fail("원본 CSV 헤더가 비어 있습니다", path=str(path))
        if len(headers) != len(set(headers)):
            fail("원본 CSV에 중복 헤더가 있습니다", path=str(path), headers=headers)
        rows: list[dict[str, str]] = []
        for row_number, raw_row in enumerate(reader, start=2):
            if None in raw_row:
                fail("원본 CSV 행의 열 수가 헤더보다 많습니다", path=str(path), row=row_number)
            rows.append({header: (raw_row.get(header) or "") for header in headers})
    return headers, rows, hashlib.sha256(content).hexdigest()


def require_unique_source_ids(rows: list[dict[str, str]], path: Path) -> None:
    seen: set[str] = set()
    for row_number, row in enumerate(rows, start=2):
        item_id = row.get("번호", "").strip()
        if not item_id or item_id == "?":
            fail("원본 번호는 비어 있거나 ?일 수 없습니다", path=str(path), row=row_number)
        if item_id in seen:
            fail("원본에 중복 번호가 있습니다", path=str(path), row=row_number, item_id=item_id)
        seen.add(item_id)


def normalize_legacy_status(value: str, warnings: list[dict[str, Any]], item_id: str) -> str:
    raw = value.strip()
    aliases = {
        "": "보류",
        "?": "보류",
        "활성": "진행",
        "접수": "진행",
        "거래중": "진행",
        "계약완료": "완료",
        "거래완료": "완료",
        "종료": "완료",
    }
    normalized = aliases.get(raw, raw)
    if normalized not in STATUS_VALUES:
        fail("원본 상태 값을 표준 상태로 바꿀 수 없습니다", item_id=item_id, value=value)
    if normalized != raw:
        warnings.append({"item_id": item_id, "field": "상태", "from": value, "to": normalized})
    return normalized


def normalize_legacy_yn(field: str, value: str, warnings: list[dict[str, Any]], item_id: str) -> str:
    raw = value.strip()
    aliases = {
        "": "?",
        "가능": "Y",
        "있음": "Y",
        "예": "Y",
        "불가": "N",
        "없음": "N",
        "아니오": "N",
    }
    normalized = aliases.get(raw, raw.upper())
    if normalized not in YN_UNKNOWN_VALUES:
        warnings.append({"item_id": item_id, "field": field, "from": value, "to": "?"})
        return "?"
    if normalized != raw:
        warnings.append({"item_id": item_id, "field": field, "from": value, "to": normalized})
    return normalized


def split_legacy_region(value: str) -> tuple[str, str | None]:
    """Split a legacy combined region into broad region and neighborhood when explicit."""
    parts = value.strip().split()
    if len(parts) < 3:
        return value, None
    neighborhood = parts[-1]
    if not neighborhood.endswith(("동", "읍", "면", "리", "가")):
        return value, None
    return " ".join(parts[:-1]), neighborhood


def standard_row_from_source(
    source: dict[str, str],
    columns: list[str],
    source_aliases: dict[str, str] | None = None,
) -> dict[str, str]:
    aliases = source_aliases or {}
    row = {column: "?" for column in columns}
    for source_field, value in source.items():
        target = aliases.get(source_field, source_field)
        if target in row:
            row[target] = normalize_input_value(target, value)
    return row


def write_csv_rows_new(path: Path, columns: list[str], rows: list[dict[str, str]], kind: str) -> None:
    if path.exists():
        fail("기존 출력 파일을 덮어쓰지 않습니다", path=str(path))
    validate_rows(rows, kind)
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, raw_temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".normalize", dir=path.parent)
    temporary = Path(raw_temporary)
    try:
        os.chmod(temporary, 0o600)
        with os.fdopen(descriptor, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=columns, lineterminator="\n", extrasaction="raise")
            writer.writeheader()
            writer.writerows(rows)
            handle.flush()
            os.fsync(handle.fileno())
        os.link(temporary, path)
        temporary.unlink()
        fsync_directory(path.parent)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise


def command_normalize_pair(args: argparse.Namespace) -> dict[str, Any]:
    listing_input = resolve_path(args.listing_input)
    detail_input = resolve_path(args.detail_input)
    output_dir = resolve_path(args.output_dir)
    listing_output = output_dir / "매물.csv"
    detail_output = output_dir / "매물상세.csv"

    if listing_input == detail_input:
        fail("검색용 원본과 상세 원본은 달라야 합니다")
    if listing_input in {listing_output, detail_output} or detail_input in {listing_output, detail_output}:
        fail("원본과 출력 경로는 달라야 합니다", output_dir=str(output_dir))
    if listing_output.exists() or detail_output.exists():
        fail(
            "표준 출력 파일이 이미 있어 덮어쓰지 않습니다",
            listing=str(listing_output),
            detail=str(detail_output),
        )

    listing_headers, source_listings, listing_sha = read_source_csv(listing_input)
    detail_headers, source_details, detail_sha = read_source_csv(detail_input)
    require_unique_source_ids(source_listings, listing_input)
    require_unique_source_ids(source_details, detail_input)

    detail_by_id = {row["번호"].strip(): row for row in source_details}
    listing_ids = {row["번호"].strip() for row in source_listings}
    warnings: list[dict[str, Any]] = []
    normalized_listings: list[dict[str, str]] = []

    for source in source_listings:
        item_id = source["번호"].strip()
        row = standard_row_from_source(source, LISTING_COLUMNS)
        row["번호"] = item_id
        row["상태"] = normalize_legacy_status(source.get("상태", ""), warnings, item_id)
        for field in ("반려", "옵션"):
            row[field] = normalize_legacy_yn(field, source.get(field, ""), warnings, item_id)
        if is_unknown(row["동네"]):
            broad_region, neighborhood = split_legacy_region(row["지역"])
            if neighborhood:
                row["지역"] = broad_region
                row["동네"] = neighborhood
        if is_unknown(row["접수일"]):
            detail_received = detail_by_id.get(item_id, {}).get("접수일", "").strip()
            row["접수일"] = normalize_input_value("접수일", detail_received) if detail_received else "?"
        normalized_listings.append(row)

    normalized_details: list[dict[str, str]] = []
    all_detail_ids = list(dict.fromkeys([row["번호"].strip() for row in source_details] + [row["번호"].strip() for row in source_listings]))
    listing_source_by_id = {row["번호"].strip(): row for row in source_listings}
    for item_id in all_detail_ids:
        source = detail_by_id.get(item_id, {"번호": item_id})
        row = standard_row_from_source(source, DETAIL_COLUMNS, DETAIL_SOURCE_ALIASES)
        row["번호"] = item_id
        legacy_school_minutes = listing_source_by_id.get(item_id, {}).get("학군(분)", "").strip()
        if legacy_school_minutes and legacy_school_minutes != "?":
            preserved = f"기존 학군(분): {legacy_school_minutes}"
            current = row.get("학군상세", "?").strip()
            row["학군상세"] = preserved if is_unknown(current) else f"{current}; {preserved}"
        normalized_details.append(row)

    validate_rows(normalized_listings, "listing")
    validate_rows(normalized_details, "detail")
    output_dir.mkdir(parents=True, exist_ok=True)
    write_csv_rows_new(listing_output, LISTING_COLUMNS, normalized_listings, "listing")
    try:
        write_csv_rows_new(detail_output, DETAIL_COLUMNS, normalized_details, "detail")
    except Exception:
        # Do not delete a successfully created output because another process may
        # have replaced it. A new output directory makes retries unambiguous.
        raise

    listing_summary = validate_rows(normalized_listings, "listing")
    detail_summary = validate_rows(normalized_details, "detail")
    orphan_detail_ids = sorted(set(detail_by_id) - listing_ids)
    unmapped_listing_headers = [header for header in listing_headers if header not in LISTING_COLUMNS and header not in LEGACY_ONLY_FIELDS]
    unmapped_detail_headers = [header for header in detail_headers if header not in DETAIL_COLUMNS and header not in DETAIL_SOURCE_ALIASES and header != "접수일"]
    return {
        "ok": True,
        "source": {
            "listing": str(listing_input),
            "listing_sha256": listing_sha,
            "detail": str(detail_input),
            "detail_sha256": detail_sha,
        },
        "output": {
            "listing": str(listing_output),
            "listing_sha256": sha256_file(listing_output),
            "detail": str(detail_output),
            "detail_sha256": sha256_file(detail_output),
        },
        "rows": {"listing": listing_summary["rows"], "detail": detail_summary["rows"]},
        "mapping": {
            "공시가격(만원)": "공시가격",
            "매물상세.접수일": "매물.접수일",
            "학군(분)": "매물상세.학군상세 (원문 보존)",
            "지역의 명시적 읍면동": "지역 + 동네",
        },
        "added_unknown_fields": {
            "listing": [column for column in LISTING_COLUMNS if column not in listing_headers and column != "접수일"],
            "detail": [column for column in DETAIL_COLUMNS if column not in detail_headers and column not in DETAIL_SOURCE_ALIASES.values()],
        },
        "unmapped_headers": {"listing": unmapped_listing_headers, "detail": unmapped_detail_headers},
        "orphan_detail_ids": orphan_detail_ids,
        "warnings": warnings + listing_summary["warnings"] + detail_summary["warnings"],
    }


@contextmanager
def file_lock(target: Path) -> Iterable[None]:
    target.parent.mkdir(parents=True, exist_ok=True)
    lock_path = target.with_name(f".{target.name}.lock")
    with lock_path.open("a+", encoding="utf-8") as lock_handle:
        fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(lock_handle.fileno(), fcntl.LOCK_UN)


def write_csv_new(path: Path, columns: list[str]) -> None:
    if path.exists():
        fail("기존 파일을 덮어쓰지 않습니다", path=str(path))
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, raw_temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".init", dir=path.parent)
    temporary = Path(raw_temporary)
    try:
        os.chmod(temporary, 0o600)
        with os.fdopen(descriptor, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=columns, lineterminator="\n")
            writer.writeheader()
            handle.flush()
            os.fsync(handle.fileno())
        os.link(temporary, path)
        temporary.unlink()
        fsync_directory(path.parent)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise


def restore_displaced_file(displaced: Path, target: Path, recovery: Path, mode: int) -> None:
    """Preserve a displaced original and restore a copy to its public path."""
    os.chmod(displaced, 0o600)
    os.replace(displaced, recovery)
    descriptor, raw_restore = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".restore", dir=target.parent)
    restore = Path(raw_restore)
    try:
        with recovery.open("rb") as source, os.fdopen(descriptor, "wb") as destination:
            shutil.copyfileobj(source, destination)
            destination.flush()
            os.fsync(destination.fileno())
        os.chmod(restore, mode)
        os.replace(restore, target)
        fsync_directory(target.parent)
    except Exception:
        if restore.exists():
            restore.unlink()
        raise


def write_csv_atomic(
    path: Path,
    columns: list[str],
    rows: list[dict[str, str]],
    kind: str,
    expected_sha: str,
) -> tuple[str, str]:
    validate_rows(rows, kind)
    backup_dir = path.parent / ".maemul-backups"
    backup_dir.mkdir(mode=0o700, parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    backup = backup_dir / f"{path.name}.{timestamp}.bak"

    mode = stat.S_IMODE(path.stat().st_mode)
    temporary: Path | None = None
    preserve_temporary = False
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8-sig",
            newline="",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            writer = csv.DictWriter(handle, fieldnames=columns, lineterminator="\n", extrasaction="raise")
            writer.writeheader()
            writer.writerows(rows)
            handle.flush()
            os.fsync(handle.fileno())
        os.chmod(temporary, mode)
        read_csv_table(temporary, kind)
        if atomic_exchange(temporary, path):
            displaced_sha = sha256_file(temporary)
            if displaced_sha != expected_sha:
                try:
                    atomic_exchange(temporary, path)
                except Exception as rollback_error:
                    recovery = backup_dir / f"{path.name}.{timestamp}.conflict-original"
                    try:
                        preserve_temporary = True
                        restore_displaced_file(temporary, path, recovery, mode)
                        temporary = None
                    except Exception as recovery_error:
                        preserve_temporary = temporary is not None and temporary.exists()
                        fail(
                            "충돌 원본 자동 복구에 실패했습니다",
                            path=str(path),
                            displaced_path=str(temporary) if preserve_temporary else None,
                            recovery_path=str(recovery) if recovery.exists() else None,
                            rollback_error=str(rollback_error),
                            recovery_error=str(recovery_error),
                        )
                    fail(
                        "파일 충돌을 감지했고 원본을 복원했습니다",
                        path=str(path),
                        expected_sha=expected_sha,
                        actual_sha=displaced_sha,
                        recovery_path=str(recovery),
                        rollback_error=str(rollback_error),
                    )
                fail(
                    "파일이 쓰기 직전에 변경되었습니다",
                    path=str(path),
                    expected_sha=expected_sha,
                    actual_sha=displaced_sha,
                )
            try:
                os.chmod(temporary, 0o600)
                os.replace(temporary, backup)
            except Exception as backup_error:
                # The displaced file is still the confirmed original. Restore
                # it if securing the backup cannot be completed.
                if temporary.exists():
                    try:
                        atomic_exchange(temporary, path)
                    except Exception as rollback_error:
                        recovery = backup_dir / f"{path.name}.{timestamp}.backup-failure-original"
                        try:
                            preserve_temporary = True
                            restore_displaced_file(temporary, path, recovery, mode)
                            temporary = None
                        except Exception as recovery_error:
                            preserve_temporary = temporary is not None and temporary.exists()
                            fail(
                                "백업 실패 후 원본 자동 복구에 실패했습니다",
                                path=str(path),
                                displaced_path=str(temporary) if preserve_temporary else None,
                                recovery_path=str(recovery) if recovery.exists() else None,
                                backup_error=str(backup_error),
                                rollback_error=str(rollback_error),
                                recovery_error=str(recovery_error),
                            )
                        fail(
                            "백업 실패를 감지했고 원본을 복원했습니다",
                            path=str(path),
                            recovery_path=str(recovery),
                            backup_error=str(backup_error),
                            rollback_error=str(rollback_error),
                        )
                raise
            temporary = None
        else:
            # Portable fallback: recheck immediately before replacement. Native
            # exchange above provides the stronger compare-and-swap guarantee.
            require_expected_hash(path, expected_sha)
            descriptor = os.open(backup, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with path.open("rb") as source, os.fdopen(descriptor, "wb") as destination:
                shutil.copyfileobj(source, destination)
                destination.flush()
                os.fsync(destination.fileno())
            backup_sha = sha256_file(backup)
            if backup_sha != expected_sha:
                fail(
                    "백업 중 원본이 변경되었습니다",
                    path=str(path),
                    expected_sha=expected_sha,
                    actual_sha=backup_sha,
                    backup=str(backup),
                )
            require_expected_hash(path, expected_sha)
            os.replace(temporary, path)
            temporary = None
        fsync_directory(path.parent)
        fsync_directory(backup_dir)
    except Exception:
        if temporary and temporary.exists() and not preserve_temporary:
            temporary.unlink()
        raise
    return str(backup), sha256_file(path)


def write_workbook_atomic(
    path: Path,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, str]],
    kind: str,
    expected_sha: str,
) -> tuple[str, str]:
    """Rewrite one standard tab while preserving the other workbook tabs."""
    require_openpyxl()
    require_xlsx_path(path)
    validate_rows(rows, kind)
    backup_dir = path.parent / ".maemul-backups"
    backup_dir.mkdir(mode=0o700, parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    backup = backup_dir / f"{path.name}.{timestamp}.bak"
    descriptor, raw_temporary = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=".xlsx", dir=path.parent)
    os.close(descriptor)
    temporary = Path(raw_temporary)
    try:
        shutil.copy2(path, temporary)
        workbook = load_workbook(temporary)
        try:
            if sheet_name not in workbook.sheetnames:
                fail("Excel 매물장에 필요한 탭이 없습니다", path=str(path), sheet=sheet_name)
            sheet = workbook[sheet_name]
            if sheet.max_row > 1:
                sheet.delete_rows(2, sheet.max_row - 1)
            for row_index, row in enumerate(rows, start=2):
                for column_index, field in enumerate(columns, start=1):
                    sheet.cell(row_index, column_index, row[field])
            sheet.auto_filter.ref = f"A1:{sheet.cell(max(len(rows) + 1, 1), len(columns)).coordinate}"
            workbook.save(temporary)
        finally:
            workbook.close()
        read_workbook_snapshot(temporary, kind, sheet_name)

        require_expected_hash(path, expected_sha)
        descriptor = os.open(backup, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with path.open("rb") as source, os.fdopen(descriptor, "wb") as destination:
            shutil.copyfileobj(source, destination)
            destination.flush()
            os.fsync(destination.fileno())
        if sha256_file(backup) != expected_sha:
            fail("백업 중 Excel 원본이 변경되었습니다", path=str(path), backup=str(backup))
        require_expected_hash(path, expected_sha)
        os.chmod(temporary, stat.S_IMODE(path.stat().st_mode))
        os.replace(temporary, path)
        fsync_directory(path.parent)
        fsync_directory(backup_dir)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise
    return str(backup), sha256_file(path)


def empty_profile_store() -> dict[str, Any]:
    return {"version": 1, "active_profile": None, "profiles": {}}


def next_profile_name(profiles: dict[str, Any], base: str) -> str:
    """Return a stable, human-readable profile name without overwriting one."""
    base = base.strip()
    if not base:
        fail("프로필 이름 기준은 비어 있을 수 없습니다")
    if base not in profiles:
        return base
    suffix = 2
    while f"{base} {suffix}" in profiles:
        suffix += 1
    return f"{base} {suffix}"


def load_profile_store(path: Path) -> dict[str, Any]:
    if not path.exists():
        return empty_profile_store()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        fail("프로필 저장소를 읽을 수 없습니다", path=str(path), error=str(exc))
    if data.get("version") != 1 or not isinstance(data.get("profiles"), dict):
        fail("프로필 저장소 형식이 올바르지 않습니다", path=str(path))
    return data


def write_profile_store(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, raw_temp = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    temporary = Path(raw_temp)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.chmod(temporary, 0o600)
        os.replace(temporary, path)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise


def command_profile(args: argparse.Namespace, store_path: Path) -> dict[str, Any]:
    if args.profile_action in {"list", "show", "next-name"}:
        data = load_profile_store(store_path)
    else:
        with file_lock(store_path):
            data = load_profile_store(store_path)
            if args.profile_action == "start-new":
                onboarding = data.get("onboarding")
                if onboarding:
                    previous_name = onboarding.get("previous_active_profile")
                    return {
                        "ok": True,
                        "started": False,
                        "profile_store": str(store_path),
                        "active_profile": None,
                        "onboarding": onboarding,
                        "previous_profile": data["profiles"].get(previous_name),
                        "profiles_preserved": len(data["profiles"]),
                    }
                previous_name = data.get("active_profile")
                onboarding = {
                    "mode": "new-ledger",
                    "previous_active_profile": previous_name,
                    "started_at": utc_now(),
                }
                data["onboarding"] = onboarding
                data["active_profile"] = None
                write_profile_store(store_path, data)
                return {
                    "ok": True,
                    "started": True,
                    "profile_store": str(store_path),
                    "active_profile": None,
                    "onboarding": onboarding,
                    "previous_profile": data["profiles"].get(previous_name),
                    "profiles_preserved": len(data["profiles"]),
                }
            if args.profile_action == "cancel-new":
                onboarding = data.pop("onboarding", None)
                if not onboarding:
                    return {
                        "ok": True,
                        "cancelled": False,
                        "profile_store": str(store_path),
                        "active_profile": data.get("active_profile"),
                        "profile": data["profiles"].get(data.get("active_profile")),
                    }
                previous_name = onboarding.get("previous_active_profile")
                restored_name = previous_name if previous_name in data["profiles"] else None
                data["active_profile"] = restored_name
                write_profile_store(store_path, data)
                return {
                    "ok": True,
                    "cancelled": True,
                    "profile_store": str(store_path),
                    "active_profile": restored_name,
                    "profile": data["profiles"].get(restored_name),
                    "profiles_preserved": len(data["profiles"]),
                }
            if args.profile_action == "set":
                onboarding = data.get("onboarding")
                if args.name in data["profiles"]:
                    if onboarding and onboarding.get("mode") in {"new-sheet", "new-ledger"}:
                        fail(
                            "새 매물장 시작 중에는 기존 프로필을 덮어쓰지 않습니다",
                            name=args.name,
                            hint="profile next-name으로 새 이름을 만든 뒤 저장하세요",
                        )
                    if not args.replace:
                        fail("같은 이름의 프로필이 이미 있습니다", name=args.name, hint="확인 후 --replace를 사용하세요")
                previous_name = onboarding.get("previous_active_profile") if onboarding else None
                profile: dict[str, Any] = {
                    "name": args.name,
                    "access": args.access,
                    "label": args.label or args.name,
                    "updated_at": utc_now(),
                }
                if args.access == "local-csv":
                    if not args.listing:
                        fail("local-csv 프로필에는 --listing이 필요합니다")
                    listing = resolve_path(args.listing)
                    read_csv_table(listing, "listing")
                    profile["listing_path"] = str(listing)
                    if args.detail:
                        detail = resolve_path(args.detail)
                        read_csv_table(detail, "detail")
                        profile["detail_path"] = str(detail)
                    else:
                        profile["detail_path"] = None
                elif args.access == "local-xlsx":
                    if not args.workbook:
                        fail("local-xlsx 프로필에는 --workbook이 필요합니다")
                    workbook_path = resolve_path(args.workbook)
                    read_workbook_snapshot(workbook_path, "listing", args.listing_sheet)
                    read_workbook_snapshot(workbook_path, "detail", args.detail_sheet)
                    profile.update(
                        workbook_path=str(workbook_path),
                        listing_sheet=args.listing_sheet,
                        detail_sheet=args.detail_sheet,
                    )
                else:
                    if not args.sheet_id or not args.connector or not args.account:
                        fail("google-sheet 프로필에는 --sheet-id, --connector, --account가 필요합니다")
                    profile.update(
                        sheet_id=args.sheet_id,
                        spreadsheet_url=args.spreadsheet_url or f"https://docs.google.com/spreadsheets/d/{args.sheet_id}/edit",
                        connector=args.connector,
                        account=args.account,
                        listing_sheet=args.listing_sheet,
                        detail_sheet=args.detail_sheet,
                    )
                if previous_name and previous_name in data["profiles"] and previous_name != args.name:
                    profile["previous_profile"] = previous_name
                data["profiles"][args.name] = profile
                if args.activate or not data.get("active_profile"):
                    data["active_profile"] = args.name
                if onboarding and data["active_profile"] == args.name:
                    data.pop("onboarding", None)
                write_profile_store(store_path, data)
                return {
                    "ok": True,
                    "profile_store": str(store_path),
                    "profile": profile,
                    "active_profile": data["active_profile"],
                    "previous_profile": data["profiles"].get(previous_name),
                }
            if args.profile_action == "activate":
                if args.name not in data["profiles"]:
                    fail("프로필을 찾을 수 없습니다", name=args.name)
                data["active_profile"] = args.name
                cancelled_onboarding = data.pop("onboarding", None)
                write_profile_store(store_path, data)
                return {
                    "ok": True,
                    "active_profile": args.name,
                    "profile": data["profiles"][args.name],
                    "cancelled_onboarding": bool(cancelled_onboarding),
                }

    if args.profile_action == "list":
        return {
            "ok": True,
            "profile_store": str(store_path),
            "active_profile": data.get("active_profile"),
            "onboarding": data.get("onboarding"),
            "profiles": list(data["profiles"].values()),
        }

    if args.profile_action == "next-name":
        return {
            "ok": True,
            "profile_store": str(store_path),
            "name": next_profile_name(data["profiles"], args.base),
        }

    name = args.name or data.get("active_profile")
    if not name:
        return {
            "ok": True,
            "profile_store": str(store_path),
            "active_profile": None,
            "onboarding": data.get("onboarding"),
            "profile": None,
        }
    profile = data["profiles"].get(name)
    if not profile:
        fail("활성 프로필이 저장소에 없습니다", name=name)
    result = dict(profile)
    if profile["access"] == "local-csv":
        result["listing_exists"] = Path(profile["listing_path"]).is_file()
        detail_path = profile.get("detail_path")
        result["detail_exists"] = bool(detail_path and Path(detail_path).is_file())
    elif profile["access"] == "local-xlsx":
        result["workbook_exists"] = Path(profile["workbook_path"]).is_file()
    return {
        "ok": True,
        "profile_store": str(store_path),
        "active_profile": name,
        "onboarding": data.get("onboarding"),
        "profile": result,
    }


def compare_value(raw: str, criterion: dict[str, Any]) -> bool | None:
    if is_unknown(raw):
        return None
    field = criterion["field"]
    operator = criterion["op"]
    value = criterion.get("value")
    if operator not in SUPPORTED_OPERATORS:
        fail("지원하지 않는 검색 연산자입니다", operator=operator)

    if field in NUMERIC_FIELDS:
        left: Any = parse_number(raw, field)
        if operator == "between":
            if not isinstance(value, list) or len(value) != 2:
                fail("between 값은 두 원소 배열이어야 합니다", field=field)
            right: Any = [parse_number(value[0], field), parse_number(value[1], field)]
        elif operator in {"in", "not-in"}:
            if not isinstance(value, list):
                fail("in/not-in 값은 배열이어야 합니다", field=field)
            right = [parse_number(item, field) for item in value]
        else:
            right = parse_number(value, field)
    else:
        left = raw.strip().casefold()
        if operator == "between":
            if not isinstance(value, list) or len(value) != 2:
                fail("between 값은 두 원소 배열이어야 합니다", field=field)
            right = [str(value[0]).strip().casefold(), str(value[1]).strip().casefold()]
        elif operator in {"in", "not-in"}:
            if not isinstance(value, list):
                fail("in/not-in 값은 배열이어야 합니다", field=field)
            right = [str(item).strip().casefold() for item in value]
        else:
            right = str(value).strip().casefold()

    if operator == "eq":
        return left == right
    if operator == "ne":
        return left != right
    if operator == "in":
        return left in right
    if operator == "not-in":
        return left not in right
    if operator == "contains":
        return right in left
    if operator == "lte":
        return left <= right
    if operator == "gte":
        return left >= right
    return right[0] <= left <= right[1]


def validate_criteria(criteria: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    if not isinstance(criteria, dict):
        fail("검색 조건은 JSON 객체여야 합니다")
    hard = criteria.get("hard", [])
    soft = criteria.get("soft", [])
    if not isinstance(hard, list) or not isinstance(soft, list):
        fail("hard와 soft는 배열이어야 합니다")
    for criterion in [*hard, *soft]:
        if not isinstance(criterion, dict):
            fail("각 검색 조건은 객체여야 합니다")
        if criterion.get("field") not in LISTING_COLUMNS:
            fail("검색 조건 필드가 스키마에 없습니다", field=criterion.get("field"))
        if criterion.get("op") not in SUPPORTED_OPERATORS:
            fail("지원하지 않는 검색 연산자입니다", operator=criterion.get("op"))
        if "value" not in criterion:
            fail("검색 조건에 value가 없습니다", criterion=criterion)
        if criterion.get("value") is None:
            fail("검색 조건 value는 null일 수 없습니다", criterion=criterion)
        if criterion["field"] in NUMERIC_FIELDS and criterion["op"] == "contains":
            fail("숫자 필드에는 contains를 사용할 수 없습니다", field=criterion["field"])
    return hard, soft


def classify_row(row: dict[str, str], hard: list[dict[str, Any]]) -> tuple[str, list[str], list[str]]:
    unknown: list[str] = []
    failed: list[str] = []
    for criterion in hard:
        result = compare_value(row[criterion["field"]], criterion)
        label = criterion.get("label") or criterion["field"]
        if result is None:
            unknown.append(label)
        elif not result:
            failed.append(label)
    if failed:
        return "excluded", unknown, failed
    if unknown:
        return "needs_verification", unknown, failed
    return "match", unknown, failed


def transaction_price(row: dict[str, str]) -> tuple[float, float]:
    infinity = float("inf")
    transaction = row["거래"]
    if transaction == "매매":
        return (infinity if is_unknown(row["매매가(만원)"]) else float(row["매매가(만원)"]), 0.0)
    if transaction == "전세":
        return (infinity if is_unknown(row["보증금(만원)"]) else float(row["보증금(만원)"]), 0.0)
    deposit = infinity if is_unknown(row["보증금(만원)"]) else float(row["보증금(만원)"])
    rent = infinity if is_unknown(row["월세(만원)"]) else float(row["월세(만원)"])
    return deposit, rent


def receipt_rank(row: dict[str, str]) -> int:
    raw = row["접수일"]
    if is_unknown(raw):
        return 0
    try:
        return int(raw.replace("-", ""))
    except ValueError:
        return 0


def decorated_row(row: dict[str, str], soft: list[dict[str, Any]], unknown: list[str]) -> dict[str, Any]:
    satisfied: list[str] = []
    soft_unknown: list[str] = []
    for criterion in soft:
        result = compare_value(row[criterion["field"]], criterion)
        label = criterion.get("label") or criterion["field"]
        if result is True:
            satisfied.append(label)
        elif result is None:
            soft_unknown.append(label)
    return {
        "row": row,
        "soft_score": len(satisfied),
        "soft_satisfied": satisfied,
        "soft_unknown": soft_unknown,
        "hard_unknown": unknown,
    }


def search_rows(
    rows: list[dict[str, str]],
    hard: list[dict[str, Any]],
    soft: list[dict[str, Any]],
    include_hold: bool,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    matches: list[dict[str, Any]] = []
    possible: list[dict[str, Any]] = []
    for row in rows:
        if row["상태"] == "완료" or (row["상태"] == "보류" and not include_hold):
            continue
        classification, unknown, _ = classify_row(row, hard)
        if classification == "excluded":
            continue
        decorated = decorated_row(row, soft, unknown)
        if classification == "match":
            matches.append(decorated)
        else:
            possible.append(decorated)

    def sort_key(item: dict[str, Any]) -> tuple[Any, ...]:
        row = item["row"]
        price_a, price_b = transaction_price(row)
        return (-item["soft_score"], price_a, price_b, -receipt_rank(row), row["번호"])

    matches.sort(key=sort_key)
    possible.sort(key=sort_key)
    return matches, possible


def relaxation_counts(
    rows: list[dict[str, str]],
    hard: list[dict[str, Any]],
    include_hold: bool,
) -> list[dict[str, Any]]:
    """Count all one-condition relaxations in one pass over the rows."""
    counts = [[0, 0] for _ in hard]
    for row in rows:
        if row["상태"] == "완료" or (row["상태"] == "보류" and not include_hold):
            continue
        results = [compare_value(row[criterion["field"]], criterion) for criterion in hard]
        for removed_index in range(len(hard)):
            remaining = results[:removed_index] + results[removed_index + 1 :]
            if False in remaining:
                continue
            counts[removed_index][1 if None in remaining else 0] += 1
    return [
        {
            "removed": criterion.get("label") or criterion["field"],
            "verified_count": counts[index][0],
            "needs_verification_count": counts[index][1],
        }
        for index, criterion in enumerate(hard)
    ]


def read_command_table(args: argparse.Namespace, kind: str) -> tuple[Path, str | None, list[dict[str, str]], dict[str, Any], str]:
    workbook_value = getattr(args, "workbook", None)
    if workbook_value:
        path = resolve_path(workbook_value)
        sheet_name = getattr(args, "sheet", None) or ("매물" if kind == "listing" else "매물상세")
        rows, summary, snapshot_sha = read_workbook_snapshot(path, kind, sheet_name)
        return path, sheet_name, rows, summary, snapshot_sha
    path = resolve_path(args.file)
    rows, summary, snapshot_sha = read_csv_snapshot(path, kind)
    return path, None, rows, summary, snapshot_sha


def target_metadata(path: Path, sheet_name: str | None) -> dict[str, Any]:
    if sheet_name:
        return {"workbook": str(path), "sheet": sheet_name}
    return {"file": str(path)}


def command_search(args: argparse.Namespace) -> dict[str, Any]:
    path, sheet_name, rows, summary, snapshot_sha = read_command_table(args, "listing")
    criteria = parse_json(args.criteria_json, "검색 조건")
    hard, soft = validate_criteria(criteria)
    matches, possible = search_rows(rows, hard, soft, args.include_hold)

    relaxations = relaxation_counts(rows, hard, args.include_hold)

    return {
        "ok": True,
        **target_metadata(path, sheet_name),
        "sha256": snapshot_sha,
        "validated_rows": summary["rows"],
        "matches_count": len(matches),
        "needs_verification_count": len(possible),
        "matches": matches[: args.limit],
        "needs_verification": possible[: args.limit],
        "relaxations": relaxations,
    }


def require_expected_hash(path: Path, expected: str) -> None:
    actual = sha256_file(path)
    if actual != expected:
        fail("파일이 확인 이후 변경되었습니다", path=str(path), expected_sha=expected, actual_sha=actual)


def find_row(rows: list[dict[str, str]], item_id: str) -> tuple[int, dict[str, str]]:
    for index, row in enumerate(rows):
        if row["번호"] == item_id:
            return index, row
    fail("번호에 해당하는 행을 찾을 수 없습니다", item_id=item_id)


def generated_id(rows: list[dict[str, str]]) -> str:
    numbers: list[int] = []
    for row in rows:
        item_id = row["번호"]
        if item_id.startswith("P") and item_id[1:].isdigit():
            numbers.append(int(item_id[1:]))
    return f"P{(max(numbers, default=0) + 1):03d}"


def command_mutation(args: argparse.Namespace) -> dict[str, Any]:
    raw_target = args.workbook if getattr(args, "workbook", None) else args.file
    path = resolve_path(raw_target)
    with file_lock(path):
        path, sheet_name, rows, _, snapshot_sha = read_command_table(args, "listing")
        if snapshot_sha != args.expected_sha:
            fail(
                "파일이 확인 이후 변경되었습니다",
                path=str(path),
                expected_sha=args.expected_sha,
                actual_sha=snapshot_sha,
            )
        if args.command == "add":
            payload = parse_json(args.record_json, "매물")
            records = payload if isinstance(payload, list) else [payload]
            if not records or len(records) > 1000 or any(not isinstance(record, dict) for record in records):
                fail("매물 JSON은 객체 또는 1~1000개의 객체 배열이어야 합니다")
            added: list[dict[str, str]] = []
            warnings: list[dict[str, Any]] = []
            for record in records:
                unknown_fields = sorted(set(record) - set(LISTING_COLUMNS))
                if unknown_fields:
                    fail("스키마에 없는 컬럼입니다", fields=unknown_fields)
                row = {column: "?" for column in LISTING_COLUMNS}
                for field in ("매매가(만원)", "보증금(만원)", "월세(만원)"):
                    row[field] = ""
                row.update({key: normalize_input_value(key, value) for key, value in record.items()})
                row["번호"] = row["번호"] if not is_unknown(row["번호"]) else generated_id(rows)
                if "상태" in record and row["상태"] not in STATUS_VALUES:
                    fail("명시한 상태 값이 올바르지 않습니다", value=row["상태"], allowed=sorted(STATUS_VALUES))
                missing_core = [field for field in ("종류", "거래", "지역") if is_unknown(row[field])]
                if "상태" not in record:
                    row["상태"] = "보류" if missing_core else "진행"
                if missing_core:
                    warnings.append({"item_id": row["번호"], "fields": missing_core, "warning": "검색 핵심값 미확인으로 보류"})
                if "접수일" not in record:
                    row["접수일"] = date.today().isoformat()
                if any(existing["번호"] == row["번호"] for existing in rows):
                    fail("이미 존재하는 번호입니다", item_id=row["번호"])
                rows.append(row)
                added.append(row)
            changed = added[0] if isinstance(payload, dict) else added
        else:
            index, current = find_row(rows, args.id)
            before = dict(current)
            if args.command == "complete":
                if current["상태"] == "완료":
                    fail("이미 완료 상태입니다", item_id=args.id)
                current["상태"] = "완료"
            else:
                changes = parse_json(args.changes_json, "변경값")
                if not isinstance(changes, dict):
                    fail("변경값 JSON은 객체여야 합니다")
                unknown_fields = sorted(set(changes) - set(LISTING_COLUMNS))
                if unknown_fields:
                    fail("스키마에 없는 컬럼입니다", fields=unknown_fields)
                if "번호" in changes:
                    fail("번호는 update로 변경할 수 없습니다")
                current.update({key: normalize_input_value(key, value) for key, value in changes.items()})
            rows[index] = current
            changed = {"before": before, "after": current}

            warnings = []

        if sheet_name:
            backup, new_hash = write_workbook_atomic(path, sheet_name, LISTING_COLUMNS, rows, "listing", args.expected_sha)
        else:
            backup, new_hash = write_csv_atomic(path, LISTING_COLUMNS, rows, "listing", args.expected_sha)
        return {
            "ok": True,
            **target_metadata(path, sheet_name),
            "backup": backup,
            "sha256": new_hash,
            "changed": changed,
            "warnings": warnings,
        }


def command_detail_upsert(args: argparse.Namespace) -> dict[str, Any]:
    raw_target = args.workbook if getattr(args, "workbook", None) else args.file
    path = resolve_path(raw_target)
    with file_lock(path):
        path, sheet_name, rows, _, snapshot_sha = read_command_table(args, "detail")
        if snapshot_sha != args.expected_sha:
            fail(
                "파일이 확인 이후 변경되었습니다",
                path=str(path),
                expected_sha=args.expected_sha,
                actual_sha=snapshot_sha,
            )
        changes = parse_json(args.changes_json, "상세 변경값")
        if not isinstance(changes, dict):
            fail("상세 변경값 JSON은 객체여야 합니다")
        unknown_fields = sorted(set(changes) - set(DETAIL_COLUMNS))
        if unknown_fields:
            fail("상세 스키마에 없는 컬럼입니다", fields=unknown_fields)
        if "번호" in changes and normalize_input_value("번호", changes["번호"]) != args.id:
            fail("상세 번호는 대상 번호와 같아야 합니다")
        try:
            index, current = find_row(rows, args.id)
            before: dict[str, str] | None = dict(current)
        except ToolError:
            index = len(rows)
            current = {column: "?" for column in DETAIL_COLUMNS}
            current["번호"] = args.id
            before = None
        current.update({key: normalize_input_value(key, value) for key, value in changes.items()})
        current["번호"] = args.id
        if index == len(rows):
            rows.append(current)
        else:
            rows[index] = current
        if sheet_name:
            backup, new_hash = write_workbook_atomic(path, sheet_name, DETAIL_COLUMNS, rows, "detail", args.expected_sha)
        else:
            backup, new_hash = write_csv_atomic(path, DETAIL_COLUMNS, rows, "detail", args.expected_sha)
        return {
            "ok": True,
            **target_metadata(path, sheet_name),
            "backup": backup,
            "sha256": new_hash,
            "changed": {"before": before, "after": current},
        }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--profile-store", default=None, help="Override profile store path for tests or isolated use")
    subparsers = parser.add_subparsers(dest="command", required=True)

    def add_table_target(command_parser: argparse.ArgumentParser) -> None:
        target = command_parser.add_mutually_exclusive_group(required=True)
        target.add_argument("--file", help="Legacy/local CSV path")
        target.add_argument("--workbook", help="Local .xlsx workbook path")
        command_parser.add_argument("--sheet", help="Workbook sheet name; defaults to 매물 or 매물상세")

    profile = subparsers.add_parser("profile")
    profile_actions = profile.add_subparsers(dest="profile_action", required=True)
    profile_actions.add_parser("list")
    profile_show = profile_actions.add_parser("show")
    profile_show.add_argument("--name")
    profile_actions.add_parser("start-new")
    profile_actions.add_parser("cancel-new")
    profile_next_name = profile_actions.add_parser("next-name")
    profile_next_name.add_argument("--base", default="기본매물장")
    profile_activate = profile_actions.add_parser("activate")
    profile_activate.add_argument("--name", required=True)
    profile_set = profile_actions.add_parser("set")
    profile_set.add_argument("--name", required=True)
    profile_set.add_argument("--access", choices=["local-csv", "local-xlsx", "google-sheet"], required=True)
    profile_set.add_argument("--listing")
    profile_set.add_argument("--detail")
    profile_set.add_argument("--workbook")
    profile_set.add_argument("--sheet-id")
    profile_set.add_argument("--spreadsheet-url")
    profile_set.add_argument("--connector")
    profile_set.add_argument("--account")
    profile_set.add_argument("--listing-sheet", default="매물")
    profile_set.add_argument("--detail-sheet", default="매물상세")
    profile_set.add_argument("--label")
    profile_set.add_argument("--activate", action="store_true")
    profile_set.add_argument("--replace", action="store_true")

    init_ledger = subparsers.add_parser("init-ledger")
    init_ledger.add_argument("--listing", required=True)
    init_ledger.add_argument("--detail", required=True)

    init_workbook = subparsers.add_parser("init-workbook")
    init_workbook.add_argument("--workbook", required=True)
    init_workbook.add_argument("--listing-sheet", default="매물")
    init_workbook.add_argument("--detail-sheet", default="매물상세")

    normalize_pair = subparsers.add_parser("normalize-pair")
    normalize_pair.add_argument("--listing-input", required=True)
    normalize_pair.add_argument("--detail-input", required=True)
    normalize_pair.add_argument("--output-dir", required=True)

    validate = subparsers.add_parser("validate")
    add_table_target(validate)
    validate.add_argument("--kind", choices=["listing", "detail"], required=True)

    file_hash = subparsers.add_parser("hash")
    file_hash.add_argument("--file", required=True)

    inspect = subparsers.add_parser("inspect")
    add_table_target(inspect)
    inspect.add_argument("--id", required=True)

    search = subparsers.add_parser("search")
    add_table_target(search)
    search.add_argument("--criteria-json", required=True)
    search.add_argument("--limit", type=int, default=10)
    search.add_argument("--include-hold", action="store_true")

    add = subparsers.add_parser("add")
    add_table_target(add)
    add.add_argument("--record-json", required=True)
    add.add_argument("--expected-sha", required=True)

    update = subparsers.add_parser("update")
    add_table_target(update)
    update.add_argument("--id", required=True)
    update.add_argument("--changes-json", required=True)
    update.add_argument("--expected-sha", required=True)

    complete = subparsers.add_parser("complete")
    add_table_target(complete)
    complete.add_argument("--id", required=True)
    complete.add_argument("--expected-sha", required=True)

    detail = subparsers.add_parser("detail-upsert")
    add_table_target(detail)
    detail.add_argument("--id", required=True)
    detail.add_argument("--changes-json", required=True)
    detail.add_argument("--expected-sha", required=True)
    return parser


def run(args: argparse.Namespace) -> dict[str, Any]:
    store_path = resolve_path(args.profile_store) if args.profile_store else default_profile_store()
    if args.command == "profile":
        return command_profile(args, store_path)
    if args.command == "normalize-pair":
        return command_normalize_pair(args)
    if args.command == "init-workbook":
        return command_init_workbook(args)
    if args.command == "init-ledger":
        listing = resolve_path(args.listing)
        detail = resolve_path(args.detail)
        if listing == detail:
            fail("검색용 파일과 상세 파일은 달라야 합니다")
        listing_exists = listing.exists()
        detail_exists = detail.exists()
        if listing_exists and detail_exists:
            fail("두 파일이 이미 존재하므로 덮어쓰지 않습니다", listing=str(listing), detail=str(detail))
        if listing_exists:
            read_csv_table(listing, "listing")
            write_csv_new(detail, DETAIL_COLUMNS)
        elif detail_exists:
            read_csv_table(detail, "detail")
            write_csv_new(listing, LISTING_COLUMNS)
        else:
            write_csv_new(listing, LISTING_COLUMNS)
            try:
                write_csv_new(detail, DETAIL_COLUMNS)
            except Exception:
                # Preserve the completed listing rather than risk deleting a path
                # another process may have replaced after creation. Re-running
                # init-ledger will validate it and create only the missing detail.
                raise
        return {
            "ok": True,
            "listing": str(listing),
            "detail": str(detail),
            "listing_sha256": sha256_file(listing),
            "detail_sha256": sha256_file(detail),
        }
    if args.command == "validate":
        path, sheet_name, _, summary, snapshot_sha = read_command_table(args, args.kind)
        return {"ok": True, **target_metadata(path, sheet_name), "kind": args.kind, "sha256": snapshot_sha, **summary}
    if args.command == "hash":
        path = resolve_path(args.file)
        if not path.is_file():
            fail("파일을 찾을 수 없습니다", path=str(path))
        return {"ok": True, "file": str(path), "sha256": sha256_file(path)}
    if args.command == "inspect":
        path, sheet_name, rows, _, snapshot_sha = read_command_table(args, "listing")
        _, row = find_row(rows, args.id)
        return {"ok": True, **target_metadata(path, sheet_name), "sha256": snapshot_sha, "row": row}
    if args.command == "search":
        if args.limit < 1 or args.limit > 100:
            fail("limit은 1부터 100 사이여야 합니다")
        return command_search(args)
    if args.command in {"add", "update", "complete"}:
        return command_mutation(args)
    if args.command == "detail-upsert":
        return command_detail_upsert(args)
    fail("알 수 없는 명령입니다", command=args.command)


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        emit(run(args))
        return 0
    except ToolError as exc:
        try:
            details = json.loads(str(exc))
        except json.JSONDecodeError:
            details = {"message": str(exc)}
        emit({"ok": False, "error": details})
        return 2
    except Exception as exc:  # pragma: no cover - defensive CLI boundary
        emit({"ok": False, "error": {"message": "예상하지 못한 오류", "type": type(exc).__name__, "detail": str(exc)}})
        return 3


if __name__ == "__main__":
    sys.exit(main())
