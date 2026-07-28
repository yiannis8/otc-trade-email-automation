from __future__ import annotations

import csv
import html
import io
import json
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from email.message import EmailMessage
from email.policy import default
from pathlib import Path
from typing import Any, Iterable

EMAIL_COLUMNS = [
    "Structure",
    "Format",
    "Rates Curve Index",
    "Issuer / Spread",
    "Currency",
    "Size",
    "Bloomberg Ticker 1",
    "Bloomberg Ticker 2",
    "Bloomberg Ticker 3",
    "Bloomberg Ticker 4",
    "Bloomberg Ticker 5",
    "Reoffer / Upfront (%)",
    "Strike Date",
    "Tenor (m)",
    "Autocall (Yes/No)",
    "Frequency",
    "First Observation in (m)",
    "Autocall Trigger Level (%)",
    "Autocall Step Down/Up (%)",
    "Coupon Type",
    "Memory",
    "Coupon Trigger Level (%)",
    "Coupon Step Down/Up (%)",
    "Coupon p.a. (%)",
    "Strike Level (%)",
    "Downside Leverage (%)",
    "Barrier Type",
    "Barrier Level",
    "Status",
    "Error Message",
    "Reference ID",
]

ALIASES = {
    "direction": ["direction", "side"],
    "size": ["size", "notional", "amount"],
    "assets": ["assets", "underlyings", "basket", "underlying assets"],
    "maturity_date": ["maturity date", "maturity", "redemption date"],
    "currency": ["currency", "ccy"],
    "reference_date": ["reference date", "strike date", "initial valuation date"],
    "ko_barrier": ["ko barrier", "autocall barrier", "autocall trigger"],
    "coupon_barrier": ["coupon barrier", "coupon trigger"],
    "ki_barrier": ["ki barrier", "barrier level", "knock-in barrier"],
    "ki_monitoring": ["ki monitoring type", "barrier type", "knock-in monitoring"],
    "put_type": ["put type", "downside type"],
    "put_strike": ["put strike", "strike level"],
    "put_leverage": ["put leverage", "downside leverage"],
    "memory": ["memory", "memory coupon"],
    "coupons": ["coupons per observation", "coupon per observation", "coupon"],
    "coupon_dates": ["coupon dates", "payment dates"],
    "ko_dates": ["ko dates", "autocall dates", "observation dates"],
    "funding_type": ["funding type", "format"],
    "swap_benchmark": ["swap benchmark", "rates curve index", "benchmark"],
    "swap_spread": ["swap spread", "issuer spread", "spread"],
    "issue_date": ["issue date", "settlement date"],
    "reoffer": ["reoffer / upfront (%)", "reoffer", "upfront", "upfront (%)", "bid"],
    "coupon_pa": ["coupon p.a. (%)", "coupon pa", "annual coupon", "coupon p.a."],
    "strike_forward": ["strike forward (bd)", "strike forward", "bd"],
}

COUNTERPARTY_SUBJECT = {
    "MORGAN STANLEY": "MS",
    "MS": "MS",
    "BNPP": "BNP",
    "BNP": "BNP",
    "SOCIETE GENERALE": "SG",
    "SG": "SG",
    "CITI": "CITI",
    "CITIBANK": "CITI",
    "JPM": "JPM",
    "JPMORGAN": "JPM",
    "BARCLAYS": "BARCLAYS",
    "HSBC": "HSBC",
    "UBS": "UBS",
}


@dataclass
class MappingResult:
    fields: dict[str, str]
    source: dict[str, str]
    unmapped_source_fields: dict[str, str] = field(default_factory=dict)
    review_notes: list[str] = field(default_factory=list)
    trade_id: str = ""
    subject: str = "@AGILE"


def _normalise_label(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"[^a-z0-9%(). ]+", "", text)
    return re.sub(r"\s+", " ", text).strip()


def _clean_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d-%b-%y")
    return str(value).strip()


def _read_rows(path_or_bytes: str | Path | bytes, filename: str | None = None) -> list[list[Any]]:
    if isinstance(path_or_bytes, bytes):
        raw = path_or_bytes
        suffix = Path(filename or "input.csv").suffix.lower()
    else:
        path = Path(path_or_bytes)
        raw = path.read_bytes()
        suffix = path.suffix.lower()

    if suffix in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook

        wb = load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
    else:
        text = raw.decode("utf-8-sig", errors="replace")
        rows = list(csv.reader(io.StringIO(text)))

    rows = [row for row in rows if any(_clean_value(value) for value in row)]
    if not rows:
        raise ValueError("The file contains no usable rows.")
    return rows


def load_trade_file(
    path_or_bytes: str | Path | bytes,
    filename: str | None = None,
) -> list[tuple[dict[str, str], str]]:
    """Load one or many trades from CSV/XLSX.

    Supported layouts:
    1. Single key/value trade: field names in column A, values in column B.
    2. Multi-trade matrix: field names in column A, one trade per following column.
    3. Standard table: headers in the first row, one trade per following row.
    """
    rows = _read_rows(path_or_bytes, filename)
    max_columns = max(len(row) for row in rows)
    first_row = rows[0]
    first_cell = _clean_value(first_row[0] if first_row else "")
    first_row_trade_ids = [
        _clean_value(first_row[index] if index < len(first_row) else "")
        for index in range(1, max_columns)
    ]
    labelled_rows = sum(1 for row in rows[1:] if row and _clean_value(row[0]))

    # Matrix layout: blank top-left cell, trade IDs across row 1, fields down column A.
    is_matrix = (
        max_columns > 2
        and not first_cell
        and sum(bool(value) for value in first_row_trade_ids) >= 2
        and labelled_rows >= max(2, (len(rows) - 1) // 2)
    )
    if is_matrix:
        trades: list[tuple[dict[str, str], str]] = []
        for column_index in range(1, max_columns):
            trade_id = _clean_value(first_row[column_index] if column_index < len(first_row) else "")
            source: dict[str, str] = {}
            for row in rows[1:]:
                key = _clean_value(row[0] if row else "")
                value = _clean_value(row[column_index] if column_index < len(row) else "")
                if key:
                    source[key] = value
            if source and (trade_id or any(source.values())):
                trades.append((source, trade_id or f"Trade {column_index}"))
        if not trades:
            raise ValueError("No trades were found in the multi-column file.")
        return trades

    # Single key/value layout.
    key_value_rows = sum(
        1 for row in rows if len(row) >= 2 and _clean_value(row[0])
    )
    is_key_value = key_value_rows >= max(2, len(rows) // 2)
    if is_key_value:
        source: dict[str, str] = {}
        trade_id = ""
        for row_index, row in enumerate(rows):
            key = _clean_value(row[0] if row else "")
            value = _clean_value(row[1] if len(row) > 1 else "")
            if row_index == 0 and not key and value:
                trade_id = value
                continue
            if key:
                source[key] = value
        return [(source, trade_id)]

    # Standard tabular layout: first row contains headers and each later row is a trade.
    headers = [_clean_value(value) for value in rows[0]]
    trades = []
    trade_id_headers = {"trade id", "trade_id", "id", "name", "reference id"}
    for row_number, row in enumerate(rows[1:], start=1):
        source = {
            header: _clean_value(row[index] if index < len(row) else "")
            for index, header in enumerate(headers)
            if header
        }
        if not any(source.values()):
            continue
        trade_id = ""
        for header, value in source.items():
            if _normalise_label(header) in trade_id_headers and value:
                trade_id = value
                break
        trades.append((source, trade_id or f"Trade {row_number}"))

    if not trades:
        raise ValueError("The file does not contain any trade rows.")
    return trades


def load_key_value_file(
    path_or_bytes: str | Path | bytes,
    filename: str | None = None,
) -> tuple[dict[str, str], str]:
    """Backward-compatible helper that returns the first trade only."""
    return load_trade_file(path_or_bytes, filename)[0]


def _canonical_values(source: dict[str, str]) -> tuple[dict[str, str], dict[str, str]]:
    normalized_source = {_normalise_label(key): value for key, value in source.items()}
    used: set[str] = set()
    canonical: dict[str, str] = {}
    for canonical_name, aliases in ALIASES.items():
        for alias in aliases:
            normalized_alias = _normalise_label(alias)
            if normalized_alias in normalized_source:
                canonical[canonical_name] = normalized_source[normalized_alias]
                used.add(normalized_alias)
                break
    unmapped = {
        key: value
        for key, value in source.items()
        if _normalise_label(key) not in used
    }
    return canonical, unmapped


def _split_values(value: str) -> list[str]:
    text = str(value or "").strip()
    if not text or text.upper() in {"NA", "N/A", "NONE", "NULL"}:
        return []
    text = text.strip("[]")
    return [part.strip().strip("'\"") for part in text.split(",") if part.strip().strip("'\"")]


def _parse_date(value: str) -> datetime | None:
    value = str(value or "").strip()
    for date_format in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, date_format)
        except ValueError:
            continue
    return None


def _first_date(value: str) -> datetime | None:
    values = _split_values(value)
    return _parse_date(values[0]) if values else _parse_date(value)


def _months_between(start: datetime | None, end: datetime | None) -> str:
    if not start or not end:
        return ""
    months = (end.year - start.year) * 12 + end.month - start.month
    if end.day < start.day:
        months -= 1
    return str(max(months, 0))


def _percentage_number(value: str) -> str:
    values = _split_values(value)
    text = values[0] if values else str(value or "")
    text = text.strip().replace("%", "").replace(",", "")
    if not text or text.upper() in {"NA", "N/A", "NONE"}:
        return ""
    try:
        number = float(text)
        return f"{number:g}"
    except ValueError:
        return text


def _yes_no(value: str) -> str:
    text = str(value or "").strip().upper()
    if text in {"TRUE", "YES", "Y", "1"}:
        return "Yes"
    if text in {"FALSE", "NO", "N", "0"}:
        return "No"
    return str(value or "").strip()


def _tickers(assets: str) -> list[str]:
    result = []
    for asset in _split_values(assets):
        ticker = re.sub(r"\s+(Index|Equity|Curncy|Comdty|Corp|Govt)$", "", asset, flags=re.I).strip()
        result.append(ticker)
    return result[:5]


def _frequency(dates_value: str) -> tuple[str, str]:
    dates = [_parse_date(value) for value in _split_values(dates_value)]
    dates = [date for date in dates if date]
    if len(dates) < 2:
        return "", ""
    gaps = []
    for left, right in zip(dates, dates[1:]):
        months = (right.year - left.year) * 12 + right.month - left.month
        if right.day < left.day - 5:
            months -= 1
        gaps.append(max(months, 0))
    gap = Counter(gaps).most_common(1)[0][0]
    labels = {1: "Monthly", 3: "Quarterly", 6: "Semi Annually", 12: "Annually"}
    return labels.get(gap, f"Every {gap} months"), str(gap)


def _step(value: str) -> str:
    numbers = []
    for part in _split_values(value):
        try:
            numbers.append(float(part.replace("%", "")))
        except ValueError:
            return "Variable"
    if len(numbers) < 2 or len(set(numbers)) == 1:
        return ""
    differences = [round(right - left, 8) for left, right in zip(numbers, numbers[1:])]
    if len(set(differences)) == 1:
        return f"{differences[0]:g}"
    return "Variable"


def _subject_from_direction(direction: str) -> str:
    upper = str(direction or "").upper()
    for key, code in COUNTERPARTY_SUBJECT.items():
        if key in upper:
            return f"@AGILE {code}"
    token = re.split(r"\s+", upper.strip())[0] if upper.strip() else ""
    return f"@AGILE {token}".strip()


def map_trade(
    source: dict[str, str],
    trade_id: str = "",
    manual_overrides: dict[str, str] | None = None,
) -> MappingResult:
    canonical, unmapped = _canonical_values(source)
    notes: list[str] = []
    tickers = _tickers(canonical.get("assets", ""))
    strike_date = _first_date(canonical.get("reference_date", ""))
    maturity = _parse_date(canonical.get("maturity_date", ""))
    ko_dates = canonical.get("ko_dates", "")
    first_ko = _first_date(ko_dates)
    frequency, _ = _frequency(ko_dates)

    has_ko = bool(_split_values(ko_dates))
    has_put = bool(
        canonical.get("put_type", "")
        and canonical.get("put_type", "").upper() not in {"NA", "N/A", "NONE"}
    )
    if has_ko and has_put:
        structure = "Autocallable & BRC"
    elif has_ko:
        structure = "Autocallable"
    elif has_put:
        structure = "BRC"
    else:
        structure = ""
        notes.append("Structure could not be classified automatically.")

    knock_in = canonical.get("ki_barrier", "")
    knock_in_missing = not _split_values(knock_in)
    barrier_type = "None" if knock_in_missing else (canonical.get("ki_monitoring", "") or "Review")
    barrier_level = "" if knock_in_missing else _percentage_number(knock_in)

    coupon_barrier = canonical.get("coupon_barrier", "")
    coupon_type = "Conditional" if _split_values(coupon_barrier) else "Fixed"

    fields = {column: "" for column in EMAIL_COLUMNS}
    fields.update(
        {
            "Structure": structure,
            "Format": canonical.get("funding_type", ""),
            "Rates Curve Index": canonical.get("swap_benchmark", ""),
            "Issuer / Spread": _percentage_number(canonical.get("swap_spread", "")),
            "Currency": canonical.get("currency", ""),
            "Size": re.sub(r"[^0-9.-]", "", canonical.get("size", "")),
            "Reoffer / Upfront (%)": _percentage_number(canonical.get("reoffer", "")),
            "Strike Date": strike_date.strftime("%d/%m/%Y") if strike_date else "",
            "Tenor (m)": _months_between(strike_date, maturity),
            "Autocall (Yes/No)": "Yes" if has_ko else "No",
            "Frequency": frequency,
            "First Observation in (m)": _months_between(strike_date, first_ko),
            "Autocall Trigger Level (%)": _percentage_number(canonical.get("ko_barrier", "")),
            "Autocall Step Down/Up (%)": _step(canonical.get("ko_barrier", "")),
            "Coupon Type": coupon_type,
            "Memory": _yes_no(canonical.get("memory", "")),
            "Coupon Trigger Level (%)": _percentage_number(coupon_barrier),
            "Coupon Step Down/Up (%)": _step(coupon_barrier),
            "Coupon p.a. (%)": _percentage_number(canonical.get("coupon_pa", "")),
            "Strike Level (%)": _percentage_number(canonical.get("put_strike", "")),
            "Downside Leverage (%)": _percentage_number(canonical.get("put_leverage", "")),
            "Barrier Type": barrier_type,
            "Barrier Level": barrier_level,
        }
    )
    for index, ticker in enumerate(tickers, start=1):
        fields[f"Bloomberg Ticker {index}"] = ticker

    if not fields["Reoffer / Upfront (%)"]:
        notes.append("Reoffer / Upfront (%) is missing and needs a manual input or another source.")
    if not fields["Coupon p.a. (%)"] and canonical.get("coupons"):
        notes.append(
            "Coupons per observation exists, but Coupon p.a. is not annualised automatically. Review this value before sending."
        )
    if has_put and not canonical.get("strike_forward"):
        fields["Status"] = "Error"
        fields["Error Message"] = "Strike Forward (Bd) should be greater than zero"
        notes.append("Strike Forward (Bd) is missing. Add it if Agile requires this field.")

    if manual_overrides:
        for key, value in manual_overrides.items():
            if key in fields:
                fields[key] = str(value)

    return MappingResult(
        fields=fields,
        source=source,
        unmapped_source_fields=unmapped,
        review_notes=notes,
        trade_id=trade_id,
        subject=_subject_from_direction(canonical.get("direction", "")),
    )


def map_trades(trades: Iterable[tuple[dict[str, str], str]]) -> list[MappingResult]:
    return [map_trade(source, trade_id) for source, trade_id in trades]


def _field_rows(fields_or_rows: dict[str, str] | Iterable[dict[str, str]]) -> list[dict[str, str]]:
    if isinstance(fields_or_rows, dict):
        return [fields_or_rows]
    return list(fields_or_rows)


def build_html_table(fields_or_rows: dict[str, str] | Iterable[dict[str, str]]) -> str:
    rows = _field_rows(fields_or_rows)
    headers = "".join(f"<th>{html.escape(column)}</th>" for column in EMAIL_COLUMNS)
    body_rows = []
    for fields in rows:
        values = "".join(
            f"<td>{html.escape(str(fields.get(column, '')))}</td>"
            for column in EMAIL_COLUMNS
        )
        body_rows.append(f"<tr>{values}</tr>")
    body = "".join(body_rows)
    return f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ font-family: Arial, sans-serif; font-size: 12px; color: #111; }}
.table-wrap {{ overflow-x: auto; max-width: 100%; }}
table {{ border-collapse: collapse; white-space: nowrap; }}
th, td {{ border: 1px solid #b7b7b7; padding: 6px 8px; vertical-align: top; }}
th {{ background: #f2f2f2; font-weight: 600; }}
tbody tr:nth-child(even) {{ background: #fafafa; }}
</style>
</head>
<body>
<div class="table-wrap"><table><thead><tr>{headers}</tr></thead><tbody>{body}</tbody></table></div>
</body>
</html>"""


def build_eml(
    to_address: str,
    subject: str,
    fields_or_rows: dict[str, str] | Iterable[dict[str, str]],
) -> bytes:
    message = EmailMessage(policy=default)
    message["To"] = to_address
    message["Subject"] = subject
    message.set_content("This message contains an HTML trade table. Please open it in an HTML-capable email client.")
    message.add_alternative(build_html_table(fields_or_rows), subtype="html")
    return message.as_bytes()


def combined_subject(results: Iterable[MappingResult]) -> str:
    subjects = {result.subject for result in results if result.subject}
    if len(subjects) == 1:
        return next(iter(subjects))
    if not subjects:
        return "@AGILE"
    return "@AGILE MULTI"


def save_outputs(
    results: MappingResult | Iterable[MappingResult],
    output_dir: str | Path,
    to_address: str = "Agile@marexfp.com",
) -> dict[str, Path]:
    result_list = [results] if isinstance(results, MappingResult) else list(results)
    if not result_list:
        raise ValueError("There are no mapped trades to save.")

    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    html_path = output / "generated_trade_email.html"
    eml_path = output / "generated_trade_email.eml"
    json_path = output / "mapping_result.json"
    subject = combined_subject(result_list)
    field_rows = [result.fields for result in result_list]

    html_path.write_text(build_html_table(field_rows), encoding="utf-8")
    eml_path.write_bytes(build_eml(to_address, subject, field_rows))
    json_path.write_text(
        json.dumps(
            {
                "subject": subject,
                "trades": [
                    {
                        "trade_id": result.trade_id,
                        "fields": result.fields,
                        "review_notes": result.review_notes,
                        "unmapped_source_fields": result.unmapped_source_fields,
                    }
                    for result in result_list
                ],
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return {"html": html_path, "eml": eml_path, "json": json_path}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert one or many trade rows into an Agile-style email.")
    parser.add_argument("input_file")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--to", default="Agile@marexfp.com")
    args = parser.parse_args()

    loaded_trades = load_trade_file(args.input_file)
    mapped_results = map_trades(loaded_trades)
    outputs = save_outputs(mapped_results, args.output_dir, args.to)
    print(json.dumps({key: str(value) for key, value in outputs.items()}, indent=2))
    for result in mapped_results:
        if result.review_notes:
            print(f"\nReview notes for {result.trade_id or 'trade'}:")
            for note in result.review_notes:
                print(f"- {note}")
