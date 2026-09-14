"""
Pull the 21 BCB SGS credit benchmark series (balance, nominal rate, NPL90
for INSS/public/CLT payroll, vehicle, mortgage, unsecured personal, overdraft),
pivot to one row per (month, product), and upsert into a Google Sheet via the
Sheets API using a service account.

Runs on a schedule via .github/workflows/bcb_sgs_sync.yml — same logic as the
original n8n workflow, ported to run on GitHub Actions instead of a local box.
"""
import json
import os
from datetime import date, datetime, timezone
from typing import Any

import gspread
import requests
from google.oauth2.service_account import Credentials

SERIES = [
    # INSS
    {"product": "INSS payroll loans", "metric": "balance", "series": 20578, "unit": "BRL_mn"},
    {"product": "INSS payroll loans", "metric": "nominal_rate_pa", "series": 20746, "unit": "pct_pa"},
    {"product": "INSS payroll loans", "metric": "npl90", "series": 21118, "unit": "pct"},
    # Public payroll
    {"product": "Public-sector payroll loans", "metric": "balance", "series": 20577, "unit": "BRL_mn"},
    {"product": "Public-sector payroll loans", "metric": "nominal_rate_pa", "series": 20745, "unit": "pct_pa"},
    {"product": "Public-sector payroll loans", "metric": "npl90", "series": 21117, "unit": "pct"},
    # CLT / private payroll
    {"product": "Private-sector / CLT payroll loans", "metric": "balance", "series": 20576, "unit": "BRL_mn"},
    {"product": "Private-sector / CLT payroll loans", "metric": "nominal_rate_pa", "series": 20744, "unit": "pct_pa"},
    {"product": "Private-sector / CLT payroll loans", "metric": "npl90", "series": 21116, "unit": "pct"},
    # Vehicle
    {"product": "Vehicle financing", "metric": "balance", "series": 20581, "unit": "BRL_mn"},
    {"product": "Vehicle financing", "metric": "nominal_rate_pa", "series": 20749, "unit": "pct_pa"},
    {"product": "Vehicle financing", "metric": "npl90", "series": 21121, "unit": "pct"},
    # Mortgage
    {"product": "Mortgage", "metric": "balance", "series": 20612, "unit": "BRL_mn"},
    {"product": "Mortgage", "metric": "nominal_rate_pa", "series": 20774, "unit": "pct_pa"},
    {"product": "Mortgage", "metric": "npl90", "series": 21151, "unit": "pct"},
    # Unsecured personal
    {"product": "Unsecured Personal Loans", "metric": "balance", "series": 29965, "unit": "BRL_mn"},
    {"product": "Unsecured Personal Loans", "metric": "nominal_rate_pa", "series": 29974, "unit": "pct_pa"},
    {"product": "Unsecured Personal Loans", "metric": "npl90", "series": 29992, "unit": "pct"},
    # Overdraft
    {"product": "Overdraft", "metric": "balance", "series": 20573, "unit": "BRL_mn"},
    {"product": "Overdraft", "metric": "nominal_rate_pa", "series": 20741, "unit": "pct_pa"},
    {"product": "Overdraft", "metric": "npl90", "series": 21113, "unit": "pct"},
]

SHEET_TAB = "raw_credit_benchmark"
SHEET_HEADERS = [
    "reference_month", "product", "balance_brl_bn", "nominal_rate_pa", "npl90", "updated_at",
]


def bcb_date_range(today: date) -> tuple[str, str]:
    # last completed month
    end = today.replace(day=1)
    end = date(end.year, end.month, 1)
    # step back one day from the 1st of this month to get last day of prior month
    from datetime import timedelta
    end = end - timedelta(days=1)

    # start 3 months before end -> 4 months total
    start_month = end.month - 3
    start_year = end.year
    while start_month <= 0:
        start_month += 12
        start_year -= 1
    start = date(start_year, start_month, 1)

    return start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y")


def fetch_series(series_id: int, date_inicial: str, data_final: str) -> list[dict[str, Any]]:
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{series_id}/dados"
    resp = requests.get(
        url,
        params={"formato": "json", "dataInicial": date_inicial, "dataFinal": data_final},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def normalize_and_pivot(date_inicial: str, data_final: str) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], dict[str, Any]] = {}

    for cfg in SERIES:
        rows = fetch_series(cfg["series"], date_inicial, data_final)
        for row in rows:
            value = float(str(row["valor"]).replace(",", "."))
            if cfg["unit"] == "BRL_mn":
                value = value / 1000  # R$ million -> R$ billion

            day, month, year = row["data"].split("/")
            reference_month = f"{year}-{month}"
            key = (reference_month, cfg["product"])

            if key not in grouped:
                grouped[key] = {
                    "reference_month": reference_month,
                    "product": cfg["product"],
                    "balance_brl_bn": None,
                    "nominal_rate_pa": None,
                    "npl90": None,
                }

            if cfg["metric"] == "balance":
                grouped[key]["balance_brl_bn"] = value
            elif cfg["metric"] == "nominal_rate_pa":
                grouped[key]["nominal_rate_pa"] = value
            elif cfg["metric"] == "npl90":
                grouped[key]["npl90"] = value

    complete = [
        r for r in grouped.values()
        if r["balance_brl_bn"] is not None
        and r["nominal_rate_pa"] is not None
        and r["npl90"] is not None
    ]
    incomplete = [r for r in grouped.values() if r not in complete]
    if incomplete:
        print(f"Skipping {len(incomplete)} incomplete row(s) (not yet published by BCB):")
        for r in incomplete:
            print(f"  {r['reference_month']} / {r['product']}")

    complete.sort(key=lambda r: (r["reference_month"], r["product"]))
    return complete


def open_sheet() -> gspread.Worksheet:
    creds_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    spreadsheet_id = os.environ["SPREADSHEET_ID"]

    creds = Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)

    try:
        ws = spreadsheet.worksheet(SHEET_TAB)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=SHEET_TAB, rows=1000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)

    return ws


def upsert_rows(ws: gspread.Worksheet, rows: list[dict[str, Any]]) -> None:
    existing = ws.get_all_values()
    if not existing:
        ws.append_row(SHEET_HEADERS)
        existing = [SHEET_HEADERS]

    key_index = {(r[0], r[1]): i for i, r in enumerate(existing[1:], start=2) if len(r) >= 2}

    now_iso = datetime.now(timezone.utc).isoformat()

    updates = []
    appends = []

    for row in rows:
        key = (row["reference_month"], row["product"])
        values = [
            row["reference_month"],
            row["product"],
            row["balance_brl_bn"],
            row["nominal_rate_pa"],
            row["npl90"],
            now_iso,
        ]
        if key in key_index:
            updates.append((key_index[key], values))
        else:
            appends.append(values)

    for row_num, values in updates:
        ws.update(f"A{row_num}:F{row_num}", [values])

    if appends:
        ws.append_rows(appends)

    print(f"Updated {len(updates)} row(s), appended {len(appends)} new row(s).")


def main() -> None:
    today = date.today()
    date_inicial, data_final = bcb_date_range(today)
    print(f"Fetching BCB SGS series from {date_inicial} to {data_final}")

    rows = normalize_and_pivot(date_inicial, data_final)
    print(f"{len(rows)} complete (month, product) row(s) ready to write.")

    ws = open_sheet()
    upsert_rows(ws, rows)


if __name__ == "__main__":
    main()
