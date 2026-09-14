"""
Pull the 21 BCB SGS credit benchmark series (balance, nominal rate, NPL90
for INSS/public/CLT payroll, vehicle, mortgage, unsecured personal, overdraft),
pivot to one row per (month, product), and POST them to a Google Apps Script
Web App bound to the target Sheet, which does the actual upsert into cells.

Runs on a schedule via .github/workflows/bcb_sgs_sync.yml — same logic as the
original n8n workflow, ported to run on GitHub Actions instead of a local box.
No Google Cloud project / service account needed: the Apps Script Web App
is deployed under your own Google login directly from the Sheet's
Extensions > Apps Script menu (see scripts/apps_script/Code.gs).
"""
import os
from datetime import date
from typing import Any

import requests

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


def push_to_sheet(rows: list[dict[str, Any]]) -> None:
    web_app_url = os.environ["APPS_SCRIPT_URL"]
    token = os.environ["APPS_SCRIPT_TOKEN"]

    resp = requests.post(
        web_app_url,
        json={"token": token, "rows": rows},
        timeout=30,
    )
    resp.raise_for_status()

    result = resp.json()
    if not result.get("ok"):
        raise RuntimeError(f"Apps Script upsert failed: {result}")

    print(f"Updated {result.get('updated', 0)} row(s), appended {result.get('appended', 0)} new row(s).")


def main() -> None:
    today = date.today()
    date_inicial, data_final = bcb_date_range(today)
    print(f"Fetching BCB SGS series from {date_inicial} to {data_final}")

    rows = normalize_and_pivot(date_inicial, data_final)
    print(f"{len(rows)} complete (month, product) row(s) ready to write.")

    push_to_sheet(rows)


if __name__ == "__main__":
    main()
