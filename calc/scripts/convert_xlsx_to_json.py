import json
import sys
from datetime import datetime, timezone

import openpyxl

SRC = sys.argv[1] if len(sys.argv) > 1 else r"..\..\calculator_database.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else r"..\src\data\default_database.json"


def sheet_rows(path, name):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[name]
    rows = [list(r) for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
    header = [str(h).strip() for h in rows[0]]
    return [{header[i]: (v.isoformat() if isinstance(v, datetime) else v) for i, v in enumerate(r) if i < len(header)} for r in rows[1:]]


def build(path):
    db = {
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "yandexUrl": "",
        "sourceFile": "calculator_database.xlsx",
    }
    db["constructions"] = sheet_rows(path, "Constructions")
    comps = sheet_rows(path, "Components")
    grouped = {}
    for c in comps:
        grouped.setdefault(c.get("Category", "OTHER"), []).append(c)
    db["components"] = grouped
    db["additionalWorks"] = sheet_rows(path, "AdditionalWorks")
    db["services"] = sheet_rows(path, "Services")
    db["discounts"] = sheet_rows(path, "Discounts")
    db["coefficients"] = sheet_rows(path, "CalculationCoefficients")
    db["formulas"] = sheet_rows(path, "Formulas")
    db["units"] = sheet_rows(path, "Units")
    return db


if __name__ == "__main__":
    data = build(SRC)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print("wrote", OUT, "with", sum(len(v) for k, v in data.items() if isinstance(v, list)) + sum(len(v) for k, v in data["components"].items()), "records")