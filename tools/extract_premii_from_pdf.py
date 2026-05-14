#!/usr/bin/env python3
import argparse
import csv
import re
from pathlib import Path

import pdfplumber


def normalize_fio(text: str) -> str:
    t = re.sub(r"\s+", " ", text or "").strip()
    t = re.sub(r"\s+\d+$", "", t)
    return t


def parse_page(text: str):
    fio_match = re.search(r"Табельный номер\s*(.*?)\s*\(фамилия, имя, отчество\)", text, re.S | re.I)
    dep_match = re.search(r"\(фамилия, имя, отчество\)\s*(.*?)\s*\(структурное подразделение\)", text, re.S | re.I)

    if not fio_match:
        return None

    fio = normalize_fio(fio_match.group(1).replace("\n", " "))
    restaurant = ""
    if dep_match:
        restaurant = re.sub(r"\s+", " ", dep_match.group(1)).strip()

    amounts = re.findall(r"\d[\d ]*\.\d{2}", text)
    if not amounts:
        return None

    amount = float(amounts[-1].replace(" ", ""))

    if re.search(r"В сумме\s*Минус", text, re.I) or re.search(r"\bВзыскание\b", text, re.I):
        amount = -abs(amount)

    return fio, restaurant, amount


def main():
    parser = argparse.ArgumentParser(description="Extract bonus amounts from T-11 PDF files")
    parser.add_argument("pdf", nargs="+", help="Input PDF files")
    parser.add_argument("-o", "--output", required=True, help="Output CSV path")
    args = parser.parse_args()

    totals = {}

    for p in args.pdf:
        path = Path(p)
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                row = parse_page(text)
                if not row:
                    continue
                fio, restaurant, amount = row
                key = (fio, restaurant)
                totals[key] = totals.get(key, 0.0) + amount

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)

    with out.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(["ФИО", "Ресторан", "Премии"])
        for (fio, restaurant), amount in sorted(totals.items()):
            w.writerow([fio, restaurant, f"{amount:.2f}"])

    print(f"Saved {len(totals)} rows to {out}")


if __name__ == "__main__":
    main()
