#!/usr/bin/env python3
"""
Author: Jinhwan Kwon
Last Modified: Feb 2026

Initialize people.csv and pair_scores.csv from an Excel signup file.

- people.csv is preloaded from Excel
- pair_scores.csv is created empty with headers
- phone is the primary key (digits only)
- gender defaults to 'U'
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from typing import Dict, Optional

import pandas as pd


PHONE_RE = re.compile(r"\d+")


def normalize_phone(raw: str) -> Optional[str]:
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    digits = "".join(PHONE_RE.findall(s))
    if len(digits) < 7:
        return None
    return digits


def pick_default_sheet(xlsx_path: str) -> str:
    xl = pd.ExcelFile(xlsx_path)
    best = None
    best_n = -1
    for name in xl.sheet_names:
        m = re.search(r"round\s*(\d+)", name, flags=re.IGNORECASE)
        if m:
            n = int(m.group(1))
            if n > best_n:
                best_n = n
                best = name
    return best or xl.sheet_names[0]


def find_column(df, *candidates: str) -> Optional[str]:
    lowered = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lowered:
            return lowered[cand.lower()]
    for c in df.columns:
        cl = str(c).strip().lower()
        for cand in candidates:
            if cand.lower() in cl:
                return c
    return None


def load_existing_people(path: str) -> Dict[str, Dict[str, str]]:
    people = {}
    if not os.path.exists(path):
        return people

    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            phone = normalize_phone(row.get("phone", ""))
            if phone:
                people[phone] = row
    return people


def main() -> int:
    ap = argparse.ArgumentParser(description="Initialize people.csv and pair_scores.csv from Excel.")
    ap.add_argument("--excel", required=True, help="Path to Excel signup file")
    ap.add_argument("--sheet", default=None, help="Sheet name (default: newest Round N)")
    ap.add_argument("--db-dir", default="db", help="Directory for CSV database")
    ap.add_argument("--overwrite", action="store_true", help="Overwrite existing people.csv rows")
    args = ap.parse_args()

    os.makedirs(args.db_dir, exist_ok=True)

    people_csv = os.path.join(args.db_dir, "people.csv")
    pairs_csv = os.path.join(args.db_dir, "pair_scores.csv")

    sheet = args.sheet or pick_default_sheet(args.excel)
    df = pd.read_excel(args.excel, sheet_name=sheet)

    col_first = find_column(df, "first name", "first")
    col_last = find_column(df, "last name", "last")
    col_email = find_column(df, "email")
    col_phone = find_column(df, "mobile", "phone")

    if not col_phone:
        print("ERROR: Could not find phone/mobile column in Excel.", file=sys.stderr)
        return 1

    existing = load_existing_people(people_csv)
    people_out: Dict[str, Dict[str, str]] = {} if args.overwrite else dict(existing)

    added = 0
    skipped = 0

    for _, row in df.iterrows():
        phone = normalize_phone(row.get(col_phone))
        if not phone:
            continue

        if phone in people_out and not args.overwrite:
            skipped += 1
            continue

        people_out[phone] = {
            "phone": phone,
            "first_name": str(row.get(col_first) or "").strip() if col_first else "",
            "last_name": str(row.get(col_last) or "").strip() if col_last else "",
            "email": str(row.get(col_email) or "").strip() if col_email else "",
            "gender": "U",
        }
        added += 1

    # Write people.csv
    with open(people_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["phone", "first_name", "last_name", "email", "gender"],
        )
        writer.writeheader()
        for phone in sorted(people_out.keys()):
            writer.writerow(people_out[phone])

    # Initialize pair_scores.csv if missing
    if not os.path.exists(pairs_csv):
        with open(pairs_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["phone_a", "phone_b", "score", "locked", "note"],
            )
            writer.writeheader()

    print(f"Initialized database from sheet '{sheet}'")
    print(f"people.csv: {len(people_out)} total ({added} added, {skipped} skipped)")
    print("pair_scores.csv: initialized (empty)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
