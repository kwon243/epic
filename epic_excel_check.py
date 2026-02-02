#!/usr/bin/env python3
"""
Author: Jinhwan Kwon
Last Modified: Feb 2026

epic_excel_check.py

Interactive validator + fixer for response sheets.

What it does:
- Loads a specific sheet (e.g., "Round 4")
- Detects "bad" entries:
  - missing/invalid phone
  - missing first/last name (if columns exist)
  - duplicate phone rows (keeps the first, flags the rest)
- For bad entries:
  - checks people.csv for similar entries and overrides
  - offers manual edit
  - offers delete row
- Identifies new numbers and checks for typos
- Writes updates back to the Excel file

Email column handling:
- If email column is missing, we skip it. (legacy)
- If present but blank, we can fill from people.csv when overriding.
"""

from __future__ import annotations

import csv
import os
import re
import shutil
import time
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd


PHONE_RE = re.compile(r"\d+")

def clean_cell(x) -> str:
    """
    Convert pandas/Excel cell values into a safe string.
    Treat NaN/NA/None as empty.
    """
    if x is None:
        return ""
    try:
        # handles pandas.NaT, numpy.nan, pandas.NA
        if pd.isna(x):
            return ""
    except Exception:
        pass

    s = str(x).strip()
    # Extra guard for cases where something already became a string
    if s.lower() in ("nan", "none", "<na>", "na"):
        return ""
    return s

def normalize_email(email: str) -> str:
    """
    Normalize email for matching.
    Return "" if it isn't a plausible email.
    """
    e = clean_cell(email).lower()
    # Only treat as matchable if it looks like an email
    if "@" not in e:
        return ""
    return e

def normalize_phone(raw: object) -> Optional[str]:
    """
    Strict phone normalization:
    - Extract digits
    - Accept 10 digits, or 11 digits starting with '1' (US country code)
    - Otherwise return None
    """
    if raw is None:
        return None

    s = clean_cell(raw)
    if not s:
        return None

    digits = "".join(PHONE_RE.findall(s))

    # Strip leading country code 1 if present
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

    # Require 10 digits
    if len(digits) != 10:
        return None

    return digits

def _col_like(df: pd.DataFrame, *cands: str) -> Optional[str]:
    cols = list(df.columns)
    lowered = {str(c).strip().lower(): c for c in cols}
    for cand in cands:
        key = cand.strip().lower()
        if key in lowered:
            return lowered[key]
    for c in cols:
        cl = str(c).strip().lower()
        for cand in cands:
            if cand.strip().lower() in cl:
                return c
    return None

def _norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())

def _phone_hamming_1(a: str, b: str) -> bool:
    # One-digit difference on same-length digit strings
    if not a or not b or len(a) != len(b):
        return False
    if not (a.isdigit() and b.isdigit()):
        return False
    diffs = sum(1 for x, y in zip(a, b) if x != y)
    return diffs == 1

def find_similar_people(
    people: Dict[str, Dict[str, str]],
    first: str,
    last: str,
    email: str,
    phone: str,
) -> List[Dict[str, str]]:
    """
    Return a list of candidate people rows from people.csv that look like
    the same person as the Excel entry.
    """
    first_n = _norm_name(first)
    last_n = _norm_name(last)
    email_n = normalize_email(email)
    phone_n = (phone or "").strip()

    candidates = []

    # 1) Exact email match (best signal)
    if email_n:
        for p in people.values():
            db_e = normalize_email(p.get("email", ""))
            if db_e and db_e == email_n:
                candidates.append(p)

    # 2) Exact name match
    if first_n and last_n:
        for p in people.values():
            if _norm_name(p.get("first_name", "")) == first_n and _norm_name(p.get("last_name", "")) == last_n:
                candidates.append(p)

    # 3) Near-phone match for 10-digit numbers
    if phone_n and phone_n.isdigit():
        for p in people.values():
            db_phone = (p.get("phone") or "").strip()
            if not db_phone.isdigit():
                continue
            if len(phone_n) == 10 and len(db_phone) == 10:
                if phone_n[:9] == db_phone[:9] or _phone_hamming_1(phone_n, db_phone):
                    candidates.append(p)

    # De-dupe by phone, keep order
    seen = set()
    uniq = []
    for p in candidates:
        ph = p.get("phone")
        if ph and ph not in seen:
            seen.add(ph)
            uniq.append(p)
    return uniq

def load_people_csv(people_csv: str) -> Dict[str, Dict[str, str]]:
    people: Dict[str, Dict[str, str]] = {}
    if not os.path.exists(people_csv):
        return people

    with open(people_csv, "r", newline="", encoding="utf-8-sig") as f:
        r = csv.DictReader(f)
        for row in r:
            phone = normalize_phone(row.get("phone", ""))
            if not phone:
                continue
            gender = (row.get("gender") or "U").strip().upper() or "U"
            if gender not in ("M", "F", "U"):
                gender = "U"
            people[phone] = {
                "phone": phone,
                "first_name": (row.get("first_name") or "").strip(),
                "last_name": (row.get("last_name") or "").strip(),
                "email": (row.get("email") or "").strip(),
                "gender": gender,
            }
    return people


def save_people_csv(people_csv: str, people: Dict[str, Dict[str, str]]) -> None:
    os.makedirs(os.path.dirname(people_csv) or ".", exist_ok=True)
    fieldnames = ["phone", "first_name", "last_name", "email", "gender"]
    with open(people_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for phone in sorted(people.keys()):
            w.writerow(people[phone])


def _prompt(msg: str, default: Optional[str] = None) -> str:
    if default is None:
        return input(msg).strip()
    resp = input(f"{msg} [{default}]: ").strip()
    return resp if resp else default


def _yes_no(msg: str, default: bool = False) -> bool:
    d = "y" if default else "n"
    resp = input(f"{msg} (y/n) [{d}]: ").strip().lower()
    if not resp:
        return default
    return resp in ("y", "yes", "1", "true", "t")

def _backup_excel(excel_path: str) -> str:
    """
    Create a timestamped backup of the Excel file in a sibling 'backup' directory.
    """
    base_dir = os.path.dirname(os.path.abspath(excel_path))
    backup_dir = os.path.join(base_dir, "backup")
    os.makedirs(backup_dir, exist_ok=True)

    ts = time.strftime("%Y%m%d_%H%M%S")
    base_name = os.path.basename(excel_path)
    bak_name = f"{base_name}.bak_{ts}"
    bak_path = os.path.join(backup_dir, bak_name)

    shutil.copy2(excel_path, bak_path)
    return bak_path

def _write_sheet_back(excel_path: str, sheet_name: str, df: pd.DataFrame) -> None:
    # Replace just the target sheet; leave other sheets intact
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df.to_excel(w, sheet_name=sheet_name, index=False)

def _is_missing_db_value(x: str) -> bool:
    """
    Treat empty / 'nan' / 'none' as missing for DB fields.
    """
    s = clean_cell(x)
    return s == ""


def validate_and_fix_sheet(
    excel_path: str,
    sheet_name: str,
    people_csv: str,
) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
    """
    Returns (updated_df, updated_people_dict).
    Writes changes back to Excel (after backup) and optionally updates people.csv.
    """
    people = load_people_csv(people_csv)
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Column mapping
    c_first = _col_like(df, "First Name", "first")
    c_last = _col_like(df, "Last Name", "last")
    c_email = _col_like(df, "Email Address", "email")  # may be None
    c_phone = _col_like(df, "Mobile Number", "phone", "mobile")

    if not c_phone:
        raise ValueError("Could not find a phone/mobile column in this sheet.")

    # Force phone column to be text so edits don't fail due to int64 dtype
    df[c_phone] = df[c_phone].astype("string")

    # Add a working normalized phone column (not written unless you want it)
    norm_phones = []
    for _, row in df.iterrows():
        norm_phones.append(normalize_phone(row.get(c_phone)))
    df["_norm_phone"] = pd.Series(norm_phones, dtype="object")

    # Update missing records with new data from xlsx
    people_changed = False

    for idx in range(len(df)):
        ph = df.at[idx, "_norm_phone"]
        if not isinstance(ph, str) or not ph:
            continue
        if ph not in people:
            continue

        db = people[ph]

        xl_first = clean_cell(df.at[idx, c_first]) if c_first else ""
        xl_last = clean_cell(df.at[idx, c_last]) if c_last else ""
        xl_email = clean_cell(df.at[idx, c_email]) if c_email else ""

        # Only update DB if DB is missing and Excel has a non-empty value
        if xl_first and _is_missing_db_value(db.get("first_name", "")):
            db["first_name"] = xl_first
            people_changed = True

        if xl_last and _is_missing_db_value(db.get("last_name", "")):
            db["last_name"] = xl_last
            people_changed = True

        # Email: must look like an email to be considered "updating info"
        if normalize_email(xl_email) and not normalize_email(db.get("email", "")):
            db["email"] = xl_email
            people_changed = True

    if people_changed:
        save_people_csv(people_csv, people)
        print(f"Updated people.csv from Excel for existing phones with missing fields: {people_csv}")

    # Find duplicate phones (excluding None)
    dup_mask = df["_norm_phone"].notna() & df["_norm_phone"].duplicated(keep="first")

    # Bad entry criteria
    def row_is_bad(i: int) -> bool:
        phone = df.at[i, "_norm_phone"]
        if phone is None or pd.isna(phone):
            return True
        # Missing name columns only if they exist
        if c_first and not clean_cell(df.at[i, c_first]):
            return True
        if c_last and not clean_cell(df.at[i, c_last]):
            return True
        if bool(dup_mask.iloc[i]):
            return True
        return False

    bad_indices = [i for i in range(len(df)) if row_is_bad(i)]

    if bad_indices:
        print("\nFound entries that look invalid or incomplete:\n")
    else:
        print("\nNo obvious formatting issues found.\n")

    # Interactive fix loop
    for i in bad_indices:
        print("-" * 72)
        raw_phone = df.at[i, c_phone]
        phone = df.at[i, "_norm_phone"]

        first = clean_cell(df.at[i, c_first]) if c_first else ""
        last = clean_cell(df.at[i, c_last]) if c_last else ""
        email = clean_cell(df.at[i, c_email]) if c_email else ""

        reasons = []
        if phone is None:
            reasons.append("invalid/missing phone")
        if c_first and not first:
            reasons.append("missing first name")
        if c_last and not last:
            reasons.append("missing last name")
        if bool(dup_mask.iloc[i]):
            reasons.append("duplicate phone (later duplicate)")

        print(f"Row {i+2} issues: {', '.join(reasons)}")  # +2 accounts for header + 0-index
        print(f"  Raw phone: {raw_phone!r}")
        print(f"  Parsed phone: {phone!r}")
        if c_first:
            print(f"  First: {first!r}")
        if c_last:
            print(f"  Last:  {last!r}")
        if c_email:
            print(f"  Email: {email!r}")

        db_match = people.get(phone) if phone else None
        if db_match:
            print("  Match in people.csv:")
            print(f"    - {db_match['first_name']} {db_match['last_name']} | {db_match['email']} | {db_match['gender']}")

        print("\nActions:")
        print("  1) Search people.csv for similar people")
        print("  2) Edit this row manually")
        print("  3) Delete this row")
        print("  4) Leave as-is (skip)")

        choice = _prompt("Choose 1-4: ", "4")

        if choice == "1":
            # Unified search: phone (if parsed), email, name, near-phone (if possible)
            sims = find_similar_people(
                people,
                first=first,
                last=last,
                email=email,
                phone=phone or "",   # phone is normalized if valid, else None
            )

            # If phone is invalid, also try using the raw phone cell as potential email input
            if not sims:
                sims = find_similar_people(
                    people,
                    first=first,
                    last=last,
                    email=str(raw_phone),  # might be an email typed into the phone field
                    phone="",
                )

            if not sims:
                print("No likely matches found in people.csv by phone/name/email.")
                continue

            print("\nPossible matches in people.csv:")
            for idx, cand in enumerate(sims[:8], start=1):
                print(
                    f"  {idx}) {cand['phone']} | {cand.get('first_name','')} {cand.get('last_name','')} | "
                    f"{cand.get('email','')} | {cand.get('gender','U')}"
                )

            which = _prompt(f"Pick match number (1-{min(8, len(sims))}): ", "1").strip()
            try:
                k = int(which)
                cand = sims[k - 1]
            except Exception:
                print("Invalid selection. Skipping.")
                continue

            new_phone = cand["phone"]
            if not _yes_no(f"Apply this match and set Excel phone to {new_phone}?", default=True):
                continue

            # Apply: set phone, and fill missing fields from DB (don’t overwrite non-empty)
            df.at[i, c_phone] = str(new_phone)
            df.at[i, "_norm_phone"] = str(new_phone)

            if c_first and not first:
                df.at[i, c_first] = cand.get("first_name", "")
            if c_last and not last:
                df.at[i, c_last] = cand.get("last_name", "")
            if c_email and not email:
                df.at[i, c_email] = clean_cell(cand.get("email", ""))

            print("Row updated using selected people.csv match.")

        elif choice == "2":
            # Allow entering new phone
            new_phone_raw = _prompt("Phone (digits preferred)", str(raw_phone) if raw_phone is not None else "")
            new_phone = normalize_phone(new_phone_raw)
            if not new_phone:
                print("Invalid phone. Edit cancelled.")
                continue
            df.at[i, c_phone] = new_phone
            df.at[i, "_norm_phone"] = new_phone

            if c_first:
                df.at[i, c_first] = _prompt("First name", first)
            if c_last:
                df.at[i, c_last] = _prompt("Last name", last)
            if c_email:
                df.at[i, c_email] = _prompt("Email", email)

            print("Row edited.")

        elif choice == "3":
            df.at[i, "_delete"] = True
            print("Marked for deletion.")

        else:
            print("Skipped.")

    # Apply deletions
    if "_delete" in df.columns:
        before = len(df)
        df = df[df["_delete"] != True].copy()  # noqa: E712
        after = len(df)
        if after != before:
            print(f"\nDeleted {before - after} row(s).\n")

    # Recompute norm phone and duplicates after edits
    df["_norm_phone"] = pd.Series([normalize_phone(v) for v in df[c_phone].tolist()], dtype="object")
    dup_mask = df["_norm_phone"].notna() & df["_norm_phone"].duplicated(keep="first")
    if dup_mask.any():
        print("WARNING: There are still duplicate phones after cleanup. You may want to fix them.")
        # not forcing, but you could.

    # Offer to add missing people to people.csv
    phones_in_sheet = sorted({p for p in df["_norm_phone"].tolist() if isinstance(p, str) and p})
    missing = [p for p in phones_in_sheet if p not in people]
    if missing:
        print("\nPhones in this sheet missing from people.csv:")
        for p in missing[:50]:
            print(f"  - {p}")
        if len(missing) > 50:
            print(f"  ... and {len(missing) - 50} more")

        if _yes_no("\nReview and resolve missing entries now?", default=True):
            updated_sheet = False

            for p in missing:
                row = df[df["_norm_phone"] == p].iloc[0]

                first = clean_cell(row.get(c_first)) if c_first else ""
                last = clean_cell(row.get(c_last)) if c_last else ""
                email = clean_cell(row.get(c_email)) if c_email else ""

                print("\nMissing entry:")
                print(f"  Excel phone={p}")
                print(f"  Name={first} {last}".strip())
                if c_email:
                    print(f"  Email={email}")

                # Find similar existing people in DB
                sims = find_similar_people(people, first=first, last=last, email=email, phone=p)

                if sims:
                    print("\nPossible matches in people.csv:")
                    for idx, cand in enumerate(sims[:8], start=1):
                        print(
                            f"  {idx}) {cand['phone']} | {cand.get('first_name','')} {cand.get('last_name','')} | "
                            f"{cand.get('email','')} | {cand.get('gender','U')}"
                        )

                    print("\nActions:")
                    print("  1) Use a match and update the Excel phone to that DB phone (recommended)")
                    print("  2) Add as new person to people.csv anyway")
                    print("  3) Skip for now")

                    action = _prompt("Choose 1-3: ", "1").strip()

                    if action == "1":
                        which = _prompt(f"Pick match number (1-{min(8, len(sims))}): ", "1").strip()
                        try:
                            k = int(which)
                            cand = sims[k - 1]
                        except Exception:
                            print("Invalid selection. Skipping.")
                            continue

                        new_phone = cand["phone"]
                        if _yes_no(f"Replace Excel phone {p} -> {new_phone} in the sheet?", default=True):
                            # Update ALL rows with the old phone
                            mask = df["_norm_phone"] == p
                            df.loc[mask, c_phone] = str(new_phone)
                            df.loc[mask, "_norm_phone"] = new_phone

                            # Optionally also fill missing fields from DB
                            if c_first and not first:
                                df.loc[mask, c_first] = cand.get("first_name", "")
                            if c_last and not last:
                                df.loc[mask, c_last] = cand.get("last_name", "")
                            if c_email and not email:
                                df.loc[mask, c_email] = cand.get("email", "")

                            updated_sheet = True
                            print("Sheet updated to match people.csv.")
                        continue

                    if action == "3":
                        continue
                    # else fall through to add-new

                # Add-new flow (no matches or user chose add)
                print("\nAdd as new person to people.csv:")
                print(f"  phone={p}")
                if first or last or email:
                    print(f"  from sheet: {first} {last} | {email}")

                if not _yes_no("  Add this person?", default=True):
                    continue

                people[p] = {
                    "phone": p,
                    "first_name": first,
                    "last_name": last,
                    "email": email,
                    "gender": "U",
                }

            # Save DB if changed
            save_people_csv(people_csv, people)
            print(f"\nUpdated people.csv: {people_csv}\n")

            # If we updated the sheet phones, recompute and keep going
            if updated_sheet:
                df["_norm_phone"] = [normalize_phone(v) for v in df[c_phone].tolist()]

    # Clean helper columns before writing back
    df_out = df.drop(columns=[c for c in ["_norm_phone", "_delete"] if c in df.columns], errors="ignore")

    # Backup + write
    bak = _backup_excel(excel_path)
    _write_sheet_back(excel_path, sheet_name, df_out)
    print(f"Excel updated. Backup created at: {bak}")

    return df_out, people
