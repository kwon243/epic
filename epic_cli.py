#!/usr/bin/env python3
"""
Author: Jinhwan Kwon
Last Modified: Feb 2026

epic_cli.py

Menu-driven CLI wrapper for make_groups.py and DB maintenance.

Assumptions:
- db/ is in the same directory as make_groups.py and epic_cli.py (no prompting).
- default Excel file for signups is responses.xlsx (same directory as scripts).

Top-level:
  1) Make groups (guided prompts, runs make_groups.py)
  2) Modify DB
     - People: search/add/edit/remove
     - Locked pairs: search/add/edit/remove locked pairs with score=255
"""

from __future__ import annotations

import csv
import os
import pandas as pd
import re
import subprocess
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from epic_excel_check import validate_and_fix_sheet


PHONE_RE = re.compile(r"\d+")

PEOPLE_FIELDS = ["phone", "first_name", "last_name", "email", "gender"]
PAIR_FIELDS = ["phone_a", "phone_b", "score", "locked", "note"]


# ----------------------------
# Utilities
# ----------------------------

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

def canon_pair(a: str, b: str) -> Tuple[str, str]:
    return (a, b) if a < b else (b, a)

def pick_newest_round_sheet(excel_path: str) -> str:
    """
    Return the newest 'Round N' sheet name from the Excel file.
    Raises ValueError if none found.
    """
    xl = pd.ExcelFile(excel_path)
    best = None
    best_n = -1

    for name in xl.sheet_names:
        m = re.match(r"round\s*(\d+)$", name, flags=re.IGNORECASE)
        if m:
            n = int(m.group(1))
            if n > best_n:
                best_n = n
                best = name

    if not best:
        raise ValueError("No sheets named like 'Round N' were found.")

    return best

def prompt(msg: str, default: Optional[str] = None) -> str:
    if default is None:
        return input(msg).strip()
    resp = input(f"{msg} [{default}]: ").strip()
    return resp if resp else default


def prompt_yes_no(msg: str, default: bool = False) -> bool:
    d = "y" if default else "n"
    resp = input(f"{msg} (y/n) [{d}]: ").strip().lower()
    if not resp:
        return default
    return resp in ("y", "yes", "1", "true", "t")


def print_header(title: str) -> None:
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72 + "\n")


def safe_int(s: str, default: int = 0) -> int:
    try:
        return int(str(s).strip())
    except Exception:
        return default


# ----------------------------
# CSV DB helpers
# ----------------------------

@dataclass
class PairMeta:
    score: int
    locked: bool
    note: str = ""


def ensure_db_files(db_dir: str) -> Tuple[str, str]:
    os.makedirs(db_dir, exist_ok=True)
    people_path = os.path.join(db_dir, "people.csv")
    pairs_path = os.path.join(db_dir, "pair_scores.csv")

    if not os.path.exists(people_path):
        with open(people_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=PEOPLE_FIELDS)
            w.writeheader()

    if not os.path.exists(pairs_path):
        with open(pairs_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=PAIR_FIELDS)
            w.writeheader()

    return people_path, pairs_path


def load_people(people_path: str) -> Dict[str, Dict[str, str]]:
    people: Dict[str, Dict[str, str]] = {}
    with open(people_path, "r", newline="", encoding="utf-8-sig") as f:
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


def save_people(people_path: str, people: Dict[str, Dict[str, str]]) -> None:
    with open(people_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=PEOPLE_FIELDS)
        w.writeheader()
        for phone in sorted(people.keys()):
            w.writerow(people[phone])


def load_pairs(pairs_path: str) -> Dict[Tuple[str, str], PairMeta]:
    pairs: Dict[Tuple[str, str], PairMeta] = {}
    with open(pairs_path, "r", newline="", encoding="utf-8-sig") as f:
        r = csv.DictReader(f)
        for row in r:
            a = normalize_phone(row.get("phone_a", ""))
            b = normalize_phone(row.get("phone_b", ""))
            if not a or not b or a == b:
                continue
            ca, cb = canon_pair(a, b)
            score = safe_int(row.get("score", 0), 0)
            locked = str(row.get("locked", "")).strip().lower() in ("1", "true", "t", "yes", "y")
            note = (row.get("note") or "").strip()
            pairs[(ca, cb)] = PairMeta(score=score, locked=locked, note=note)
    return pairs


def save_pairs(pairs_path: str, pairs: Dict[Tuple[str, str], PairMeta]) -> None:
    with open(pairs_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=PAIR_FIELDS)
        w.writeheader()
        for (a, b), meta in sorted(pairs.items(), key=lambda kv: (kv[0][0], kv[0][1])):
            w.writerow(
                {
                    "phone_a": a,
                    "phone_b": b,
                    "score": int(meta.score),
                    "locked": bool(meta.locked),
                    "note": meta.note or "",
                }
            )


# ----------------------------
# Menu: Make groups (wrapper)
# ----------------------------

def make_groups_flow(project_dir: str) -> None:
    print_header("Make groups")

    default_excel = os.path.join(project_dir, "responses.xlsx")
    excel = prompt("Excel path (weekly signups)", default_excel).strip().strip('"').strip("'")
    if not os.path.isfile(excel):
        print(f"File not found: {excel}")
        input("Press Enter...")
        return

    # Round number -> sheet name
    round_num = prompt("Round number (blank = newest)", default="").strip()

    if round_num:
        if not round_num.isdigit():
            print("Round number must be numeric (e.g., 4).")
            input("Press Enter...")
            return
        sheet = f"Round {int(round_num)}"
    else:
        try:
            sheet = pick_newest_round_sheet(excel)
            print(f"Using newest sheet: {sheet}")
        except Exception as e:
            print(f"Could not determine newest round: {e}")
            input("Press Enter...")
            return

    # db dir assumed
    db_dir = os.path.join(project_dir, "db")
    people_path, _ = ensure_db_files(db_dir)

    # Validate + fix responses before running grouping
    try:
        validate_and_fix_sheet(excel_path=excel, sheet_name=sheet, people_csv=people_path)
    except Exception as e:
        print(f"Response check failed: {e}")
        input("Press Enter...")
        return

    # Ask gender-aware matching after checks
    gender_mindful = prompt_yes_no("Enable gender-bias?", default=False)

    out = prompt("Output file path (blank = default output/)", default="").strip()

    cmd = [sys.executable, os.path.join(project_dir, "make_groups.py"), "--excel", excel, "--db-dir", db_dir, "--sheet", sheet]
    if gender_mindful:
        cmd += ["--gender-mindful"]
    if out:
        cmd += ["--out", out]

    print("\nRunning:\n  " + " ".join(f'"{c}"' if " " in c else c for c in cmd) + "\n")

    try:
        subprocess.run(cmd, check=False)
    except Exception as e:
        print(f"Failed to run make_groups.py: {e}")
        input("Press Enter...")


# ----------------------------
# Menu: Modify DB
# ----------------------------

def people_menu(people_path: str) -> None:
    while True:
        people = load_people(people_path)

        print_header("Modify DB -> People")
        print(f"People loaded: {len(people)}")
        print("1) Search people")
        print("2) Add person")
        print("3) Edit person")
        print("4) Remove person")
        print("5) Back")
        choice = prompt("Choose: ").strip()

        if choice == "1":
            q = prompt("Search (phone fragment or name/email): ").strip().lower()
            results = []
            for p in people.values():
                hay = " ".join([p["phone"], p["first_name"], p["last_name"], p["email"], p["gender"]]).lower()
                if q in hay:
                    results.append(p)
            if not results:
                print("No matches.")
            else:
                for p in results[:50]:
                    print(f'- {p["phone"]} | {p["first_name"]} {p["last_name"]} | {p["email"]} | {p["gender"]}')
                if len(results) > 50:
                    print(f"... and {len(results) - 50} more.")
            input("\nPress Enter...")

        elif choice == "2":
            raw_phone = prompt("Phone: ").strip()
            phone = normalize_phone(raw_phone)
            if not phone:
                print("Invalid phone.")
                input("Press Enter...")
                continue
            if phone in people:
                print("That phone already exists in people.csv.")
                input("Press Enter...")
                continue

            first = prompt("First name: ").strip()
            last = prompt("Last name: ").strip()
            email = prompt("Email: ").strip()
            gender = prompt("Gender (M/F/U)", default="U").strip().upper()
            if gender not in ("M", "F", "U"):
                gender = "U"

            people[phone] = {
                "phone": phone,
                "first_name": first,
                "last_name": last,
                "email": email,
                "gender": gender,
            }
            save_people(people_path, people)
            print("Added.")
            input("Press Enter...")

        elif choice == "3":
            raw_phone = prompt("Phone of person to edit: ").strip()
            phone = normalize_phone(raw_phone)
            if not phone or phone not in people:
                print("Not found.")
                input("Press Enter...")
                continue

            p = people[phone]
            print(f'Editing {p["phone"]} ({p["first_name"]} {p["last_name"]})')

            p["first_name"] = prompt("First name", default=p["first_name"]).strip()
            p["last_name"] = prompt("Last name", default=p["last_name"]).strip()
            p["email"] = prompt("Email", default=p["email"]).strip()
            g = prompt("Gender (M/F/U)", default=p["gender"]).strip().upper()
            p["gender"] = g if g in ("M", "F", "U") else "U"

            people[phone] = p
            save_people(people_path, people)
            print("Updated.")
            input("Press Enter...")

        elif choice == "4":
            raw_phone = prompt("Phone of person to remove: ").strip()
            phone = normalize_phone(raw_phone)
            if not phone or phone not in people:
                print("Not found.")
                input("Press Enter...")
                continue

            p = people[phone]
            ok = prompt_yes_no(f'Remove {p["first_name"]} {p["last_name"]} ({p["phone"]})?', default=False)
            if ok:
                people.pop(phone, None)
                save_people(people_path, people)
                print("Removed.")
            else:
                print("Cancelled.")
            input("Press Enter...")

        elif choice == "5":
            return
        else:
            print("Invalid choice.")
            input("Press Enter...")


def locked_pairs_menu(pairs_path: str) -> None:
    """
    Only manages locked pairs (forbidden pairs).
    Convention: locked=True and usually score=255.
    """
    while True:
        pairs = load_pairs(pairs_path)
        locked = {k: v for k, v in pairs.items() if v.locked}

        print_header("Modify DB -> Locked pairs (forbidden)")
        print(f"Locked pairs: {len(locked)}")
        print("1) View/Search pairs")
        print("2) Add pair")
        print("3) Edit locked pair note")
        print("4) Remove pair")
        print("5) Back")
        choice = prompt("Choose: ").strip()

        if choice == "1":
            q = prompt("Search by phone or leave empty to view all: ").strip()
            qn = normalize_phone(q) or q
            results = []
            for (a, b), meta in locked.items():
                if qn in a or qn in b:
                    results.append(((a, b), meta))
            if not results:
                print("No matches.")
            else:
                for (a, b), meta in results[:80]:
                    print(f"- {a} <-> {b} | score={meta.score} | locked={meta.locked} | note={meta.note}")
                if len(results) > 80:
                    print(f"... and {len(results) - 80} more.")
            input("\nPress Enter...")

        elif choice == "2":
            a = normalize_phone(prompt("Phone A: ").strip())
            b = normalize_phone(prompt("Phone B: ").strip())
            if not a or not b or a == b:
                print("Invalid phones.")
                input("Press Enter...")
                continue

            ca, cb = canon_pair(a, b)
            note = prompt("Note (optional): ", default="").strip()

            pairs[(ca, cb)] = PairMeta(score=255, locked=True, note=note)
            save_pairs(pairs_path, pairs)
            print(f"Added locked pair {ca} <-> {cb} (score=255).")
            input("Press Enter...")

        elif choice == "3":
            a = normalize_phone(prompt("Phone A: ").strip())
            b = normalize_phone(prompt("Phone B: ").strip())
            if not a or not b or a == b:
                print("Invalid phones.")
                input("Press Enter...")
                continue
            ca, cb = canon_pair(a, b)
            meta = pairs.get((ca, cb))
            if not meta or not meta.locked:
                print("That pair is not currently locked.")
                input("Press Enter...")
                continue

            meta.note = prompt("New note", default=meta.note).strip()
            meta.locked = True
            meta.score = 255
            pairs[(ca, cb)] = meta

            save_pairs(pairs_path, pairs)
            print("Updated note.")
            input("Press Enter...")

        elif choice == "4":
            a = normalize_phone(prompt("Phone A: ").strip())
            b = normalize_phone(prompt("Phone B: ").strip())
            if not a or not b or a == b:
                print("Invalid phones.")
                input("Press Enter...")
                continue
            ca, cb = canon_pair(a, b)
            meta = pairs.get((ca, cb))
            if not meta or not meta.locked:
                print("That pair is not currently locked.")
                input("Press Enter...")
                continue

            ok = prompt_yes_no(f"Remove locked pair {ca} <-> {cb}?", default=False)
            if ok:
                pairs.pop((ca, cb), None)
                save_pairs(pairs_path, pairs)
                print("Removed.")
            else:
                print("Cancelled.")
            input("Press Enter...")

        elif choice == "5":
            return
        else:
            print("Invalid choice.")
            input("Press Enter...")


def modify_db_flow(project_dir: str) -> None:
    # db dir is assumed (no prompt)
    db_dir = os.path.join(project_dir, "db")
    people_path, pairs_path = ensure_db_files(db_dir)

    while True:
        print_header("Modify DB")
        print(f"DB dir: {db_dir}")
        print("1) People (people.csv)")
        print("2) Locked pairs (pair_scores.csv, score=255)")
        print("3) Back")
        choice = prompt("Choose: ").strip()

        if choice == "1":
            people_menu(people_path)
        elif choice == "2":
            locked_pairs_menu(pairs_path)
        elif choice == "3":
            return
        else:
            print("Invalid choice.")
            input("Press Enter...")


# ----------------------------
# Top-level menu
# ----------------------------

def resolve_project_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def main() -> int:
    project_dir = resolve_project_dir()

    mg = os.path.join(project_dir, "make_groups.py")
    if not os.path.exists(mg):
        print(f"ERROR: make_groups.py not found next to epic_cli.py: {mg}")
        return 2

    while True:
        print_header("Epic Movement Grouping CLI")
        print("1) Make groups")
        print("2) Modify DB")
        print("3) Quit")
        choice = prompt("Choose: ").strip()

        if choice == "1":
            make_groups_flow(project_dir)
        elif choice == "2":
            modify_db_flow(project_dir)
        elif choice == "3":
            print("Bye.")
            return 0
        else:
            print("Invalid choice.")
            input("Press Enter...")


if __name__ == "__main__":
    raise SystemExit(main())
