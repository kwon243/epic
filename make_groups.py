#!/usr/bin/env python3
"""
Author: Jinhwan Kwon
Last Modified: Feb 2026

Grouping script (2-3 people) with pairwise recency scores.

- Primary key is phone number (string). Never infer from name/email. (because David Kim)
- Pair scores stored canonically: phone_a < phone_b (string compare).
- Locked pairs are forbidden; any group containing one is invalid.
- Missing pair rows are treated as score=0, locked=false.

Excel is used for weekly signups and group-size preferences.
CSV "database" (people.csv, pair_scores.csv) provides canonical person records and scores.

Usage:
python .\make_groups.py --excel .\responses.xlsx
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import random
import re
import sys
from dataclasses import dataclass
from itertools import combinations
from typing import Dict, List, Optional, Tuple


try:
    import pandas as pd
except ImportError:
    print("Missing dependency: pandas. Install with: pip install pandas openpyxl", file=sys.stderr)
    raise


# -------------------------
# Data models
# -------------------------

@dataclass(frozen=True)
class Person:
    phone: str
    first_name: str
    last_name: str
    email: str
    gender: str  # "M", "F", "U"


@dataclass
class PairMeta:
    score: int
    locked: bool
    note: str = ""


# -------------------------
# Helpers
# -------------------------

PHONE_RE = re.compile(r"\d+")

def normalize_phone(raw: str) -> Optional[str]:
    """
    Normalize phone to digits-only string.
    Returns None if it doesn't look like a phone number (too short).
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    digits = "".join(PHONE_RE.findall(s))
    # Basic sanity: US numbers often 10 digits, but allow 7+ to be flexible.
    if len(digits) < 7:
        return None
    return digits


def canon_pair(a: str, b: str) -> Tuple[str, str]:
    return (a, b) if a < b else (b, a)


def parse_preference(pref: str) -> str:
    """
    Returns one of: "2", "3", "N" (no preference).
    """
    if pref is None:
        return "N"
    s = str(pref).strip().lower()
    if "2" in s:
        return "2"
    if "3" in s:
        return "3"
    return "N"


def safe_bool(x) -> bool:
    if isinstance(x, bool):
        return x
    s = str(x).strip().lower()
    return s in ("1", "true", "t", "yes", "y")


def gender_of(phone: str, people: Dict[str, Person]) -> str:
    p = people.get(phone)
    return (p.gender if p else "U") or "U"


def display_name(phone: str, people: Dict[str, Person], fallback: Dict[str, Dict[str, str]]) -> str:
    if phone in people:
        p = people[phone]
        return f"{p.first_name.strip()} {p.last_name.strip()}".strip()
    # fallback from excel
    fb = fallback.get(phone, {})
    fn = (fb.get("first_name") or "").strip()
    ln = (fb.get("last_name") or "").strip()
    name = f"{fn} {ln}".strip()
    return name if name else phone


def display_phone(phone: str) -> str:
    # format US 10-digit if possible
    d = phone
    if len(d) == 10:
        return f"({d[0:3]}) {d[3:6]}-{d[6:10]}"
    return phone


# -------------------------
# Load DB CSVs
# -------------------------

def load_people_csv(path: str) -> Dict[str, Person]:
    people: Dict[str, Person] = {}
    if not os.path.exists(path):
        raise FileNotFoundError(f"people.csv not found: {path}")

    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        required = {"phone", "first_name", "last_name", "email", "gender"}
        missing = required - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"people.csv missing columns: {sorted(missing)}")

        for row in reader:
            phone = normalize_phone(row.get("phone", ""))
            if not phone:
                continue
            gender = (row.get("gender") or "U").strip().upper()
            if gender not in ("M", "F", "U"):
                gender = "U"
            people[phone] = Person(
                phone=phone,
                first_name=(row.get("first_name") or "").strip(),
                last_name=(row.get("last_name") or "").strip(),
                email=(row.get("email") or "").strip(),
                gender=gender,
            )
    return people


def load_pair_scores_csv(path: str) -> Dict[Tuple[str, str], PairMeta]:
    pairs: Dict[Tuple[str, str], PairMeta] = {}
    if not os.path.exists(path):
        # allow starting fresh
        return pairs

    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        required = {"phone_a", "phone_b", "score", "locked"}
        missing = required - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"pair_scores.csv missing columns: {sorted(missing)}")

        for row in reader:
            a = normalize_phone(row.get("phone_a", ""))
            b = normalize_phone(row.get("phone_b", ""))
            if not a or not b or a == b:
                continue
            ca, cb = canon_pair(a, b)

            try:
                score = int(row.get("score", 0))
            except Exception:
                score = 0
            locked = safe_bool(row.get("locked", False))
            note = (row.get("note") or "").strip()

            pairs[(ca, cb)] = PairMeta(score=score, locked=locked, note=note)
    return pairs


def save_pair_scores_csv(path: str, pairs: Dict[Tuple[str, str], PairMeta]) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)

    # Keep stable order for diff readability
    items = sorted(pairs.items(), key=lambda kv: (kv[0][0], kv[0][1]))

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["phone_a", "phone_b", "score", "locked", "note"])
        writer.writeheader()
        for (a, b), meta in items:
            writer.writerow(
                {
                    "phone_a": a,
                    "phone_b": b,
                    "score": int(meta.score),
                    "locked": bool(meta.locked),
                    "note": meta.note or "",
                }
            )


# -------------------------
# Load weekly signups from Excel
# -------------------------

def pick_default_sheet(xlsx_path: str) -> str:
    xl = pd.ExcelFile(xlsx_path)
    # Prefer "Round N" with largest N if present, else first sheet.
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


def load_weekly_signups(
    xlsx_path: str,
    sheet: Optional[str],
) -> Tuple[List[str], Dict[str, str], Dict[str, Dict[str, str]]]:
    """
    Returns:
      phones: list of participant phones
      preference_by_phone: phone -> "2"|"3"|"N"
      fallback_by_phone: phone -> {"first_name","last_name","email"} (from Excel)
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Excel not found: {xlsx_path}")

    if sheet is None:
        sheet = pick_default_sheet(xlsx_path)

    df = pd.read_excel(xlsx_path, sheet_name=sheet)

    # Flexible column lookup (Google Forms changes sometimes)
    def col_like(*cands: str) -> Optional[str]:
        cols = list(df.columns)
        lowered = {str(c).strip().lower(): c for c in cols}
        for cand in cands:
            c = lowered.get(cand.strip().lower())
            if c is not None:
                return c
        # fallback fuzzy
        for c in cols:
            cl = str(c).strip().lower()
            for cand in cands:
                if cand.strip().lower() in cl:
                    return c
        return None

    c_first = col_like("First Name", "first")
    c_last = col_like("Last Name", "last")
    c_email = col_like("Email Address", "email")
    c_phone = col_like("Mobile Number", "phone", "mobile")
    c_pref = col_like("Would you prefer to be in a group of...", "prefer", "group of")

    if not c_phone:
        raise ValueError("Could not find a phone/mobile column in the Excel sheet.")

    phones: List[str] = []
    pref: Dict[str, str] = {}
    fallback: Dict[str, Dict[str, str]] = {}

    for _, row in df.iterrows():
        raw_phone = row.get(c_phone)
        phone = normalize_phone(raw_phone)
        if not phone:
            # Skip rows with non-phone entries
            continue

        phones.append(phone)
        pref_val = parse_preference(row.get(c_pref) if c_pref else None)
        pref[phone] = pref_val

        fallback[phone] = {
            "first_name": str(row.get(c_first) or "").strip() if c_first else "",
            "last_name": str(row.get(c_last) or "").strip() if c_last else "",
            "email": str(row.get(c_email) or "").strip() if c_email else "",
        }

    # De-dupe while preserving order
    seen = set()
    uniq = []
    for p in phones:
        if p not in seen:
            seen.add(p)
            uniq.append(p)

    return uniq, pref, fallback


# -------------------------
# Scoring + constraints
# -------------------------

def get_pair_meta(pairs: Dict[Tuple[str, str], PairMeta], a: str, b: str) -> PairMeta:
    ca, cb = canon_pair(a, b)
    return pairs.get((ca, cb), PairMeta(score=0, locked=False, note=""))


def group_pairs(group: List[str]) -> List[Tuple[str, str]]:
    return [canon_pair(a, b) for a, b in combinations(group, 2)]


def group_is_valid(
    group: List[str],
    pairs: Dict[Tuple[str, str], PairMeta],
    people: Dict[str, Person],
    gender_mindful: bool,
) -> bool:
    # Locked pair check
    for a, b in combinations(group, 2):
        meta = get_pair_meta(pairs, a, b)
        if meta.locked:
            return False

    if not gender_mindful:
        return True

    genders = [gender_of(p, people) for p in group]

    # Rule 1: avoid man-woman pairs (M+F) when both known
    if len(group) == 2:
        gset = set(genders)
        if "M" in gset and "F" in gset and "U" not in gset:
            return False

    # Rule 2: avoid woman-man-man triples (F+M+M) when all known
    if len(group) == 3:
        if "U" not in genders:
            if genders.count("F") == 1 and genders.count("M") == 2:
                return False

    return True


def group_cost(group: List[str], pairs: Dict[Tuple[str, str], PairMeta]) -> int:
    cost = 0
    for a, b in combinations(group, 2):
        meta = get_pair_meta(pairs, a, b)
        cost += int(meta.score)
    return cost


def preference_penalty(group: List[str], pref: Dict[str, str]) -> int:
    # Soft penalty: 0 if satisfied or no preference, 2 if violated.
    # You can tune these if you want preferences stricter/looser.
    size = len(group)
    pen = 0
    for p in group:
        want = pref.get(p, "N")
        if want == "N":
            continue
        if want == "2" and size != 2:
            pen += 2
        if want == "3" and size != 3:
            pen += 2
    return pen


def total_cost(groups: List[List[str]], pairs: Dict[Tuple[str, str], PairMeta], pref: Dict[str, str]) -> int:
    return sum(group_cost(g, pairs) + preference_penalty(g, pref) for g in groups)


# -------------------------
# Size plan selection
# -------------------------

def all_size_plans(n: int) -> List[List[int]]:
    """
    Return all plans (list of sizes) such that 2*a + 3*b = n.
    """
    plans = []
    for b in range(n // 3 + 1):
        rem = n - 3 * b
        if rem < 0:
            continue
        if rem % 2 == 0:
            a = rem // 2
            plans.append([3] * b + [2] * a)
    return plans


def plan_penalty(plan: List[int], phones: List[str], pref: Dict[str, str]) -> int:
    """
    Estimate mismatch penalty for a plan ignoring actual pair feasibility.
    Greedy assign people who prefer 3 to 3-slots etc.
    """
    want2 = [p for p in phones if pref.get(p, "N") == "2"]
    want3 = [p for p in phones if pref.get(p, "N") == "3"]
    no = [p for p in phones if pref.get(p, "N") == "N"]

    slots2 = plan.count(2) * 2
    slots3 = plan.count(3) * 3

    # Put want3 into 3 slots first, want2 into 2 slots first.
    # Penalty counts the overflow that will likely be mismatched.
    pen = 0
    if len(want3) > slots3:
        pen += (len(want3) - slots3) * 2 # 2 -> mismatch penalty
    if len(want2) > slots2:
        pen += (len(want2) - slots2) * 2

    # Also penalize if plan has too few 3 slots given want3
    # and too few 2 slots given want2 (already captured).
    return pen


def count_no_pref(phones: List[str], pref: Dict[str, str]) -> int:
    return sum(1 for p in phones if pref.get(p, "N") == "N")


def choose_best_plans(n: int, phones: List[str], pref: Dict[str, str], top_k: int = 3) -> List[List[int]]:
    """
    Choose size plans (2/3) for n people.

    Primary objective: satisfy explicit preferences ("2"/"3") as much as possible.
    Secondary objective: among similar preference-mismatch plans, prefer MORE 3-person groups
    so that no-preference people are more likely to land in 3s than 2s.
    """
    plans = all_size_plans(n)

    want3_count = sum(1 for p in phones if pref.get(p, "N") == "3")
    want2_count = sum(1 for p in phones if pref.get(p, "N") == "2")
    no_pref_count = count_no_pref(phones, pref)

    def sort_key(pl: List[int]):
        ppen = plan_penalty(pl, phones, pref)

        num3 = pl.count(3)
        num2 = pl.count(2)

        # If there are many no-preference people, it's especially reasonable to bias toward 3s.
        # We'll treat "more 3s" as a tie-breaker, not a primary objective.
        # (Lower key is better.)
        return (
            ppen,                        # 1) minimize preference mismatch
            -num3,                       # 2) prefer more groups of 3
            num2,                        # 3) prefer fewer groups of 2
            abs(num3 - (want3_count // 3)),  # 4) weak alignment with amount of want3
        )

    plans.sort(key=sort_key)
    return plans[: max(1, min(top_k, len(plans)))]



# -------------------------
# Group construction heuristic
# -------------------------

def build_groups_once(
    phones: List[str],
    pref: Dict[str, str],
    people: Dict[str, Person],
    pairs: Dict[Tuple[str, str], PairMeta],
    size_plan: List[int],
    gender_mindful: bool,
    rng: random.Random,
    topk: int = 8,
) -> Optional[List[List[str]]]:
    remaining = phones[:]
    rng.shuffle(remaining)

    groups: List[List[str]] = []
    sizes = size_plan[:]
    rng.shuffle(sizes)

    # Prefer placing "hard preference" people first (those with 2 or 3 pref)
    def pref_rank(p: str) -> int:
        w = pref.get(p, "N")
        return 0 if w == "N" else 1

    remaining.sort(key=pref_rank, reverse=True)

    for size in sizes:
        if len(remaining) < size:
            return None

        # pick a seed: someone with preferences first, otherwise random
        seed_idx = 0
        seed = remaining.pop(seed_idx)

        if size == 2:
            candidates = []
            for j, other in enumerate(remaining):
                g = [seed, other]
                if not group_is_valid(g, pairs, people, gender_mindful):
                    continue
                c = group_cost(g, pairs) + preference_penalty(g, pref)
                candidates.append((c, j, other))

            if not candidates:
                return None

            candidates.sort(key=lambda x: x[0])
            pick = rng.choice(candidates[: min(topk, len(candidates))])
            _, j, other = pick
            remaining.pop(j)
            groups.append([seed, other])

        elif size == 3:
            candidates = []
            # Try pairs among remaining
            for (j, a), (k, b) in combinations(list(enumerate(remaining)), 2):
                g = [seed, a, b]
                if not group_is_valid(g, pairs, people, gender_mindful):
                    continue
                c = group_cost(g, pairs) + preference_penalty(g, pref)
                candidates.append((c, j, k, a, b))

            if not candidates:
                return None

            candidates.sort(key=lambda x: x[0])
            pick = rng.choice(candidates[: min(topk, len(candidates))])
            _, j, k, a, b = pick
            # remove higher index first
            for idx in sorted([j, k], reverse=True):
                remaining.pop(idx)
            groups.append([seed, a, b])
        else:
            raise ValueError("Size plan can only contain 2 or 3.")

    # Final sanity
    used = [p for g in groups for p in g]
    if sorted(used) != sorted(phones):
        return None

    return groups


def find_best_groups(
    phones: List[str],
    pref: Dict[str, str],
    people: Dict[str, Person],
    pairs: Dict[Tuple[str, str], PairMeta],
    gender_mindful: bool,
    restarts: int,
    seed: Optional[int],
) -> List[List[str]]:
    rng = random.Random(seed)

    best: Optional[List[List[str]]] = None
    best_cost = 10**18

    plans = choose_best_plans(len(phones), phones, pref, top_k=3)

    for plan in plans:
        for _ in range(restarts):
            attempt = build_groups_once(
                phones=phones,
                pref=pref,
                people=people,
                pairs=pairs,
                size_plan=plan,
                gender_mindful=gender_mindful,
                rng=rng,
            )
            if attempt is None:
                continue

            c = total_cost(attempt, pairs, pref)
            if c < best_cost:
                best_cost = c
                best = attempt

                # If perfect (0 cost), we can stop early
                if best_cost == 0:
                    return best

    if best is None:
        raise RuntimeError(
            "Could not find a valid grouping. Common causes:\n"
            "- Too many locked pairs among the current participants\n"
            "- Gender-mindful constraints make grouping impossible\n"
            "- Odd participant count with heavy size preferences\n"
        )
    return best


# -------------------------
# Pair score updates
# -------------------------

def decay_scores(pairs: Dict[Tuple[str, str], PairMeta]) -> None:
    for meta in pairs.values():
        if meta.locked:
            continue
        meta.score = max(0, int(meta.score) - 1)


def set_group_scores_to_3(
    pairs: Dict[Tuple[str, str], PairMeta],
    groups: List[List[str]],
) -> None:
    for g in groups:
        for a, b in combinations(g, 2):
            ca, cb = canon_pair(a, b)
            meta = pairs.get((ca, cb))
            if meta is None:
                pairs[(ca, cb)] = PairMeta(score=3, locked=False, note="")
            else:
                if not meta.locked:
                    meta.score = 3


# -------------------------
# Output formatting
# -------------------------

def print_groups(groups: List[List[str]], people: Dict[str, Person], fallback: Dict[str, Dict[str, str]], pairs: Dict[Tuple[str, str], PairMeta], pref: Dict[str, str]) -> None:
    print("\nProposed groups:\n")
    for i, g in enumerate(groups, start=1):
        g_sorted = g[:]  # keep stable-ish
        c = group_cost(g_sorted, pairs)
        ppen = preference_penalty(g_sorted, pref)
        print(f"Group {i} (size {len(g_sorted)}): cost={c}, pref_penalty={ppen}")
        for phone in g_sorted:
            name = display_name(phone, people, fallback)
            gender = gender_of(phone, people)
            want = pref.get(phone, "N")
            want_s = {"2": "pref 2", "3": "pref 3", "N": "no pref"}.get(want, "no pref")
            print(f"  - {name} [{gender}, {want_s}]  {display_phone(phone)}")
        print()

    print(f"Total objective cost = {total_cost(groups, pairs, pref)}\n")


def write_formatted_output(
    out_path: str,
    groups: List[List[str]],
    people: Dict[str, Person],
    fallback: Dict[str, Dict[str, str]],
) -> None:
    lines: List[str] = []
    lines.append(f"Weekly Groups ({dt.date.today().isoformat()})")
    lines.append("")

    for i, g in enumerate(groups, start=1):
        lines.append(f"Group {i} (size {len(g)}):")
        for phone in g:
            name = display_name(phone, people, fallback)
            lines.append(f"  - {name} | {display_phone(phone)}")
        lines.append("")

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# -------------------------
# Main
# -------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description="Create weekly groups of 2–3 minimizing recency scores.")
    ap.add_argument("--excel", required=True, help="Path to the Excel file with weekly signups.")
    ap.add_argument("--sheet", default=None, help="Sheet name to use (default: newest Round N).")
    ap.add_argument("--db-dir", default="db", help="Directory containing people.csv and pair_scores.csv.")
    ap.add_argument("--gender-mindful", action="store_true", help="Enable gender-mindful matching constraints.")
    ap.add_argument("--restarts", type=int, default=500, help="Random restarts per size plan (higher = better but slower).")
    ap.add_argument("--seed", type=int, default=None, help="Random seed for reproducibility.")
    ap.add_argument("--out", default=None, help="Output file path (default: output/groups_YYYY-MM-DD.txt).")
    args = ap.parse_args()

    people_path = os.path.join(args.db_dir, "people.csv")
    pair_path = os.path.join(args.db_dir, "pair_scores.csv")

    people = load_people_csv(people_path)
    pair_scores = load_pair_scores_csv(pair_path)

    phones, pref, fallback = load_weekly_signups(args.excel, args.sheet)

    if not phones:
        print("No valid phone numbers found in the selected Excel sheet.", file=sys.stderr)
        return 2

    # Warn about signups not present in people.csv
    missing_people = [p for p in phones if p not in people]
    if missing_people:
        print("WARNING: These phones are in Excel but missing from people.csv (gender will be U, name from Excel):")
        for p in missing_people:
            print(f"  - {p}")
        print()

    groups = find_best_groups(
        phones=phones,
        pref=pref,
        people=people,
        pairs=pair_scores,
        gender_mindful=args.gender_mindful,
        restarts=args.restarts,
        seed=args.seed,
    )

    print_groups(groups, people, fallback, pair_scores, pref)

    resp = input("Approve these groups? Type 'yes' to approve: ").strip().lower()
    if resp not in ("yes", "y"):
        print("Not approved. No files were written and pair_scores.csv was not updated.")
        return 0

    # Determine output path
    if args.out:
        out_path = args.out
    else:
        os.makedirs("output", exist_ok=True)
        out_path = os.path.join("output", f"groups_{dt.date.today().isoformat()}.txt")

    # Write formatted output
    write_formatted_output(out_path, groups, people, fallback)
    print(f"\nWrote groups to: {out_path}")

    # Update pair_scores: decay first, then set new pairs to 3
    # (only after approval)
    backup_path = pair_path + f".bak_{dt.date.today().isoformat()}"
    if os.path.exists(pair_path):
        with open(pair_path, "rb") as src, open(backup_path, "wb") as dst:
            dst.write(src.read())
        print(f"Backed up pair_scores.csv to: {backup_path}")

    decay_scores(pair_scores)
    set_group_scores_to_3(pair_scores, groups)
    save_pair_scores_csv(pair_path, pair_scores)
    print(f"Updated pair scores: {pair_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
