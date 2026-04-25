#!/usr/bin/env python3
import csv
import hashlib
import sqlite3
from pathlib import Path


BASE_DIR = Path(
    "/Users/doujinzhuang/Library/CloudStorage/OneDrive-UAB-TheUniversityofAlabamaatBirmingham/DouLab_UAB/Service/ATTIS/final/review_site"
)
DB_PATH = BASE_DIR / "review_portal.db"
INPUT_CSV = BASE_DIR / "reviewer_blazer_ids.csv"


def hash_secret(secret: str) -> str:
    return hashlib.sha256(secret.encode("utf-8")).hexdigest()


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout = 30000;")
    return conn


def main() -> None:
    if not INPUT_CSV.exists():
        raise FileNotFoundError(f"Missing file: {INPUT_CSV}")

    rows = []
    with INPUT_CSV.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            reviewer = (row.get("Reviewer") or "").strip()
            blazer_id = (row.get("BlazerID") or "").strip()
            if reviewer and blazer_id:
                rows.append((reviewer, hash_secret(blazer_id)))

    if not rows:
        raise ValueError("No valid reviewer/BlazerID rows found.")

    with get_conn() as conn:
        known = {
            r["reviewer"]
            for r in conn.execute("SELECT DISTINCT reviewer FROM assignments").fetchall()
        }
        updated = 0
        missing = []
        for reviewer, pin_hash in rows:
            if reviewer not in known:
                missing.append(reviewer)
                continue
            conn.execute(
                "INSERT OR REPLACE INTO reviewer_auth (reviewer, pin_hash) VALUES (?, ?)",
                (reviewer, pin_hash),
            )
            # force re-login after password change
            conn.execute("DELETE FROM sessions WHERE reviewer = ?", (reviewer,))
            updated += 1

    print(f"Updated passwords for {updated} reviewers.")
    if missing:
        print("Names not found in assignments:")
        for name in missing:
            print(" -", name)


if __name__ == "__main__":
    main()
