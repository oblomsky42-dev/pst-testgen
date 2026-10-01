"""The expected-values manifest written next to every generated PST.

schema_version 1 - see README.md, "Manifest format". Consumers should check
schema_version and ignore keys they do not know.
"""

from __future__ import annotations

import csv
import json
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from . import __version__

SCHEMA_VERSION = 1


def _item(m: dict) -> dict:
    return {"subject": m["subject"], "date": m["date"].strftime("%Y-%m-%d %H:%M:%S"),
            "message_id": m["message_id"]}


def build_manifest(pst_file: str, seed: int, parameters: dict, generated: list[dict],
                   soft: set, hard: set, failed_operations: int = 0) -> dict:
    """Expected values for a scan of the PST. ``generated`` holds message_spec() dicts."""
    created = Counter(m["date"].year for m in generated)
    visible_year = Counter(m["date"].year for m in generated if m["subject"] not in hard)
    visible_month = Counter(f"{m['date']:%Y-%m}" for m in generated if m["subject"] not in hard)
    hidden_year = Counter(m["date"].year for m in generated if m["subject"] in hard)

    def by_key(counter):
        return {str(k): counter[k] for k in sorted(counter)}

    return {
        "schema_version": SCHEMA_VERSION,
        "generator": f"pst-testgen {__version__}",
        "pst_file": Path(pst_file).name,  # a bare name: no local paths in shared fixtures
        "generated_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "seed": seed,
        "parameters": parameters,
        "total_messages_created": len(generated),
        "year_range": [min(created), max(created)] if created else [],
        "by_year_all_created": by_key(created),
        "expected": {
            "plain_scan": {
                "note": "a scan without recovery sees live and soft-deleted messages",
                "total_messages": sum(visible_year.values()),
                "by_year": by_key(visible_year),
                "by_month": by_key(visible_month),
            },
            "recover_scan": {
                "note": "a scan that also reads orphaned nodes may add some of the "
                        "hard-deleted messages - between 0 and max_additional_recovered, "
                        "every one listed in deletions.hard_deleted_items. Outlook 2016+ "
                        "normally leaves none.",
                "max_total_messages": len(generated),
                "max_additional_recovered": len(hard),
                "by_year_if_all_recovered": by_key(created),
            },
        },
        "deletions": {
            "enabled": bool(soft or hard),
            "mode": parameters.get("delete_mode") if (soft or hard) else None,
            "requested_pct": parameters.get("delete_pct", 0),
            "soft_deleted_count": len(soft),
            "hard_deleted_count": len(hard),
            "failed_operations": failed_operations,
            "hard_deleted_by_year": by_key(hidden_year),
            "soft_deleted_items": [_item(m) for m in generated if m["subject"] in soft],
            "hard_deleted_items": [_item(m) for m in generated if m["subject"] in hard],
        },
        "warnings": [
            "Do NOT compact this PST or let Outlook run maintenance on it - both rewrite "
            "the file, and the manifest then no longer describes it.",
        ],
    }


def write_manifest(manifest: dict, pst_path: str) -> tuple[Path, Path]:
    """Write <stem>_expected_manifest.json and .csv next to the PST."""
    stem = Path(pst_path).with_suffix("")
    json_path = Path(f"{stem}_expected_manifest.json")
    csv_path = Path(f"{stem}_expected_manifest.csv")
    json_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    plain = manifest["expected"]["plain_scan"]
    sections = (
        ("=== Expected: plain scan ===", plain["by_year"], plain["total_messages"]),
        ("=== Upper bound with recovery (every hard delete recovered) ===",
         manifest["by_year_all_created"], manifest["total_messages_created"]),
        ("=== Hard-deleted (invisible to a plain scan) ===",
         manifest["deletions"]["hard_deleted_by_year"],
         manifest["deletions"]["hard_deleted_count"]),
    )
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        for title, counts, total in sections:
            w.writerow([title])
            w.writerow(["year", "count"])
            for year, n in counts.items():
                w.writerow([year, n])
            w.writerow(["TOTAL", total])
            w.writerow([])
    return json_path, csv_path
