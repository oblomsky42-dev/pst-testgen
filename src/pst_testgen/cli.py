r"""pst-testgen: build a PST whose right answer is known, plus a manifest of that answer.

    pst-testgen --out C:\test\t.pst --count 2000 --delete-pct 10 --delete-mode both

Deletion modes:
  soft  moved to Deleted Items - still in the folder tree, a plain scan sees it.
  hard  also purged from Deleted Items (Shift+Delete) - a plain scan does not see
        it. Whether anything stays recoverable is up to the PST writer: Outlook
        2016+ drops the index entries at once, so the manifest gives the
        hard-deleted set as an upper bound for recovery tools, not a target.
  both  roughly half each.
"""

from __future__ import annotations

import argparse
import os
import random
import sys

from . import __version__
from .manifest import build_manifest, write_manifest
from .outlook import HAS_WIN32, MapiFolderWriter, OutlookPst, delete_messages
from .plan import build_year_plan, choose_deletions, message_spec

FOLDER_NAME = "TestMessages"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pst-testgen",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Generate a PST with a known date distribution (and optional "
                    "deletions) through Outlook, plus an expected-values manifest.",
        epilog="Requires Windows, classic desktop Outlook and pywin32. "
               "The PST must not exist yet.")
    p.add_argument("--out", required=True, help=r"output PST path, e.g. C:\test\t.pst")
    p.add_argument("--count", type=int, metavar="N",
                   help="generate exactly N messages (otherwise fill --target-gb)")
    p.add_argument("--target-gb", type=float, default=1.0,
                   help="target file size in GB when --count is not given (default: 1)")
    p.add_argument("--start-year", type=int, default=2015)
    p.add_argument("--end-year", type=int, default=2025)
    p.add_argument("--seed", type=int, default=1337, help="RNG seed: same seed, same PST content")
    p.add_argument("--attach-every", type=int, default=50,
                   help="add a random-bytes attachment to every N-th message (default: 50)")
    p.add_argument("--attach-mb", type=int, default=5,
                   help="attachment size in MB (default: 5; 0 = no attachments)")
    p.add_argument("--body-kb", type=int, default=8, help="body ballast per message in KB")
    p.add_argument("--delete-pct", type=float, default=0.0, metavar="PCT",
                   help="delete this %% of the messages afterwards (default: 0)")
    p.add_argument("--delete-mode", choices=["soft", "hard", "both"], default="both",
                   help="soft = to Deleted Items, hard = purged (Shift+Delete), both = ~50/50")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def validate(args) -> str | None:
    if args.end_year < args.start_year:
        return "--end-year must be >= --start-year"
    if args.attach_every < 1 or args.attach_mb < 0 or args.body_kb < 0:
        return "--attach-every must be >= 1, --attach-mb and --body-kb >= 0"
    if args.count is not None and args.count < 1:
        return "--count must be >= 1"
    if not 0 <= args.delete_pct <= 100:
        return "--delete-pct must be between 0 and 100"
    if os.path.exists(args.out):
        return (f"{args.out} already exists - Outlook would add to it and the manifest "
                "would no longer describe the file")
    return None


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    problem = validate(args)
    if problem:
        print(f"error: {problem}")
        return 1
    if not HAS_WIN32:
        print("error: pst-testgen needs Windows with classic Outlook and pywin32 "
              "(pip install pywin32)")
        return 1

    if args.count:
        total, target_bytes = args.count, float("inf")
        print(f"Messages      : {total:,}")
    else:
        target_bytes = int(args.target_gb * 1024 ** 3)
        per_message = args.body_kb * 1024 + 2048 + args.attach_mb * 1024 ** 2 / args.attach_every
        total = int(target_bytes / per_message)
        print(f"Target size   : {args.target_gb} GB (~{total:,} messages)")
    print(f"Years         : {args.start_year}-{args.end_year}")
    print(f"Attachments   : {args.attach_mb} MB every {args.attach_every} messages")
    print(f"Seed          : {args.seed}")
    plan = build_year_plan(args.start_year, args.end_year, total, args.seed)

    print("\nConnecting to Outlook...")
    pst = OutlookPst(args.out)
    print(f"PST           : {pst.path}")
    try:
        return populate(pst, args, plan, target_bytes)
    finally:
        # also on errors and Ctrl+C: never leave the test PST attached to the profile
        pst.close()


def populate(pst: OutlookPst, args, plan: dict, target_bytes: float) -> int:
    rng = random.Random(args.seed)
    folder = pst.folder(FOLDER_NAME)
    attach_path = None
    if args.attach_mb > 0:
        attach_path = os.path.join(os.path.dirname(pst.path), "_ballast_attachment.bin")
        with open(attach_path, "wb") as fh:
            block = os.urandom(1024 * 1024)  # random, so PST compression cannot shrink it
            for _ in range(args.attach_mb):
                fh.write(block)
    filler = "X" * (args.body_kb * 1024)
    years = [y for y in sorted(plan) for _ in range(plan[y])]
    rng.shuffle(years)

    generated = []
    writer = MapiFolderWriter(pst.store.StoreID, folder.EntryID)
    print("\nGenerating... (Ctrl+C stops early; the PST stays valid)")
    try:
        for index, year in enumerate(years):
            spec = message_spec(index, year, args.seed, rng)
            body = (f"From: {spec['sender']}\nTo: {spec['recipient']}\n"
                    f"Date: {spec['date']:%Y-%m-%d %H:%M:%S}\n\n{filler}")
            with_attachment = attach_path if index % args.attach_every == 0 else None
            stored = writer.create(spec, body, with_attachment)
            if index == 0:
                if stored != spec["date"]:
                    print(f"  [date check] FAILED - wrote {spec['date']}, the store kept "
                          f"{stored}; stopping, the test PST would be invalid.")
                    return 1
                print(f"  [date check] OK - delivery time stored exactly ({stored} UTC)")
            generated.append(spec)
            if (index + 1) % 500 == 0:
                size = os.path.getsize(pst.path)
                print(f"  {index + 1:,} messages, {size / 1024 ** 2:,.0f} MB", flush=True)
                if size >= target_bytes:
                    print("  target size reached")
                    break
    except KeyboardInterrupt:
        print("\nInterrupted - the PST is valid with the messages written so far.")
    finally:
        writer.close()
        if attach_path and os.path.exists(attach_path):
            os.remove(attach_path)
    print(f"Generated {len(generated):,} messages.")

    soft, hard, failed = set(), set(), 0
    if args.delete_pct > 0 and generated:
        soft, hard = choose_deletions([m["subject"] for m in generated],
                                      args.delete_pct, args.delete_mode, rng)
        deleted_items = pst.deleted_items()
        if hard and deleted_items is None:
            print("  WARNING: Deleted Items not found - hard deletes stay soft.")
            soft, hard = soft | hard, set()
        failed = delete_messages(folder, deleted_items, soft, hard)
        print(f"Deleted: {len(soft):,} soft (in Deleted Items), {len(hard):,} hard (purged)"
              + (f", {failed} operation(s) FAILED" if failed else ""))

    parameters = {k: getattr(args, k) for k in (
        "count", "target_gb", "start_year", "end_year", "attach_every", "attach_mb",
        "body_kb", "delete_pct", "delete_mode")}
    manifest = build_manifest(pst.path, args.seed, parameters, generated, soft, hard, failed)
    json_path, _csv_path = write_manifest(manifest, pst.path)
    print(f"Manifest: {json_path}")
    print("\n!! Do not compact this PST or let Outlook run maintenance on it.")
    return 0 if not failed else 2


def entry_point() -> None:
    sys.exit(main())
