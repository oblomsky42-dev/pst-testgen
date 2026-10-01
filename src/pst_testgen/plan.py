"""What to generate and what to delete - pure functions, no Outlook needed."""

from __future__ import annotations

import random
from datetime import datetime, timedelta

SENDERS = [
    "j.smith@example.com", "a.jones@example.com", "m.brown@example.com",
    "s.davis@example.com", "k.wilson@example.com", "r.taylor@example.com",
    "l.moore@example.com", "p.anderson@example.com", "c.thomas@example.com",
    "d.jackson@example.com", "external.partner@vendor.example",
    "notifications@system.example", "legal@corp.example.com",
]


def build_year_plan(start_year: int, end_year: int, total_messages: int, seed: int) -> dict:
    """Per-year message counts: ramp up, peak around 60% of the range, taper off."""
    rng = random.Random(seed)
    years = list(range(start_year, end_year + 1))
    n = len(years)
    weights = []
    for i in range(n):
        pos = i / max(1, n - 1)
        weights.append(max(0.15, 1.0 - abs(pos - 0.6) * 1.4) * rng.uniform(0.75, 1.25))
    wsum = sum(weights)
    plan = {y: int(total_messages * weights[i] / wsum) for i, y in enumerate(years)}
    plan[years[int(n * 0.6)]] += total_messages - sum(plan.values())
    return plan


def random_datetime_in_year(year: int, rng: random.Random) -> datetime:
    start = datetime(year, 1, 1)
    span = (datetime(year, 12, 31, 23, 59, 59) - start).total_seconds()
    return start + timedelta(seconds=rng.randint(0, int(span)))


def message_spec(index: int, year: int, seed: int, rng: random.Random) -> dict:
    """Subject, date, parties and Message-ID of the index-th message."""
    dt = random_datetime_in_year(year, rng)
    sender, recipient = rng.choice(SENDERS), rng.choice(SENDERS)
    return {
        "index": index,
        "date": dt,
        "sender": sender,
        "recipient": recipient,
        # the subject carries the date, so every item can be checked on its own
        "subject": f"[{dt:%Y-%m-%d}] Test message {index} from {sender}",
        "message_id": f"<test-{seed}-{index}@pst-testgen.example>",
    }


def choose_deletions(subjects: list[str], delete_pct: float, delete_mode: str,
                     rng: random.Random) -> tuple[set, set]:
    """(soft, hard) sets of subjects to delete. 'both' splits roughly in half."""
    n_delete = int(len(subjects) * delete_pct / 100)
    victims = rng.sample(subjects, n_delete) if n_delete else []
    if delete_mode == "soft":
        return set(victims), set()
    if delete_mode == "hard":
        return set(), set(victims)
    rng.shuffle(victims)
    half = len(victims) // 2
    return set(victims[half:]), set(victims[:half])
