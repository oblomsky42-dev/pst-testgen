import random
import re

from pst_testgen.plan import build_year_plan, choose_deletions, message_spec, random_datetime_in_year


def test_year_plan_is_complete_and_reproducible():
    plan = build_year_plan(2015, 2025, 10_000, seed=42)
    assert sorted(plan) == list(range(2015, 2026))
    assert sum(plan.values()) == 10_000
    assert all(n > 0 for n in plan.values())
    assert plan == build_year_plan(2015, 2025, 10_000, seed=42)
    assert plan != build_year_plan(2015, 2025, 10_000, seed=43)


def test_single_year_plan():
    assert build_year_plan(2020, 2020, 5, seed=1) == {2020: 5}


def test_random_datetime_stays_in_year():
    rng = random.Random(0)
    assert all(random_datetime_in_year(2016, rng).year == 2016 for _ in range(1000))


def test_message_spec_carries_its_date_and_id():
    spec = message_spec(7, 2019, seed=5, rng=random.Random(1))
    assert spec["date"].year == 2019
    assert spec["subject"].startswith(f"[{spec['date']:%Y-%m-%d}] Test message 7 from ")
    assert re.fullmatch(r"<test-5-7@pst-testgen\.example>", spec["message_id"])
    assert spec["sender"].endswith(".example") or spec["sender"].endswith(".com")


def test_choose_deletions():
    subjects = [f"s{i}" for i in range(200)]
    soft, hard = choose_deletions(subjects, 10, "soft", random.Random(3))
    assert (len(soft), len(hard)) == (20, 0)
    soft, hard = choose_deletions(subjects, 10, "hard", random.Random(3))
    assert (len(soft), len(hard)) == (0, 20)
    soft, hard = choose_deletions(subjects, 25, "both", random.Random(3))
    assert len(soft) + len(hard) == 50 and abs(len(soft) - len(hard)) <= 1
    assert not soft & hard and (soft | hard) <= set(subjects)
    assert choose_deletions(subjects, 25, "both", random.Random(3)) == (soft, hard)
    assert choose_deletions(subjects, 0, "both", random.Random(3)) == (set(), set())
