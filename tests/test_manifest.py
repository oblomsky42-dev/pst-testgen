import csv
import json
import random

from pst_testgen.manifest import SCHEMA_VERSION, build_manifest, write_manifest
from pst_testgen.plan import message_spec


def _generated(n=40, seed=9):
    rng = random.Random(seed)
    return [message_spec(i, 2015 + i % 4, seed, rng) for i in range(n)]


def test_expected_values_follow_the_deletions(tmp_path):
    generated = _generated()
    soft = {generated[0]["subject"], generated[1]["subject"]}
    hard = {generated[2]["subject"], generated[3]["subject"], generated[4]["subject"]}
    m = build_manifest(str(tmp_path / "x.pst"), 9, {"delete_mode": "both", "delete_pct": 12.5},
                       generated, soft, hard)

    assert m["schema_version"] == SCHEMA_VERSION and m["pst_file"] == "x.pst"
    plain = m["expected"]["plain_scan"]
    assert plain["total_messages"] == 40 - 3
    assert sum(plain["by_year"].values()) == sum(plain["by_month"].values()) == 37
    assert all(len(k) == 7 and k[4] == "-" for k in plain["by_month"])  # "YYYY-MM"
    rec = m["expected"]["recover_scan"]
    assert (rec["max_total_messages"], rec["max_additional_recovered"]) == (40, 3)
    d = m["deletions"]
    assert (d["soft_deleted_count"], d["hard_deleted_count"]) == (2, 3)
    assert {i["subject"] for i in d["hard_deleted_items"]} == hard
    item = d["soft_deleted_items"][0]
    assert set(item) == {"subject", "date", "message_id"}
    assert item["subject"][1:11] == item["date"][:10]


def test_no_local_paths_leak_into_the_manifest(tmp_path):
    m = build_manifest(r"C:\Users\someone\Secret\case.pst", 1, {}, _generated(5), set(), set())
    assert m["pst_file"] == "case.pst"
    assert "someone" not in json.dumps(m)


def test_write_manifest(tmp_path):
    generated = _generated(10)
    m = build_manifest(str(tmp_path / "t.pst"), 9, {}, generated, set(), set())
    json_path, csv_path = write_manifest(m, str(tmp_path / "t.pst"))
    assert json_path.name == "t_expected_manifest.json"
    assert json.loads(json_path.read_text(encoding="utf-8")) == m
    rows = list(csv.reader(csv_path.open(encoding="utf-8-sig")))
    assert ["TOTAL", "10"] in rows
