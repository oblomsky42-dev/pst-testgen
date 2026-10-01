import pytest

from pst_testgen import __version__
from pst_testgen.cli import build_parser, main, validate


def _args(tmp_path, *extra):
    return build_parser().parse_args(["--out", str(tmp_path / "new.pst"), *extra])


@pytest.mark.parametrize("extra, message", [
    (["--start-year", "2020", "--end-year", "2019"], "--end-year"),
    (["--count", "0"], "--count"),
    (["--delete-pct", "150"], "--delete-pct"),
    (["--attach-every", "0"], "--attach-every"),
])
def test_bad_arguments_are_rejected(tmp_path, extra, message):
    assert message in validate(_args(tmp_path, *extra))


def test_good_arguments_pass(tmp_path):
    assert validate(_args(tmp_path, "--count", "10", "--delete-pct", "10")) is None


def test_existing_output_is_refused_before_outlook_starts(tmp_path, capsys):
    existing = tmp_path / "already.pst"
    existing.write_bytes(b"")
    assert main(["--out", str(existing), "--count", "5"]) == 1
    assert "already exists" in capsys.readouterr().out


def test_version(capsys):
    with pytest.raises(SystemExit):
        main(["--version"])
    assert __version__ in capsys.readouterr().out
