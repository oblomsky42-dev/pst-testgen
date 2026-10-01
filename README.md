# pst-testgen

[![CI](https://github.com/oblomsky42-dev/pst-testgen/actions/workflows/ci.yml/badge.svg)](https://github.com/oblomsky42-dev/pst-testgen/actions/workflows/ci.yml)
![Platform](https://img.shields.io/badge/platform-Windows%20%2B%20Outlook-blue)
![License](https://img.shields.io/badge/license-MIT-green)

*[Русская версия](README.ru.md)*

Generates Outlook PST files whose right answer is known — every message's
delivery date to the second, the counts per year and per month, which
messages were moved to Deleted Items and which were purged — and writes that
answer to a manifest next to the PST.

Tools that read mail archives — forensic scanners, eDiscovery processing,
libpff-based parsers — need ground truth to be tested against. Real mailboxes
do not come with one, and public corpora such as Enron have no answer key for
dates or deletions. pst-testgen makes PSTs that do.

It was built to test [email-date-range](https://github.com/oblomsky42-dev/email-date-range),
whose CI checks every message of a pst-testgen PST against its manifest.

## Requirements

- Windows with classic desktop Outlook (2016 or later; developed against
  Outlook 2024). The "new Outlook" has no COM interface and cannot be used.
- Python 3.10+ and pywin32 (installed with the package).

The default mailbox is not touched. The new PST is attached to the default
Outlook profile for the length of the run and detached at the end; if the run
started Outlook, it also closes it, because Outlook keeps a PST locked until
its process exits.

## Install

```bash
pip install git+https://github.com/oblomsky42-dev/pst-testgen
```

or take `pst_testgen.exe` from the
[Releases](https://github.com/oblomsky42-dev/pst-testgen/releases) page
(it still needs Outlook on the machine).

## Usage

```bash
pst-testgen --out C:\test\t.pst --count 2000 --delete-pct 10 --delete-mode both
pst-testgen --out D:\test\big.pst --target-gb 5 --attach-mb 5 --attach-every 50
```

| Option | Default | Meaning |
|---|---|---|
| `--out` | — | the new PST; it must not exist yet |
| `--count N` | — | exactly N messages; without it the PST is filled up to `--target-gb` |
| `--target-gb` | 1 | target file size |
| `--start-year`, `--end-year` | 2015, 2025 | date range; the per-year counts ramp up and taper off |
| `--seed` | 1337 | same seed → same dates, subjects, senders and deletions |
| `--attach-every`, `--attach-mb` | 50, 5 | a random-bytes attachment on every N-th message (0 MB = none) |
| `--body-kb` | 8 | body ballast per message |
| `--delete-pct` | 0 | share of messages to delete after generation |
| `--delete-mode` | both | `soft` = to Deleted Items, `hard` = purged (Shift+Delete), `both` = half each |

Each message goes into a `TestMessages` folder with:

- subject `[YYYY-MM-DD] Test message N from <sender>` — the date in the
  subject lets a tool's output be checked message by message;
- delivery and submit time set to that date, in UTC, to the second;
- read and received state (not a draft), a sender and a recipient from a
  fixed list of `example.com`-style addresses, and the Message-ID
  `<test-SEED-N@pst-testgen.example>`.

## The manifest

`<name>_expected_manifest.json` (and a per-year `.csv` summary):

| Key | Content |
|---|---|
| `schema_version` | `1` |
| `generator`, `generated_utc`, `seed`, `parameters` | how the file was made |
| `pst_file` | the PST's file name — never a full path, so the manifest can be shared |
| `total_messages_created`, `by_year_all_created` | everything written |
| `expected.plain_scan` | `total_messages`, `by_year`, `by_month` ("YYYY-MM") — what a scan of the folder tree must report: live plus soft-deleted messages |
| `expected.recover_scan` | `max_total_messages`, `max_additional_recovered` — the upper bound for a tool that also reads orphaned nodes (see below) |
| `deletions` | counts plus `soft_deleted_items` and `hard_deleted_items`, each with `subject`, `date` and `message_id` |

## Deletions, and what is left of them

`soft` deletes move a message to Deleted Items: it stays in the folder tree and
every tool must still see it. `hard` deletes purge it from Deleted Items as
well, like Shift+Delete.

What a purged message leaves behind is up to whatever wrote the PST. Outlook
2016 and later drop it from the PST's indexes at once. In a test with Outlook
2024, 1,800 purged messages left no orphaned nodes and nothing that `pffexport
-m recovered` could rebuild. The manifest therefore treats the hard-deleted set
as an upper bound for recovery, not as a target: a tool that "recovers" a
message that is not in `hard_deleted_items` is wrong; one that recovers none
of them may well be right.

## Why Extended MAPI

The obvious approach — create a `MailItem` through the Outlook object model
and set `PR_MESSAGE_DELIVERY_TIME` through `PropertyAccessor` — does not work:
Outlook resets the delivery time to "now" when it saves a new item, whether
the property is set before or after the first save, on a `MailItem` or a
`PostItem`. The write is silently ignored.

pst-testgen uses Outlook's object model only to create the PST and its folder
and to delete. The messages themselves are written through Extended MAPI
(pywin32's `win32com.mapi`): `IMAPIFolder::CreateMessage`, `SetProps` with the
delivery and submit time, `PR_MESSAGE_FLAGS = MSGFLAG_READ`, the recipient
table through `ModifyRecipients`, then `SaveChanges`. MAPI keeps every value
as given. The first message is read back after saving, and the run stops if
its delivery time is not exactly the one written — a test PST with wrong dates
is worse than none.

## Validation

PSTs of 400, 2,000 and 3,000 messages generated with Outlook 2024 were read
back with libpff (through email-date-range) and matched their manifests
exactly: totals, per-year and per-month counts, every message's timestamp and
Message-ID, and the placement of every soft-deleted message. The 400-message
file ships as a CI fixture in email-date-range after an audit for personal
data: every property libpff can reach, the raw file and the permute-decoded
file were searched for the generating machine's user, host and account names
and paths — with a positive control to show the search finds what is there.

The unit tests (plan, deletion choice, manifest, argument checks) run in CI on
Windows; the Outlook part cannot run there and is tested by hand.

## Limitations

- Windows and classic Outlook only.
- The file size for `--target-gb` is checked every 500 messages, so the
  result overshoots slightly.
- Do not compact the PST or let Outlook run maintenance on it afterwards — both
  rewrite the file, and the manifest no longer describes it.

## Development

```bash
pip install -e ".[test]" ruff
pytest
ruff check src tests
```

## License

MIT — see [LICENSE](LICENSE).
