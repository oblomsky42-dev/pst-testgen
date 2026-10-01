# Changelog

## 1.0.0 — 2026-10-01

First release as a project of its own. Until now this was
`generate_test_pst.py`, one of three scripts that grew into
[email-date-range](https://github.com/oblomsky42-dev/email-date-range).

### Fixed
- Dates never stuck: the Outlook object model resets a new message's delivery
  time when it saves it, so every message carried the generation time. Messages
  are now written through Extended MAPI, and the first one is read back; the run
  stops if its date differs.
- Every message was first saved into the default mailbox's Drafts and then
  moved, and stayed there when the move failed. Nothing touches the default
  mailbox any more.
- The manifest expected recovery tools to find every hard-deleted message. With
  Outlook 2016+ nothing is left to find, so the hard-deleted set is now an upper
  bound.
- An error during generation left the test PST attached to the Outlook profile;
  it is now always detached, and Outlook is closed if the run started it (it
  keeps the PST locked otherwise).
- The manifest recorded the PST's full local path; it now records the file name.

### Added
- `--count N`; every message has a sender, a recipient, read state and a
  Message-ID `<test-SEED-N@pst-testgen.example>`.
- Manifest `schema_version` 1 with the generation parameters.
- Refuses to write into an existing PST; argument checks.
- Unit tests and CI for everything that does not need Outlook.

### Changed
- `--target-gb` defaults to 1 GB instead of 20 GB; `--delete-mode` defaults
  to `both`.
