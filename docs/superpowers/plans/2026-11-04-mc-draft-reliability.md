# MC_Draft reliability implementation plan

**Goal:** Fix the six reviewed issues without a redesign or mobile-specific work.
**Architecture:** Retain the existing static HTML app and Excel schema. Add local draft recovery and visible generation warnings. Deliver a real downloadable workbook using the existing worksheet names and columns.
**Tech stack:** Vanilla JavaScript, SheetJS, Node's built-in test runner; Artifact Tool for the template artifact, openpyxl for independent read-only inspection.
**Spec:** User approved the six fixes in the preceding audit and excluded iPad/iPhone compatibility work.

## Constraints
- Keep department ordering, external award ordering, manual editing and the existing colour scheme.
- No account, server, cloud data storage, or unrelated feature work.
- Keep last-row no-photo semantics; explain them accurately.
- Work from main 85b8a79369cecf778299485999cfbe01d8fa431a on an isolated fix branch.

## Review focus
- Mixed synonymous awards must sort correctly; unknown awards retain their relative order.
- A restored manual draft must survive regeneration cancellation and failed imports.
- Missing/invalid workbook data must be visible without opening another tab.
- Speech-only workbooks must work without fake award records.
- Unavailable/full localStorage must not prevent editing or export.

## Tasks
1. Write and run regression tests for ranking and pure-speech/invalid award data. Fix ranking normalization and empty award handling. Verify all tests.
2. Write and run draft recovery/overwrite/reset tests. Add local draft save/restore, navigation protection, and confirmation before destructive local actions. Verify all tests.
3. Write and run visible-generation-warning tests. Put issue count and errors above the editable script, retain issue details and normalize harmless nature whitespace. Verify all tests.
4. Add a downloadable v2.5 workbook and accurate concise field instructions. Verify the workbook headers, examples, drop-downs and generated outputs.
5. Run the complete suite, syntax and whitespace checks; inspect browser interaction if available, obtain an independent code review, then push the verified fix branch and create a reviewable PR.

## Progress
- Baseline reviewed: repository contains only index.html and favicon.svg; no existing tests.
- Baseline reproductions: champion/second-place ordering is reversed; pure speech requires awards; regeneration overwrites manual edits; invalid nature silently drops an event behind an inactive tab.
- Tasks 1–4 complete: six approved fixes implemented. The workbook's four sheets were rendered and inspected; official input sheets are empty of fictional awards, examples isolated, identifiers preserved as text.
- Independent whole-branch review identified two Important issues: edits during asynchronous file reads and invalid empty imports. Both reproduced by failing tests and fixed; three regression tests added.
- Review also found a recipient omission, unsorted class awards, missing team unknown-rank reminders, and inaccurate boundary/version text. These were within the approved ranking/omission/documentation scope and fixed in the same review pass; three behavior tests added.
- Final review notes: original ceremony/no-photo semantics retained; no mobile-specific work. Cloud browser could not reach the local preview (ERR_CONNECTION_REFUSED), so no completed browser QA is claimed.
- Current verification: 33 Node regression tests passing; inline JavaScript syntax and git whitespace checks passing. Workbook native headers, empty input records, identifiers, validation ranges and isolated examples independently checked. Actual SheetJS 0.18.5 also parsed the real exported workbook: blank-template import and speech-only generation passed; copying the isolated example sheet produced all five recipients across three event types with leading-zero IDs and correct mixed ranking.
