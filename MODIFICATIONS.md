# Shared Modifications Log

All 6 sub-projects share the same submissions sheet, so a fix/change made to
one often needs to be copied to the other 5. Log each shared change here as
it's made, so it's easy to check which projects still need it applied.

Format:

## <short title> — YYYY-MM-DD
What changed and why.

Applied to:
- [ ] project-1
- [ ] project-2
- [ ] project-3
- [ ] project-4
- [ ] project-5
- [ ] project-6

---

## Cache-rebuild latency fix — 2026-08-05
Every write action was blocked by a hardcoded `Utilities.sleep(3000)` in
`rebuildDataCachePartial` (waiting for Sheets writes to "settle" before
reading them back) plus a redundant client-side `setTimeout(..., 3000)`
before reloading tables after submit — ~6s of dead time per action for
no benefit, since `submitData`/`submitStatusA`/`submitStatusB` already
block until the cache rebuild finishes. Replaced the sleep with
`SpreadsheetApp.flush()` (forces the pending write instead of guessing),
and removed the client-side delay entirely. Also merged the per-table
filter loops in each project's `getAllTableData` (were 6-7 separate full
passes over the cache array) into a single pass.

Applied to:
- [x] super-admin
- [x] companies (also fixed: duplicate `esc()` in index.html where a
      weaker second definition silently shadowed the safer one; and
      `submitStatusA`/`submitStatusB` reading the full `D:D` column
      instead of bounding to `getLastRow()`)
- [x] hr-admin (read-only app — no submit/write path, so only the
      loop-merge half of the fix applied; no sleep/setTimeout existed
      here to remove)
- [x] technician (also had a *second*, worse instance of the same
      anti-pattern: `getCompleteRowFromTable2` had `Utilities.sleep(3000)`
      immediately followed by a `SpreadsheetApp.flush()` that already
      did the real work — removed the redundant sleep. Runs on every
      "mark as Fixed" action, which also generates a PDF, so this was
      the slowest single action in the whole system: ~9s of pure
      removable padding, stacked on top of Drive API PDF-generation
      waits that were left alone)
- [ ] dneqpwhtsp-search
- [ ] team-leader

Still open (deferred, needs a decision before implementing — see chat
history 2026-08-05): stop rewriting the *entire* shared `DataCache`
sheet on every partial write (needs a per-source-sheet block index,
should land in all 6 apps at once, not incrementally); archiving old
closed rows out of `DataCache` (needs a retention policy — the
Teamleader Review table intentionally shows all Fixed/Not-Fixed rows,
including already-reviewed ones, as a history view).
