---
slug: nfl-comp-band-season-gate
kind: knowledge
updated: 2026-08-19
---

# Player modal "Projected Outcome Range" band

**Bottom line:** `computeCompBand()` (index.html:4881) used to require each
of the 10 shown comps to have **3+ NFL seasons** before counting toward the
5-minimum needed to render the band at all — for players whose comp pool
skews toward recent draft classes, this silently hid the whole section.
Fixed 2026-08-19: now counts any comp with `nfl_career_ppg != null`
(any NFL data, regardless of season count), same threshold logic (5 min)
otherwise. Example that surfaced it: De'Zhaun Stribling's 10 comps all had
real PPG data but only 3 had played 3+ seasons — band was `null` before the
fix, populated with `n:10` after.

- The per-comp-card text line under each mini-card (`renderModalComps()`,
  ~4954) was **never** gated this way — it already showed NFL stats for any
  comp with `_nfl` data. Only the aggregate band above the grid had the
  stale 3-season filter. If something like this recurs elsewhere, check for
  similarly-named "mature"/season-count filters before assuming the bug is
  in the display layer rather than the data-filtering layer.
