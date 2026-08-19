---
slug: init-order-tdz-trap
kind: knowledge
updated: 2026-08-19
---

# index.html top-level init-order trap

**Bottom line:** any `let`/`const` read by code that runs during the script's own
top-level execution (not inside a later user-triggered handler) must be declared
*before* that init code runs, or it throws `ReferenceError: Cannot access 'x'
before initialization` — silently, if the call site is wrapped in try/catch.

- The script executes top-to-bottom once on page load. Real init calls
  (`buildYearTabs(); renderHeroStats(); render();` then `applyURLState(true);`)
  live near the very end of the file (~line 5900+), specifically so every
  function/var they touch already exists.
- Found this pattern 3 times in one session (2026-08-19):
  1. A *pre-existing* dormant bug: the old "restore last-viewed page on
     refresh" code called `switchPage(saved)` too early (~old line 2531),
     before `let currentFormat`/`bkPosFilter`/etc. were declared — wrapped in
     try/catch, so it silently never worked. Fixed by moving all restore
     logic into `applyURLState()` at the true end of the script.
  2. Self-inflicted: added `syncURL()` inside `render()`, but `render()` is
     called at top-level before `let activePage` (declared much later at the
     time) existed. Fixed by moving `activePage`'s declaration up near
     `currentPos` etc. at index.html:1542, with a comment explaining why.
  3. Same class of risk for `watchlist` (index.html:1547) — declared early
     for the same reason, since `renderCards`/`renderTiers` (called by the
     top-level `render()`) reference it via `watchStarHTML()`.
- **Rule for future work:** any new page-level state var only needs to be
  declared near its own section *unless* it's read by a function that the
  top-level init block calls (`render`, `renderHeroStats`, `buildYearTabs`,
  or anything `applyURLState` calls). If so, declare it early and say why.
- Diagnosing this: `window.addEventListener('error', e => window.__errs.push(...))`
  inserted at the very top of the `<script>` block (before anything else)
  is what actually surfaced #2 — a plain try/catch around a JS call from the
  browser devtools console reported a *different*, more confusing error
  (`Cannot access before initialization` vs `is not defined`) depending on
  timing, so don't trust that message alone; capture the real stack trace.
