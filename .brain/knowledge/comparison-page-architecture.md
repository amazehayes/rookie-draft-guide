---
slug: comparison-page-architecture
kind: knowledge
updated: 2026-08-19
---

# Player Comparison page (post 2026-08-19 rebuild)

**Bottom line:** stat tables are transposed (player=row, stat=column,
horizontal scroll), and the old radar chart is gone — replaced by a
"tornado" diverging-bar chart that only renders for exactly 2 same-position
players. Everything lives inside `renderCompStats()` at index.html:3166.

- **Stat tables** — `statSection()` (3190): one `<table>` per section
  (Overview, then position-specific), sticky first column shows headshot +
  name + year/pos/round (+ label badge, Overview only). Winner per stat
  column gets bold+color+▲; `row.noWinner: true` on a stat (e.g. slot rate)
  suppresses that entirely — higher isn't "better" for descriptive stats.
- **Tornado chart** — `tornadoSection()` (3252), gated on
  `samePos && players.length === 2` only (bars need exactly 2 directions).
  3-4 player and mixed-position comparisons skip it, no fallback chart —
  just marker checklist + tables. Data comes from `COMP_TORNADO_STATS`
  (2904): per position, split into `production[]` and `athleticism[]`
  arrays. Athleticism (BMI, 40-yard dash, Speed Score, Burst Score) is now
  identical across all 4 positions, even QB (verified 154/314 QBs have full
  combine data — enough for `positionPercentile()` at 2983 to work).
  Pure HTML/CSS, no Chart.js — bars are divs with inline `width:${pctl}%`.
  Percentile + raw value render inline on one line ("79 (52.7)"), not
  stacked, to keep row height small (~23px/row, whole chart ~475px total).
- **Marker checklist** — `markerHeatmapSection()` (3311): same-position,
  2-4 players OK (not gated to exactly-2 like tornado). Shows every
  `evalMarkers()` checklist item as ✓/✗ per player.
- **Watchlist** — `watchStarHTML()`/`toggleWatchStar()` (2109/2116), state
  in `let watchlist` (1547, localStorage-backed). Star buttons appear on
  cards/tiers/table/rankings rows and the player modal; "★ Watchlist Only"
  filter toggle on Grades and Rankings pages, `?watch=1` in the URL.
- **URL state** — `syncURL()` (5900) / `applyURLState()` (5913): every page
  syncs its filters to the query string (`?page=comparison&p1=slug&p2=slug`
  for this page). Nav clicks pushState; in-page filter changes replaceState.
