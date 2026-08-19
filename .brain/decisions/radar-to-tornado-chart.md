---
slug: radar-to-tornado-chart
kind: decision
updated: 2026-08-19
---

# Comparison page profile chart: radar → tornado

**Decision:** replaced the percentile radar chart with a "tornado" /
diverging-bar chart (`tornadoSection()`, index.html:3252) for the Player
Comparison page. Only renders for exactly 2 same-position players.

**Why:** Addison asked for radar alternatives; workshopped 4 options as a
Claude Artifact (vertical dot plot, horizontal dot plot, small-multiples
bar strip, tornado) using real percentile data before writing any real-site
code. He picked the tornado chart after seeing a reference screenshot of a
similar "2 QBs, opposing bars from a center line" chart style. Iterated on
it in the artifact first: added win/lose bar muting, a final score tally,
a crown on the winner's name, split into Production/Athleticism sections,
inline percentile+raw values, then compacted row height — all validated in
the artifact before porting to index.html.

**Rejected:**
- Dot plots and bar strip — still built as Options A-C in the same
  artifact for comparison, not shipped. Better for 3+ players (no
  direction-count limit) but Addison preferred the tornado's starker
  "who's winning" read for the common 2-player case.
- An "NFL Outcome Comps" bar-chart section (ceiling/median PPG + hit-rate,
  reusing `computeComps()`/`computeCompBand()`) was built into the tornado
  section, then explicitly removed at Addison's request — he didn't want
  it there. Don't re-add without being asked. (Separately, the player
  *modal's* own NFL comps section, a different feature, was fixed the same
  session — see `computeCompBand()` at index.html:4881, unrelated to this
  decision.)
- No fallback profile chart exists for 3-4 player or mixed-position
  comparisons — deliberate scope cut, flagged to Addison, he accepted it.
  Those cases just show the marker checklist + stat tables.
