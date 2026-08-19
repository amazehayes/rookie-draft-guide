---
slug: dev-preview-workflow
kind: knowledge
updated: 2026-08-19
---

# Testing this site locally

**Bottom line:** static site, no build step — serve with a plain HTTP server
and hard-cache-bust every reload, because both the browser and the tool
chain cache aggressively.

- Preview tool config lives in the *parent* session dir's launch.json
  (`C:\Users\hayes\Documents\Claude Code\.claude\launch.json`), not inside
  this repo — entry name `rookie-draft-guide`, runs
  `python -m http.server 8765 --directory <this repo>`. Use
  `preview_start({name: "rookie-draft-guide"})`.
- **`navigate(..., force:true)` alone does NOT guarantee a fresh file** —
  hit a case where edited code kept serving the pre-edit version through
  several force-reloads. Fix: append a throwaway query string
  (`?nocache=<n>`) to the URL on every reload after an edit. The app's own
  `syncURL()` will rewrite the query string to its real page state right
  after load, which is expected, not a bug.
- If starting a plain `python -m http.server` manually in Bash for a
  one-off (e.g. testing an Artifact mockup before publishing), background
  it with `(cmd &) ; sleep 1` — a bare `cmd &` inside a `run_in_background`
  Bash call can report "completed" the instant the wrapper script's
  foreground part exits, even though the backgrounded server is still
  actually running. Verify with `curl -s -o /dev/null -w "%{http_code}"`
  before trusting it's up; don't assume "completed" means "died."
- The Browser pane's `computer` screenshot action is flaky in this
  environment — repeated 30s timeouts, sometimes clearing up after a
  `navigate` + retry, sometimes not for a whole session. When it's stuck,
  fall back to `get_page_text` and direct JS (`getBoundingClientRect`,
  `getComputedStyle`, reading `.textContent`/`.classList` on specific
  elements) — these are actually *more* reliable for verifying exact pixel
  geometry or computed values than eyeballing a screenshot anyway.
- Artifact mockups built as standalone `.html` files need an explicit
  `<meta charset="UTF-8">` — without it, em dashes and middle dots
  (`—`, `·`) render as mojibake when served via a plain local static
  server (the Artifact publishing tool's own wrapper apparently sets this
  for you, but a local `python -m http.server` preview does not).
