# ICIMC 2026 — HTML deck

Open `index.html` in Chrome or Edge (whole `html/` folder must stay together).

| Key | Action |
|---|---|
| ← → / Space | Next / previous slide |
| S | Presenter window: current + next slide, script (EN + KO), timer |
| N | Script drawer at the bottom of the audience window |
| F | Fullscreen |
| O | Slide overview |

Built with the `html-ppt` skill (`presenter-mode-reveal` pattern) on a fixed
1920×1080 stage that scales to any screen. Charts are inline SVG generated from
the manuscript data; hover a point or bar to see its value.

Rebuild after editing data or text: `node build-html.js` (no dependencies).
Web fonts load from Google Fonts; offline, the deck falls back to Georgia / Arial.
