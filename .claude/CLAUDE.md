# Phone Dashboard Project Rules

## Version Management

**IMPORTANT**: After every commit to this project, update the version number:

1. Bump version in `phone-dashboard.html` (line ~18): `<span class="version-badge">vX.X.X</span>`
2. Bump the **matching `?v=` on both asset tags** in `phone-dashboard.html`:
   `dashboard.css?v=X.X.X` and `dashboard.js?v=X.X.X`. These must equal the badge.
   GitHub Pages serves the HTML fresh but browsers cache the unversioned assets,
   so skipping this ships new HTML against stale JS and the badge lies about
   which code is actually running.
3. Add entry to `IMPROVEMENTS.md` version history table

### Versioning Scheme
- **Major (X.0.0)**: Breaking changes or major feature overhauls
- **Minor (1.X.0)**: New features
- **Patch (1.1.X)**: Bug fixes

## Key Files

| File | Purpose |
|------|---------|
| `dashboard.js` | All dashboard logic, charts, filters, data processing |
| `phone-dashboard.html` | Page structure, version badge |
| `dashboard.css` | All styling |
| `IMPROVEMENTS.md` | Feature roadmap and version history |
| `garema/index.html` | Garema dashboard page (YOLO PBX), own version badge |
| `garema/garema-dashboard.js` | Garema dashboard logic, forked from `dashboard.js` |
| `garema/yolo-adapter.js` | Converts YOLO export columns into the row shape the dashboard logic expects |

## Garema Dashboard

- Garema (YOLO PBX, single queue) has its own version badge in `garema/index.html`,
  starting at v1.0.0. When Garema files change, bump that badge and the `?v=` on all
  three of its asset tags (`../dashboard.css`, `yolo-adapter.js`, `garema-dashboard.js`).
- It shares `dashboard.css` with the YourGP dashboard so both look the same. A CSS
  change therefore needs the `?v=` bumped on both pages.
- `garema-dashboard.js` is a fork of `dashboard.js`. Fixes to shared logic (heatmaps,
  charts, daily table, staff table) usually need applying to both files.
- If a YOLO column is not recognised, add its name to `FIELD_ALIASES` in
  `yolo-adapter.js` rather than changing the dashboard logic.

## Context

- Staff are **multi-tasking receptionists**, not call center agents
- Phone metrics are just one part of their work
- Avoid efficiency scores or metrics that unfairly penalise staff busy with in-person patients
