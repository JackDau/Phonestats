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

## Context

- Staff are **multi-tasking receptionists**, not call center agents
- Phone metrics are just one part of their work
- Avoid efficiency scores or metrics that unfairly penalise staff busy with in-person patients
