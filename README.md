# PS-II Station Explorer · 2026-27

A single-file, offline browser for BITS Pilani **PS-II** station listings (Sem I, 2026-27).
Open `index.html` to browse, search, filter, and build a ranked preference list.

## Features
- Sortable table of stations with expandable detail (project, requirements, mentor/contact).
- Full-text search across station, project, description, mentor, domains, and tech.
- Filters: business domain, state, work mode, eligible branch, **project domain**, min stipend, CGPA.
- **★ Interaction flag** — marks companies that already ran an interaction-based selection this
  semester (from the crowdsourced PS Tracker), with "only / hide interaction" toggles.
- "My CGPA" eligibility highlighting.
- Drag-and-drop preference ranking drawer with CSV / JSON / shareable-link export.

## Data
`stations.json` is a PSMS portal extract (554 stations), enriched by `build_enrich.py` with:
- clean **business domain** + **project-domain tags** (from the project-titles sheet), and
- **visited / interaction** flags (from the PS Tracker), matched by company name.

Browsers block `fetch()` over `file://`, so serve the folder:

```bash
python3 -m http.server 8000   # then open http://localhost:8000
```

### Regenerating the data
```bash
python3 build_enrich.py            # dry run (prints match report)
python3 build_enrich.py --write    # rewrites stations.json in place
```
Source `.xlsx` paths are configurable via the `PROJ_XLSX` / `TRACK_XLSX` env vars.

> Unofficial student tool. Always verify details against the official PSMS portal.
