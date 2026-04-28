---
inclusion: always
---

# Tech Stack & Build

## Stack
- Pure HTML/CSS/JS — no framework, no bundler for the portfolio site itself
- Individual tools (external repos) use React + TypeScript + Vite
- Node.js scripts for build automation and content updates
- Simple Node.js HTTP server for local development (`server.js`)

## Key Files
- `index.html` — root portfolio page (source of truth, public/GitHub Pages version)
- `html/index.html` — auto-generated intranet version (DO NOT EDIT DIRECTLY)
- `build-internal.js` — generates `html/index.html` from root `index.html`
- `scripts/update-feedback.js` — batch-updates feedback systems across one-pager HTML files
- `server.js` — local dev server on port 3456

## Build Commands
```bash
# Generate intranet version of index.html
node build-internal.js

# Update feedback integration across all one-pager HTML files
node scripts/update-feedback.js

# Local dev server
node server.js
# → http://localhost:3456
```

## Deployment
- Public: GitHub Pages at `avi1840.github.io`
- Internal: intranet at `*.snifim.blroot` (redirect files in `html/redirects/`)
- The `build-internal.js` script adapts paths for intranet deployment where `html/` is the web root

## Feedback System
- All one-pager pages include a built-in feedback panel (ratings + text)
- Feedback is stored in localStorage and sent to a shared Google Sheet via Apps Script
- The Google Sheet URL (`SHEET_URL`) must never be changed
- For React/TS tools, use the `FeedbackModal` component pattern (see steering rule)

## Conventions
- All HTML pages are self-contained (inline CSS + JS, no external build step)
- Hebrew RTL throughout (`dir="rtl"`, `lang="he"`)
- Heebo font from Google Fonts
- Accessibility: skip links, ARIA labels, semantic HTML
