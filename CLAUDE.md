# CLAUDE.md — Noura Quality Management Platform

This document describes the codebase structure, conventions, and workflows for AI assistants working on this repository.

---

## Project Overview

**Noura Quality Management Platform** is a single-file Arabic-language web application for quality and staff performance tracking. It supports data entry, analytics dashboards, and reporting for a driving school environment.

- **Architecture:** Single-file SPA (no build step, no framework)
- **Language:** Vanilla JavaScript (ES6+), HTML5, CSS3
- **Backend:** Supabase (PostgreSQL cloud database)
- **UI Language:** Arabic (RTL layout)
- **Entry Point:** `index.html` (~6,800 lines, contains all HTML/CSS/JS)

---

## Repository Structure

```
/
├── index.html        # Entire application (HTML + embedded CSS + JS)
├── README.md         # Copy of the HTML file (naming artifact — treat as html)
└── CLAUDE.md         # This file
```

There are no subdirectories, build tools, package managers, or test files.

---

## Technology Stack

| Layer        | Technology                              |
|--------------|-----------------------------------------|
| UI Framework | None (vanilla HTML/CSS/JS)              |
| Language     | JavaScript ES6+                         |
| Database     | Supabase (PostgreSQL via JS SDK v2)     |
| Font         | Cairo (Google Fonts, Arabic support)    |
| Export       | SheetJS XLSX v0.20.2 (CDN)             |
| Theme        | CSS custom properties (light/dark)      |
| Storage      | localStorage (theme, attendance)        |

All external libraries are loaded from CDN — no `npm install` or local install needed.

---

## Application Configuration

### Supabase (embedded in index.html)
```javascript
const SB_URL = 'https://ljgctajxnqzkwxxhyrqi.supabase.co';
const SB_KEY = '<anon key>';
```

### Hardcoded Users
Authentication is client-side only. Users and passwords are defined in the `USERS` object inside `index.html`. Roles are `'admin'` or `'staff'`.

### Quality Target
```javascript
const TARGET = 90;  // Threshold percentage for pass/fail color coding
```

---

## Database Tables (Supabase)

| Table              | Purpose                                     |
|--------------------|---------------------------------------------|
| `violations`       | Training violations (women's/men's sections)|
| `surveys`          | Customer satisfaction surveys (1–5 stars)   |
| `facility_reports` | Facility/maintenance issue reports          |
| `building_reports` | Building crowding and wait-time observations|

Data is loaded once with `loadAll()` into global arrays `allV`, `allS`, `allF`, `allB`, then filtered client-side.

---

## Global State Variables

```javascript
let currentUser = null;   // Logged-in user object {name, role, ...}
let pendingSec  = null;   // Section waiting for authentication
let curSec      = '';     // Currently active section
let allV        = [];     // All violation records
let allS        = [];     // All survey records
let allF        = [];     // All facility report records
let allB        = [];     // All building report records
```

---

## Page & Section Structure

### Pages (controlled by `showPage(name, el)`)
| Page      | Description                        |
|-----------|------------------------------------|
| `home`    | Dashboard with widgets + cards     |
| `records` | Data tables + statistics (admin)   |

### Sections (controlled by `openSec(type)` / `showSection(t)`)
| Section    | Description                              | Access   |
|------------|------------------------------------------|----------|
| `female`   | Women's training violation form          | Staff+   |
| `male`     | Men's training violation form            | Staff+   |
| `survey`   | Customer satisfaction survey form        | Staff+   |
| `facility` | Facility issue reporting (with images)   | Staff+   |
| `building` | Building crowding observation form       | Staff+   |

### Admin-Only Sections (`adminOnly(section)`)
| Section      | Description                        |
|--------------|------------------------------------|
| `ops`        | Operations/staff performance board |
| `attendance` | Attendance tracking                |
| `rating`     | Rating analysis                    |
| `archive`    | Data archive + export              |
| `sales`      | Sales dashboard (coming soon)      |
| `incentives` | Incentives management (coming soon)|

---

## Key Functions Reference

### Authentication & Navigation
- `doLogin()` — Validate credentials from `USERS`, set `currentUser`
- `doLogout()` — Clear session, return to home
- `showPage(name, el)` — Switch between `home` and `records`
- `openSec(type)` — Open a section (prompts login if unauthenticated)
- `showSection(t)` — Navigate within a section

### Data Loading
- `loadAll()` — Fetch all 4 Supabase tables and populate global arrays
- `setDT()` — Initialize datetime picker to current time
- `updateCoworkerList()` — Populate observer dropdowns with staff names

### Form Submissions
- `submitViolation()` — Submit a training violation record
- `submitSurvey()` — Submit a customer satisfaction survey
- `submitFacility()` — Submit a facility issue report
- `submitBuilding()` — Submit a building observation

### Rendering
- `renderV(recs)` — Render violations table and stats
- `renderS(surveys)` — Render survey analytics
- `renderF(recs)` — Render facility reports table
- `renderB(recs)` — Render building reports table
- `renderHomeWidgets()` — Render home dashboard summary cards
- `renderActiveStaff()` — Render per-staff performance widgets

### Analytics & Utilities
- `updateVStats()` — Compute violation pass/fail percentages
- `countBy(arr, field)` — Group-count utility
- `getQ(d)` — Return quarter label from date string
- `barChart(entries, total, colors)` — Generate HTML bar chart markup
- `splitObservers(raw)` — Parse comma-separated observer name strings

### Export
- `getExportRows(section)` — Prepare filtered rows for export
- `dlCSV(rows, fname)` — Download as CSV
- `downloadXLSX(rows, fname)` — Download as Excel (via SheetJS)
- `exportSurvey()` — Export survey analytics report

---

## Naming Conventions

| Type              | Convention        | Example                          |
|-------------------|-------------------|----------------------------------|
| JS functions      | camelCase         | `submitViolation()`, `loadAll()` |
| CSS classes       | kebab-case        | `.nav-tab`, `.submit-btn`        |
| Global data arrays| `all` prefix      | `allV`, `allS`, `allF`, `allB`   |
| DB tables         | snake_case        | `facility_reports`               |
| localStorage keys | `nq_` prefix      | `nq_theme`                       |

---

## Styling Conventions

- All styles are embedded inside `<style>` in `index.html`
- Theme colors are defined as CSS custom properties on `:root`
- Dark mode is toggled via `data-theme="dark"` on `<html>` and stored in `localStorage` as `nq_theme`
- Layout uses CSS Grid and Flexbox
- RTL direction: `<html lang="ar" dir="rtl">`
- Responsive breakpoints: 375px, 414px, 768px, 1024px, 1400px

### Key CSS Variables
```css
--primary: #00A896      /* Teal — primary brand color */
--secondary: #1A2E35    /* Dark blue */
--success: #2E9E6B
--warning: #E07B2A
--error: #D94F4F
```

---

## Development Workflow

### Running the App
Open `index.html` directly in a browser — no build step or server required. For Supabase features, an internet connection is needed.

### Making Changes
1. Edit `index.html` directly (all HTML, CSS, and JS are in this file)
2. Test in browser by refreshing the page
3. Commit and push to the designated branch

### No Build Process
- No `npm`, `yarn`, `webpack`, `vite`, or similar tools
- No TypeScript compilation
- No test runner

---

## Security Notes

> **Important:** This application uses client-side-only authentication and exposes credentials in source code. Do not add production secrets or real personal data without implementing proper server-side authentication.

Known security concerns:
1. **Hardcoded passwords** in the `USERS` object (visible in browser dev tools)
2. **Supabase anon key** exposed in client code (normal for anon key; ensure RLS is enabled)
3. **No input sanitization** visible in current code
4. **Client-side role checks** only — admin restrictions can be bypassed in browser

When modifying authentication or access control, consider these limitations and do not increase the attack surface further without a proper security review.

---

## Editing Guidelines for AI Assistants

1. **Read `index.html` before editing** — the file is large; locate the relevant section before making changes.
2. **Preserve Arabic text** — UI labels and messages are in Arabic. Do not translate or alter them unless explicitly requested.
3. **Do not add build tooling** unless explicitly asked — this is intentionally a no-build project.
4. **Keep all code in `index.html`** — do not split into separate files unless that refactoring is the goal of the task.
5. **Match existing style** — use camelCase for JS, kebab-case for CSS classes, and follow the patterns above.
6. **No new dependencies** unless asked — additional libraries should be loaded via CDN only, consistent with the current approach.
7. **Test by inspection** — there is no test suite; validate logic changes by tracing through the relevant functions manually or in a browser.
8. **Respect RTL layout** — any new UI elements must work correctly in right-to-left context.
