# BizDeadlines 📅

**Business & Management Conference Submission Deadlines**

A static webapp inspired by [aideadlines.org](http://aideadlines.org) — live countdown timers, topic filters, and ABS ranking filters for 50+ business and management conferences.

🔗 **Live site**: `https://YOUR-USERNAME.github.io/biz-deadlines/`

---

## Features

- ⏱ **Live countdown timers** — days until each submission deadline
- 🎯 **Topic filters** — Accounting, Finance, Marketing, Operations, Strategy, IS, HR, and more
- 🏆 **ABS Ranking filters** — filter by A*, A, B
- 🔍 **Search** — full-text search across conference names and topics
- 📅 **Sort** — by deadline, conference date, or name
- 🔗 **Shareable URLs** — filters persist in the URL, so you can share exact views
- 📱 **Responsive** — works on mobile

---

## Conferences included

Covers ~50 major conferences tied to journals on the ABS 2026 list, including:

| Conference | Topics | ABS Journals |
|---|---|---|
| AOM Annual Meeting | Management, OB, Strategy, HR, ENT | A*, A, B |
| AFA Annual Meeting | Finance | A* |
| ICIS | Information Systems | A*, A, B |
| INFORMS | Operations, DS, SCM | A*, A, B |
| SMS Annual Conference | Strategy | A*, A |
| AAA Annual Meeting | Accounting | A*, A, B |
| AIB Annual Meeting | International Business | A*, A, B |
| POMS | Operations | A*, A |
| FMA | Finance | A |
| … and many more | | |

---

## Deploying to GitHub Pages

1. **Fork or create a new repo** on GitHub
2. **Upload these two files**:
   - `index.html`
   - `data.js`
3. Go to **Settings → Pages**
4. Set source to **Deploy from a branch → `main` → `/` (root)**
5. Your site will be live at `https://YOUR-USERNAME.github.io/REPO-NAME/`

---

## Adding / updating conferences

Edit `data.js`. Each conference entry looks like:

```js
{
  id: "aom2026",                          // unique id
  short: "AOM 2026",                      // short display name
  name: "Academy of Management Annual Meeting", // full name
  url: "https://aom.org/events/annual-meeting", // conference website
  deadline: "2026-01-13",                 // submission deadline (YYYY-MM-DD)
  conference_start: "2026-07-31",         // conference start (YYYY-MM-DD)
  conference_end: "2026-08-04",           // conference end (YYYY-MM-DD)
  location: "Philadelphia, PA, USA",      // location string
  topics: ["MGT","OB","ENT","STR","HR"],  // topic codes (see TOPICS in data.js)
  ranking: ["A*","A","B"],                // which ABS ranks publish here
  note: "Submission center opens early December" // optional note
}
```

**Topic codes** available: `ACC`, `COMM`, `DS`, `ECON`, `ENT`, `FIN`, `HCI`, `HIST`, `HR`, `IB`, `INS`, `IS`, `MGT`, `MKT`, `OB`, `OPS`, `PA`, `PSY`, `RE`, `SCM`, `SME`, `STR`, `TECH`

---

## URL parameters

Share filtered views via URL:
- `?topics=FIN,ACC` — filter by topics
- `?ranks=A*,A` — filter by ranking
- `?sort=conference` — sort by conference date
- `?q=ICIS` — pre-fill search
- `?past=1` — show past deadlines

Example: `index.html?ranks=A*&topics=FIN,ACC`

---

## Tech stack

Pure HTML + CSS + vanilla JS. No build step, no dependencies, no server. Just two files.

---

## Data sources

Deadline data sourced from official conference websites, cross-referenced with the ABS 2026 journal list. Deadlines are **typical** monthly estimates — always verify on the official conference website before submitting.

**Journals cross-referenced**: ABS 2026 list (256 journals, A*/A/B ratings)
