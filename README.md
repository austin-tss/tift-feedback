# Tift County Schools – Community Feedback
Static GitHub Pages front end (`/` form, `/dashboard/`) + Google Apps Script backend (`Code.gs`, bound to a Google Sheet).
Secrets (`DASHBOARD_CODE`, `GEMINI_API_KEY`) live only in Apps Script Script Properties — never in this repo.
Setup: create a Google Sheet → Extensions > Apps Script → paste Code.gs → set Script Properties → Deploy as Web app (Execute as Me, Anyone) → paste /exec URL into `config.js`.
Email confirmations are stubbed; see EMAIL_TEMPLATE.md (draft, not sent).
