# Burmelin Order Management System

A complete order + inventory system. The **frontend** (customer order form + employee
dashboard) is static HTML hosted on **Vercel** so it works on every browser including
Safari. The **backend** is a **Google Apps Script** web app that reads/writes the
Google Sheet inventory and acts as a JSON API.

## Structure

```
web/                 ← THIS is what Vercel deploys (Root Directory = web)
  index.html         ← Customer order form        →  https://YOURSITE.vercel.app/
  employee.html      ← Employee dashboard          →  https://YOURSITE.vercel.app/employee
  vercel.json        ← clean URLs config

Code.gs              ← Apps Script backend (JSON API) — NOT served by Vercel
appsscript.json      ← Apps Script manifest
Customer.html        ← Apps Script HTML fallback (legacy)
Employee.html        ← Apps Script HTML fallback (legacy)
```

The frontend talks to the backend through the API URL embedded near the top of
`web/index.html` and `web/employee.html` (the `API_URL` constant). If you ever
redeploy the Apps Script to a brand-new URL, update `API_URL` in both files.

## Deploy the frontend (Vercel)

1. Push this repo to GitHub (private is fine).
2. On vercel.com → **Add New → Project** → import this repo.
3. **Important:** set **Root Directory** to `web`.
4. Deploy. Share the URL with customers; add `/employee` for the dashboard.

## Update the backend (Google Apps Script)

Backend lives in `Code.gs`. With [clasp](https://github.com/google/clasp):

```
clasp push --force
clasp deploy --deploymentId <DEPLOYMENT_ID>
```

## Notes

- Employee password is set in `Code.gs` (`EMPLOYEE_PASS`) and validated server-side.
- Inventory is read live from the **Updated Size Chart** sheet.
- Completed orders are written to the **Sent Orders** sheet; stock is deducted on
  **Mark as Done** only.
