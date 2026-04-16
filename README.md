# TED Overview Hub (GitHub Pages)

This folder is a static web app for GitHub Pages.

## Required file

Place your workbook here:

- `data/ted_results.xlsx`

The app reads these sheets from that file:

- `Agent_2`
- `Agent_2_Results`

## Local test

From `overview_hub_web` run a simple static server (example):

```powershell
python -m http.server 8000
```

Then open:

- `http://localhost:8000`

## Publish on GitHub Pages

1. Commit and push `overview_hub_web/` including `data/ted_results.xlsx`.
2. In GitHub repo settings, enable **Pages**.
3. Choose branch/folder that contains this app (for example `main` + `/260204_Aquise_Python/overview_hub_web`).
4. Wait for deployment and open the Pages URL.

## Updating data and redeploying

1. Replace `data/ted_results.xlsx` with the newest export.
2. Commit and push.
3. GitHub Pages redeploys automatically.
