Author: kard

# TED Overview Hub

This folder is a static dashboard for the JSON acquisition output.

## Required Data

Place or copy the generated JSON output folder here:

- `JSON_Output/_manifest.json`
- `JSON_Output/Agent_2/_indexes/<yyMMdd>.json`
- `JSON_Output/Agent_2_Results/_indexes/<yyMMdd>.json`
- `JSON_Output/IMS_USA/_indexes/<yyMMdd>.json`

The weekday runner copies the project-level `JSON_Output` folder into this
dashboard folder before publishing.

The dashboard loads daily indexes on demand based on the selected date range.
The legacy `_index.json` files are retained as a fallback.

## Local Test

From this folder run a simple static server:

```powershell
python -m http.server 8000
```

Then open:

- `http://localhost:8000`

## Publishing

1. Run the JSON scraper pipeline.
2. Copy or commit `JSON_Output/` with this dashboard.
3. GitHub Pages redeploys automatically after push.

## Supabase Setup

1. Create a Supabase project.
2. In Supabase SQL Editor run:
   - `supabase/schema.sql`
3. Copy:
   - `supabase-config.example.js` -> `supabase-config.js`
4. Insert your values in `supabase-config.js`:
   - `url`
   - `anonKey`

Without Supabase config, the dashboard still loads JSON data, but auth,
comments, verification, and overrides stay disabled.
