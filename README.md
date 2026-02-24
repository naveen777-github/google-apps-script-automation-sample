# Google Apps Script Automation Sample (Sheets + REST + JSON)

A small Google Workspace automation demo built with **Google Apps Script** that imports data from a **REST API**, parses **JSON**, writes structured rows into **Google Sheets**, and generates a quick **summary dashboard** + **execution logs** for reliability.

<img width="1811" height="850" alt="Execution" src="https://github.com/user-attachments/assets/23d187df-a91b-4e59-9de9-361fe1c410aa" />

This prototype mirrors common business automation patterns: **config-driven runs**, **upsert (no duplicates)**, and **lightweight reporting** for decision-making.

---

## What it does

- Reads configuration from a `config` sheet (API URL, max pages, mode).
- Fetches paginated data from a REST API (UrlFetchApp).
- Validates + normalizes JSON fields into a structured table.
- Writes results into a `data` sheet.
- Supports:
  - **upsert** mode (update existing rows by `id`, insert new ones)
  - **append** mode (always append)
- Creates a `summary` sheet with quick insights:
  - total rows, inserted, updated, skipped
  - distinct types
  - top 5 types by count
- Writes execution info/errors into a `logs` sheet.

---

## Sheet Tabs

### 1) `config`

<img width="1920" height="1080" alt="Config" src="https://github.com/user-attachments/assets/098d5345-bba9-4e35-88aa-85c9bee5ec99" />

Key/value configuration used by the script:

| key       | value |
|----------|-------|
| api_url   | https://rickandmortyapi.com/api/location |
| max_pages | 3 |
| mode      | upsert |

**Notes**
- `max_pages`: how many pages to pull from the API
- `mode`: `upsert`

---

### 2) `data`

<img width="1920" height="1080" alt="Data" src="https://github.com/user-attachments/assets/e3a1a4e3-3a2d-4757-a6d0-933069e6a6b0" />

Output table:

`timestamp | id | name | type | dimension`

---

### 3) `summary`

<img width="1920" height="1080" alt="Summary" src="https://github.com/user-attachments/assets/808ddb13-2d2c-406f-b582-1aaaf88d9480" />



Auto-generated metrics after each run:

- Total rows in sheet
- Imported (new)
- Updated
- Skipped
- Distinct types
- Top type #1 ... #5

---

### 4) `logs`

<img width="1920" height="1080" alt="Logs" src="https://github.com/user-attachments/assets/59bb9ac5-363e-4183-b4d0-52b79642dc1f" />

Execution tracing for reliability:

`timestamp | level | message | context`

Examples:
- Starting import (shows config)
- Fetching page (shows URL/page)
- Import complete (shows inserted/updated/runtime)
- Import failed (shows error details)

---

## How to run 

1. Create a Google Sheet with 4 tabs: `config`, `data`, `summary`, `logs`
2. Open **Extensions → Apps Script**
3. Paste the script into `Code.gs` (or import the provided file contents)
4. Run `runImportFromConfig()` once to authorize permissions (UrlFetch)
5. Refresh the sheet
6. Use the menu:
   **Automation Sample → Run Import (Config)**

---

## Core functions 

- `onOpen()` → adds the custom menu
- `runImportFromConfig()` → main orchestrator
- `fetchLocationsPages_()` → REST API + pagination + JSON parsing
- `upsertRows_()` → upsert/append behavior into the `data` tab
- `writeSummary_()` → calculates summary metrics/top types
- `log_()` → writes structured logs to the `logs` tab

---

## Why this is useful
This shows the core skills needed for business automation:
- structured data design in Sheets
- API integration and JSON parsing
- workflow reliability via logging + error handling
- actionable reporting for operations
