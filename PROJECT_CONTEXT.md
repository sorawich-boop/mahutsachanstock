# Mahutsachan Cafe — Stock Tracker

## Project Overview
A web app for daily stock tracking at Mahutsachan Cafe (Khlong Thom). Replaces manual Google Sheets entry + LINE screenshot workflow.

## Files
| File | Purpose |
|------|---------|
| `stock-tracker-v2.html` | Main web app — open in any browser (iPad/PC) |
| `Code.gs` | Google Apps Script API — paste into script.google.com |

## Architecture
```
stock-tracker-v2.html  ←→  Apps Script Web App  ←→  Google Sheets
                                                  ←→  Google Drive (photos)
localStorage = cache only (server always wins on sync)
```

## Google IDs
```
BAR_SPREADSHEET_ID       = '1mnZZoNfg4MEmtwd-xzEzgma78OJpsDMasCA0KnQbaJY'
CONTAINER_SPREADSHEET_ID = '1BGBIfk-XAGO4TH41lT-WMK3_8BsDTw4HbGc5_9dcOv0'
BAR_DRIVE_FOLDER_ID      = '1kkjz8bN8SITKzcZAKQBoguorGNp3xFRH'
CONTAINER_DRIVE_FOLDER_ID= '1o5zVa-L0q75CkAzR53P_r4Vkl54EOOrN'
Apps Script Web App URL  = https://script.google.com/macros/s/AKfycbxACsuCdxqONMIFcm2BKEA1TWFtiSDU8rFswj9PH9f2f2evWniVC7Uf-Lox5iBvQmBW/exec
```

## Google Sheets Structure (new clean format)
Each spreadsheet has 2 tabs auto-created on first sync:

**Config tab** — item catalog
| name | unit | cat | min | order |

**Log tab** — every transaction
| date | by | item | recv | used | remaining | kind | photos | ts |

`kind` values: `baseline` | `entry` | `count` | `adjust`

## App Features
1. **Stock dashboard** — สต๊อกบาร์น้ำ + ภาชนะ tabs, low stock alerts, progress bars
2. **Daily entry (บันทึก)** — รับ/เบิก per item in sheet column order with col numbers, photo required for recv
3. **Physical count (นับจริง)** — actual vs system diff, resets baseline, writes count rows to Log
4. **Daily summary (รายวัน)** — full stock table per day, เซฟรูป downloads JPG for LINE
5. **History** — transaction log, delete reverses stock + writes adjust row to Sheets
6. **Add/Edit/Delete items** — pushes catalog changes to Config tab
7. **↺ รีเซ็ต** — wipes local cache and full re-pull from Sheets

## Photo Upload Flow
- Photo required when recv > 0 (photo zone slides open per item)
- Compressed to max 1400px JPEG before upload
- Drive folder structure: `parentFolder / dd-mm-yyyy / itemName / photo.jpg`
- Drive scope requires testDrive() run from editor to authorize

## Apps Script Setup (after any code change)
1. Paste Code.gs into script.google.com
2. Run `testDrive()` from editor → authorize Drive permissions
3. Deploy → **New deployment** → Web App → Execute as: Me → Anyone
4. Paste new URL into app config
5. Test: `[URL]?action=test` in incognito browser → should return JSON

## Known Issues / Notes
- Must use **incognito browser** to test URL if multiple Google accounts signed in
- After deploying, always use **New deployment** (not edit existing) or code changes don't apply
- `appsscript.json` must include Drive scope:
```json
{
  "timeZone": "Asia/Bangkok",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```

## Data Flow
### Adding stock (รับ)
1. User fills recv + attaches photo
2. Photo compressed → uploaded to Drive
3. Entry saved locally + POSTed to Apps Script
4. Apps Script appends rows to Log tab (one row per item)

### Sync
1. GET `?action=getStock` → returns Config catalog + Log current stock
2. Server always wins — local cache overwritten
3. If Log empty on first sync → pushes local state as `baseline` rows

### Delete history
1. Reverses stock locally
2. POSTs `adjustStock` → Apps Script writes compensating `adjust` rows to Log

## Item Catalog
38 bar items + 9 container items. Stored in Config tab after first sync.
Add/edit/delete via ⚙️ button in dashboard → immediately pushed to Config tab.
