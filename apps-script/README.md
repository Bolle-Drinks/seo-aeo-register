# Worklist backend

`Code.gs` turns the "BOLLE SEO & AEO worklist" spreadsheet into a tiny API so the
register page can save ticks and names for everyone, instead of only in the
viewer's own browser.

## Deploy (about five minutes, all in the browser)

1. Open the worklist spreadsheet.
2. **Extensions → Apps Script**. Delete whatever is in `Code.gs`.
3. Paste in the contents of `Code.gs` from this folder. Save.
4. **Deploy → New deployment**. Click the gear, choose **Web app**.
   - *Execute as:* **Me**
   - *Who has access:* **Anyone**
5. Click Deploy and authorise it when Google asks.
6. Copy the **Web app URL** — it ends in `/exec`.

Send that URL back and the page gets wired to it.

## Why "Anyone" access

The page is served from GitHub Pages with no login, so the browser calling this
endpoint is anonymous. Writes are gated by a shared token in `Code.gs` rather
than by a Google sign-in. Anyone who has both the page URL and reads its source
could write to the sheet — the sheet keeps full revision history, so a bad edit
is undoable, and nothing else in the Drive is reachable through this script.

If that trade is not acceptable, keep editing the sheet directly and leave the
page as a read-only mirror.
