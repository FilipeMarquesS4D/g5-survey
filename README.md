# Group 5 — City-Centre Carrier Perspective (Antwerp field survey)

A phone-friendly tool for collecting urban-freight delivery observations from the
**carrier's** point of view, with live per-member counts, photo uploads, and a
one-click "copy for AI analysis" export.

- **`index.html`** — the survey app (open it in any phone browser).
- **`backend.gs`** — Google Apps Script that receives uploads into a Google Sheet
  and stores photos in a Drive folder.

## Live link

Once GitHub Pages is enabled (Settings ▸ Pages ▸ Branch: `main`, folder `/root`),
the survey is served at:

```
https://<your-username>.github.io/g5-survey/
```

Share that link with the team — no files to pass around.

## Data backend (one person, ~3 min)

GitHub Pages only *serves* the page; the data is collected by a Google Sheet so
all four phones write to one place.

1. Create a Google Sheet → **Extensions ▸ Apps Script**.
2. Paste **`backend.gs`**, Save.
3. **Deploy ▸ New deployment ▸ Web app** → *Execute as: Me*, *Who has access: **Anyone*** → Deploy → authorise.
4. Copy the `/exec` URL.
5. Open the live survey → **⚙️ Shared upload setup** → paste the URL → **Save URL**
   (use **Test connection** to confirm).

The sheet header row, the **G5 Survey Photos** Drive folder, and de-duplication
are all created automatically on the first upload.

## How the team uses it

1. Open the link, pick your name.
2. One form per delivery stop → **Save & upload**. Works offline; queued saves
   upload automatically when back online.
3. The top dashboard shows live totals per member.
4. At the end, anyone taps **🤖 Copy AI summary** and pastes into Claude/ChatGPT,
   or pulls the full data straight from the Google Sheet.

## Members

Youssef Kheiredin · Carlos Filipe da Fonseca Nunes Marques · Thuy My Y Hoang · Theresa-Maria Mersini
