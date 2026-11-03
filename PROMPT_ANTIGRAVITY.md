# Prompt for Google Antigravity — rebuild «إجــراء إشهار عقاري» as a Flask + SQLite website

Open this whole folder as the workspace in Antigravity, then paste everything below the line into the chat.

---

You are taking over a **finished, working** single-page tool used by an Algerian notary office («مكتب الأستاذة حاجي اسيا — موثقة»). It is called **«إجــراء إشهار عقاري»**.

The user drops a Word deed (`.doc` or `.docx`). The tool:
- reads it;
- removes the old page numbering written in the deed;
- puts the whole content, unchanged, inside the official «إجــراء إشهار عقاري» box template (framed box on every page, the boxes at the top of page 1, the line «تاريخ تحرير العقد: … ، الفهرس : …»);
- keeps tables inside the box;
- writes page labels «الصـفحـــة الأولى .../...», «الثانية .../...» in the page footer, and «ـ الصـفحـــة … والأخـيــرة ـ» on the last page;
- returns a `.docx` that opens in Word 2007 and newer.

Your job is to turn it into a real website with **HTML, CSS, JavaScript, Python Flask and SQLite**, **without changing what it does or the Word file it produces**.

## What is in this folder

`index.html` is the complete working tool in one file. Open it in Chrome to see it working (access code 2806). It is the **reference**: the new site must look and behave exactly like it.

`source/` holds the same code split into files:

| File | What it is | Rule |
|---|---|---|
| `core.js` | The converter: reads the deed and builds the output .docx. | Do not modify. |
| `doc2docx.js` | Reads old Word 97-2003 `.doc` files in the browser. | Do not modify. |
| `paginate.js` | Old helper, no longer called. | Keep, do not modify. |
| `shells/tpl2/` | The Word template, as unzipped .docx parts: frame lines in the headers, the boxes, styles, and the footer with the page-label fields. | Do not modify. |
| `page/head.html` | Title and all CSS. | — |
| `page/body_tpl.html` | The page HTML. | — |
| `page/app_tpl.js` | UI logic: file input, fields, generate, download. | — |
| `page/gate_tpl.js` | The current access-code screen. | — |
| `jszip.min.js` | JSZip 3.10.1. | — |

At runtime the template is embedded in the page as JSON:

```
<script type="application/json" id="shells">{"tpl": {"<path inside the docx>": "<file text>", ...}}</script>
```

You may instead serve it as a static JSON file and load it with `fetch` before conversion, as long as the object passed to the converter is identical.

`index.html` is assembled in this order:
1. `<!doctype html><html lang="ar"><head><meta charset="utf-8"><meta name="viewport" …>`
2. `head.html`, then `</head><body>`
3. `body_tpl.html`
4. the shells JSON
5. `jszip`, `core.js`, `doc2docx.js`, `paginate.js`, `app_tpl.js`, `gate_tpl.js`
6. `</body></html>`

## Hard rules

1. **The output .docx must stay byte-for-byte identical in its content.** Do NOT rewrite, refactor, "clean up", reformat, minify, translate or "improve" `core.js`, `doc2docx.js` or anything in `shells/tpl2`. Load them exactly as they are. They contain many precise details that break easily:
   - frame line positions and exact row heights;
   - table placement inside the box (left-to-right and right-to-left tables, row-level centring removed);
   - table header font scaling;
   - bold handling for `.doc` files;
   - Word 2007 compatibility;
   - the footer page-label fields.
2. **Deed files never leave the user's computer.** All conversion stays in the browser (JavaScript).
   - Never upload deed files or their text to the server.
   - Never store them in SQLite.
   - Never log them.
   - The server only serves the site and handles accounts and usage records (see below).
3. **The UI stays the same**: Arabic, right-to-left, same texts, same fields, same colors and layout as `index.html`. No other languages.
4. **No external calls at runtime.** Serve JSZip and the fonts (Amiri, IBM Plex Sans Arabic, IBM Plex Mono) locally from `static/`. No CDNs, analytics or trackers.
5. Must work on old office PCs (Windows, Chrome/Edge). The produced file must open in **Word 2007** and newer.

## What to build

### Flask app

- `app.py` (Flask object named `app`) and `requirements.txt` (Flask, Werkzeug, nothing heavy).
- `templates/index.html`: the tool page, built from `source/page/*.html`.
- `static/`: `core.js`, `doc2docx.js`, `paginate.js`, `app_tpl.js` and `jszip.min.js` as separate files, loaded in the same order as above, plus fonts and the template JSON.
- **Remove `gate_tpl.js`** (the 2806 code screen). Replace it with real login (below). The tool page is only reachable when logged in.
- Secret key from the environment variable `SECRET_KEY`, with a generated fallback for local use.
- Secure session cookies, CSRF protection on forms, and security headers: CSP allowing only self, `X-Frame-Options: DENY`, `Referrer-Policy: no-referrer`.

### SQLite database (`instance/app.db`, created automatically)

- `users`: id, username (unique), password_hash (Werkzeug `generate_password_hash`), full_name, role (`admin` or `user`), active, created_at.
- `usage_log`: id, user_id, created_at, original_file_name, page_count_estimate (nullable). **Only metadata. Never file contents.** The browser sends a small JSON POST to `/api/log` after each successful generation.

### Pages

- `/login` and `/logout`. The login screen keeps the same look as the current code screen (office name «مكتب الأستاذة حاجي اسيا», Arabic, RTL).
- `/`: the tool (login required).
- `/admin` (admin only):
  - add and disable users, reset passwords;
  - see the usage log (who generated which file name and when), with a date filter.
  - All in Arabic, same style.
- A CLI command `flask create-admin` to create the first admin from the terminal.

### Run and deploy

- `start.bat` for Windows: creates a venv the first time, installs requirements, runs on `http://127.0.0.1:5000` and opens the browser.
- `README.md` in Arabic and English:
  - running locally on the office PC;
  - deploying online (PythonAnywhere step by step, and a VPS with gunicorn + nginx + HTTPS);
  - backing up `instance/app.db`.
- Keep a script `build_single_file.py` that still produces the one-file offline `index.html` from `source/` (USB / offline backup version).

## Acceptance test (do this before saying it is done)

1. Open the reference `index.html` (code 2806) and the new site (logged in) side by side in Chrome.
2. Load the same deed in both, a `.doc` and a `.docx`, leave the fields empty, generate and download.
3. Unzip both `.docx` files and compare `word/document.xml`, `word/header1.xml`, `word/header2.xml` and `word/footer1.xml`. They must be identical. If they differ, you changed something you must not change: revert it.
4. Open the result in Word and check:
   - the frame on every page;
   - the boxes at the top of page 1;
   - bold text kept;
   - tables inside the box;
   - the page labels at the bottom of each page;
   - «… والأخـيــرة ـ» on the last page.
5. Check that `instance/app.db` contains only users and usage metadata, never deed text.

Ask me before changing anything not listed above.
