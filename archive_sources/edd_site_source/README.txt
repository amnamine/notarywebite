Source of the current web page (single-file version that ran inside Claude).

frontend/core.js          - "إجراء إشهار عقاري": reads a .docx, puts its body into the box template (keeps tables as they are, scales them to fit), text rules (no empty lines, end-of-paragraph dash leaders, slash-numbers left-to-right), fills the cadastral box.
frontend/core556.js       - "عقد جدول الوصفي للتقسيم": finds scanned table pictures (docx) / PDF pages, builds the prompt sent to the AI, slices pages, rebuilds tables with merged cells (rowspan/colspan), total-row checks, A4 landscape output.
frontend/app.js           - page logic (upload, fields, progress, downloads).
frontend/gate.js          - entry code screen (SHA-256 of "djebari-edd:" + code) and tabs.
frontend/head.html, body.html - design (CSS tokens, Arabic RTL layout).
frontend/current_full_page.html - the whole page as published (everything inlined).
templates/ijraa_ichhar_template/ - unzipped .docx of the box template with placeholders:
    {{CONTENT}} {{QISM}} {{BALADIYA}} {{MAJMOUA}} {{HISSA}} {{WILAYA}} {{IDA_HAJM}} {{IDA_RAQM}} {{FI}} {{MUJALLAD}} {{MUJ_RAQM}} {{RASM}}
    header1.xml / header2.xml draw the box lines down to the bottom of every page.
templates/landscape_556_template/ - unzipped .docx, A4 landscape with page numbers, placeholder {{BODY}}.
