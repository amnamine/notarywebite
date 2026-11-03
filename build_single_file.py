"""
Offline single-file HTML builder for «إجراء إشهار عقاري».
Assembles the complete tool into a single standalone index.html
that can run directly in the browser offline from a USB stick.
"""
import os
import json

def build():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    claude_dir = os.path.join(base_dir, 'archive_sources', 'claude_artifact')
    out_file = os.path.join(base_dir, 'dist_executables', 'web_offline_single_file', 'index_offline_single_file.html')
    os.makedirs(os.path.dirname(out_file), exist_ok=True)

    with open(os.path.join(claude_dir, 'page', 'head.html'), 'r', encoding='utf-8') as f:
        head_html = f.read()

    with open(os.path.join(claude_dir, 'page', 'body_tpl.html'), 'r', encoding='utf-8') as f:
        body_html = f.read()

    with open(os.path.join(base_dir, 'static', 'shells.json'), 'r', encoding='utf-8') as f:
        shells_json = f.read()

    with open(os.path.join(claude_dir, 'jszip.min.js'), 'r', encoding='utf-8') as f:
        jszip_js = f.read()

    with open(os.path.join(claude_dir, 'core.js'), 'r', encoding='utf-8') as f:
        core_js = f.read()

    with open(os.path.join(claude_dir, 'doc2docx.js'), 'r', encoding='utf-8') as f:
        doc2docx_js = f.read()

    with open(os.path.join(claude_dir, 'paginate.js'), 'r', encoding='utf-8') as f:
        paginate_js = f.read()

    with open(os.path.join(claude_dir, 'page', 'app_tpl.js'), 'r', encoding='utf-8') as f:
        app_js = f.read()

    with open(os.path.join(claude_dir, 'page', 'gate_tpl.js'), 'r', encoding='utf-8') as f:
        gate_js = f.read()

    html = f"""<!doctype html>
<html lang="ar">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
{head_html}
</head>
<body>
{body_html}
<script type="application/json" id="shells">{shells_json}</script>
<script>{jszip_js}</script>
<script>{core_js}</script>
<script>{doc2docx_js}</script>
<script>{paginate_js}</script>
<script>{app_js}</script>
<script>{gate_js}</script>
</body>
</html>"""

    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Built offline single file successfully: {out_file} ({len(html)} bytes)")

if __name__ == '__main__':
    build()
