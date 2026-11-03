"""
AI Reader and PDF processing for Tool B.
- PyMuPDF rendering at ~300 DPI (~2600px width)
- PDF machine text layer extraction
- Slicing: whole page + 2 to 4 zoomed horizontal bands with ~4% overlap
- Anthropic Claude vision call using exact prompt from core556.js
"""
import io
import json
import math
import os
import re
from typing import Dict, Any, List, Optional, Tuple
from PIL import Image

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import anthropic
except ImportError:
    anthropic = None


def render_pdf_page(pdf_bytes: bytes, page_num: int) -> Tuple[bytes, str]:
    """Render one PDF page at ~300 DPI (approx 2600px wide) and extract text layer."""
    if not fitz:
        raise RuntimeError("PyMuPDF is not installed")
    
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc.load_page(page_num)
    
    # Calculate scale so width is ~2600 px or scale ~4.0
    rect = page.rect
    w1 = rect.width
    h1 = rect.height
    scale = min(2600.0 / w1, 4200.0 / h1, 4.5)
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_bytes = pix.tobytes("jpeg")
    
    text = page.get_text("text") or ""
    text_clean = re.sub(r'\s+', ' ', text).strip()
    pdf_text = text_clean[:6000] if len(text_clean) > 30 else ""
    doc.close()
    return img_bytes, pdf_text


def create_image_slices(img_bytes: bytes, max_slices: int = 5) -> List[bytes]:
    """Create whole page image and 2-4 zoomed horizontal bands with 4% overlap."""
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    w, h = img.size
    
    out_slices = []
    # 1. Whole image
    buf_whole = io.BytesIO()
    img.save(buf_whole, format="JPEG", quality=92)
    out_slices.append(buf_whole.getvalue())
    
    # Determine number of bands
    n = math.ceil((w * h) / 1.3e6)
    n = max(0, min(n, 4, max_slices - 1))
    if n < 2:
        return out_slices
        
    ov = round(h * 0.04)
    step = math.ceil(h / n)
    for i in range(n):
        y0 = max(0, i * step - ov)
        y1 = min(h, (i + 1) * step + ov)
        band = img.crop((0, y0, w, y1))
        buf = io.BytesIO()
        band.save(buf, format="JPEG", quality=92)
        out_slices.append(buf.getvalue())
        
    return out_slices


def build_ai_prompt(prev_table: Optional[Dict[str, Any]], n_images: int, pdf_text: str = "") -> str:
    """Exact prompt from core556.js."""
    lines = [
        'You are copying a scanned page from an Algerian notarial deed (Arabic, right-to-left) into structured data, so it can be rebuilt as an editable Word page that looks exactly like the scan.',
        (f'Image 1 is the WHOLE page: use it for the layout (which cells are merged, where each table starts and ends). Images 2 to {n_images} are zoomed horizontal bands of the same page, top to bottom, overlapping a little: use them to read small text and digits exactly. Every row must appear ONCE.'
         if n_images > 1 else 'The image is the whole page.'),
        'Copy EXACTLY what is printed. Never correct, compute, translate, reorder, add or drop anything. Keep every row, including every total row (المجموع الجزئي, المجموع الكلي, المجموع...).',
        'Digits: Western 0-9, exactly as printed, keeping leading zeros ("03", "07"), decimal commas, slashes and dashes ("78,57", "1000/1", "F3-01", "C+9", "//", "—").',
        'If a character cannot be read with confidence write "?" in its place (e.g. "1?7"); never guess silently.',
        'Ignore stamps, seals, signatures, logos, handwriting and colour highlighting that are not part of the printed text.',
        '',
        'Return the page as an ordered list of blocks, top to bottom:',
        '- printed text line(s) outside tables: {"type":"text","text":"...","align":"right|center|left","bold":true|false,"underline":true|false}  (one block per line)',
        '- a table: {"type":"table","header_rows":N,"continues_previous":true|false,"rows":[[cell,...],...]}',
        'Table rules (like HTML tables):',
        '- Each row lists its cells from RIGHT to LEFT as they appear on the page.',
        '- A cell is {"t":"text"} plus "cs":k when it spans k columns and "rs":k when it spans k rows (omit when 1). Add "b":true if the cell text is bold.',
        '- A cell covered by a merged cell from a row above is NOT listed again in the lower rows (exactly like HTML rowspan).',
        '- A header written on two lines inside one cell: join the lines with "\\n".',
        '- "header_rows" = how many rows at the top are column headings (0 if the table starts without headings).',
        ('- "continues_previous": true only when this table has no heading rows because it continues a table from the previous page' +
         (f' (the previous table had {prev_table["cols"]} columns; its headings were {json.dumps(prev_table.get("head", []))}).' if prev_table else '.')),
        '',
        ('The PDF page also carries this machine text layer (exact characters, but not in table order). When it contains a value you see in the image, use its exact characters:\n' + pdf_text if pdf_text else ''),
        'Reply with ONLY this JSON: {"blocks":[...],"unreadable":false}. Set "unreadable": true only if nothing on the page can be read.'
    ]
    return '\n'.join(l for l in lines if l != '')


def call_claude_vision(api_key: str, model_name: str, prompt_text: str, image_slices: List[bytes]) -> Dict[str, Any]:
    """Call Anthropic Claude with the image slices and parse the JSON response."""
    if not anthropic:
        raise RuntimeError("anthropic package is not installed")
    import base64
    client = anthropic.Anthropic(api_key=api_key)
    
    content = []
    for img_bytes in image_slices:
        b64 = base64.b64encode(img_bytes).decode('ascii')
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": b64
            }
        })
    content.append({"type": "text", "text": prompt_text})

    response = client.messages.create(
        model=model_name or "claude-3-5-sonnet-latest",
        max_tokens=4096,
        messages=[{"role": "user", "content": content}]
    )
    
    resp_text = response.content[0].text
    # Extract JSON from code block or direct text
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', resp_text)
    raw_json = json_match.group(1) if json_match else resp_text.strip()
    return json.loads(raw_json)
