"""
تطبيق مكتبي حديث (Tkinter GUI v2) — مكتب الأستاذة حاجي آسيا، موثقة
إجراء إشهار عقاري: معالجة ملفات Word ووضعها داخل الإطار الرسمي تلقائياً دون الحاجة لإنترنت.
"""
import sys
import os
import io
import re
import json
import subprocess
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Any, List, Optional, Tuple

# Resolve base directories
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    APP_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = BASE_DIR

# Check templates directory
TEMPLATES_DIR = os.path.join(BASE_DIR, 'docx_templates', 'ijraa_ichhar_template')
if not os.path.exists(TEMPLATES_DIR):
    alt_tpl = os.path.join(os.path.dirname(BASE_DIR), 'docx_templates', 'ijraa_ichhar_template')
    if os.path.exists(alt_tpl):
        TEMPLATES_DIR = alt_tpl

# ----------------- XML & DOCX CORE -----------------
FONT = 'Arial'
RF = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
SZ = {20: 20, 22: 22, 24: 24, 26: 24, 28: 26, 32: 28, 36: 30, 40: 32, 44: 34, 48: 36}
NUM_PATTERN = re.compile(r'\d+(?:\s?[\/\-]\s?\d+)+')
RPR_ORDER = [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
    'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
    'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
    'specVanish', 'oMath'
]

def map_sz(v: int) -> int:
    return SZ.get(v, 36 if v > 48 else v)

def strip_tags(x: str) -> str:
    return re.sub(r'<[^>]+>', '', x)

def decode_xml(s: str) -> str:
    return s.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')

def esc_xml(s: str) -> str:
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def plain_text(x: str) -> str:
    without_tabs = re.sub(r'<w:tab/>', ' ', x)
    without_brs = re.sub(r'<w:br[^>]*/>', ' ', without_tabs)
    return decode_xml(strip_tags(without_brs))

def top_blocks(body: str) -> List[Dict[str, str]]:
    out = []
    re_block = re.compile(r'<(/?)w:(tbl|p)(?=[ >/])[^>]*?(/?)>')
    depth = 0
    st = 0
    for m in re_block.finditer(body):
        close = m.group(1)
        tag = m.group(2)
        self_closing = m.group(3)
        if tag == 'tbl':
            if not close:
                if depth == 0:
                    st = m.start()
                depth += 1
            else:
                depth -= 1
                if depth == 0:
                    out.append({'t': 'tbl', 'x': body[st:m.end()]})
        elif depth == 0:
            if not close and self_closing:
                out.append({'t': 'p', 'x': m.group(0)})
            elif not close:
                st = m.start()
            else:
                out.append({'t': 'p', 'x': body[st:m.end()]})
    return out

def children(x: str, tag: str) -> List[str]:
    re_tag = re.compile(f'<w:tbl>|</w:tbl>|<w:{tag}(?=[ >])[^>]*>|</w:{tag}>')
    out = []
    tdepth = 0
    st = -1
    for m in re_tag.finditer(x):
        s = m.group(0)
        if s == '<w:tbl>':
            tdepth += 1
        elif s == '</w:tbl>':
            tdepth -= 1
        elif tdepth == 0:
            if s.startswith('</'):
                if st >= 0:
                    out.append(x[st:m.end()])
                    st = -1
            else:
                st = m.start()
    return out

def detect_box(text: str) -> Dict[str, str]:
    def uniq(arr: List[str]) -> List[str]:
        seen = set()
        res = []
        for x in arr:
            if x not in seen:
                seen.add(x)
                res.append(x)
        return res

    qism_matches = re.finditer(r'(?:^|[\s(])(?:ال)?قسم\s*(?:رقم)?\s*[:：]?\s*(\d{1,4})(?!\d)', text)
    qism_raw = [str(int(m.group(1))).zfill(2) for m in qism_matches]
    qism = uniq(qism_raw)

    maj_matches = re.finditer(r'مجموع[ةه]\s+(?:ال)?ملكي[ةه]\s*(?:رقم)?\s*[:：]?\s*(\d{1,5})', text)
    maj = uniq([m.group(1) for m in maj_matches])

    def freq(pattern: str, stop: re.Pattern) -> str:
        c = {}
        for m in re.finditer(pattern, text):
            words = []
            extracted = m.group(1)
            for w in re.split(r'\s+', extracted):
                if not w or stop.search(w):
                    break
                words.append(w)
                if len(words) >= 2:
                    break
            k = ' '.join(words)
            k = re.sub(r'^و(?=ال)', '', k)
            start_pos = max(0, m.start() - 250)
            end_pos = min(len(text), m.start() + 250)
            window = text[start_pos:end_pos]
            near = bool(re.search(r'قسم|مجموع[ةه] (?:ال)?ملكي|مسح|المسمى|واقع', window))
            if k:
                c[k] = c.get(k, 0) + (5 if near else 1)
        if not c:
            return ''
        sorted_c = sorted(c.items(), key=lambda item: item[1], reverse=True)
        return sorted_c[0][0]

    STOP_BAL = re.compile(r'^(بتاري|دائرة|ولاي|وولاي|الكائن|رقم|قسم|حي|مسجل|المؤرخ|الموافق)')
    STOP_WIL = re.compile(r'^(بتاري|دائرة|بلدي|الكائن|رقم|قسم|حي|المؤرخ|تحت|المسجل)')

    bal = freq(r'بلدي[ةه]\s*[:：]?\s*([^،,.:()\-0-9\n]{2,60})', STOP_BAL)
    wil = freq(r'ولاي[ةه]\s*[:：]?\s*([^،,.:()\-0-9\n]{2,40})', STOP_WIL)

    fihris_m = re.search(r'فـ*هـ*ـ*ر[ـ]*س\s*(?:ر[ـ]*ق[ـ]*م)?\s*[:：]?\s*(\d+\s*/\s*\d+)', text)
    fihris = re.sub(r'\s+', '', fihris_m.group(1)) if fihris_m else ''

    return {
        'qism': ' و '.join(qism[:6]),
        'majmoua': '، '.join(maj[:20]),
        'fihris': fihris,
        'baladiya': bal,
        'wilaya': wil,
    }

def fix_run_props(x: str, st: Dict[str, Any], sp: Dict[str, Any]) -> str:
    def repl(m: re.Match) -> str:
        rpr = m.group(1) or '<w:rPr></w:rPr>'
        rpr = re.sub(r'<w:rFonts [^>]*/>', '', rpr)
        rpr = re.sub(r'<w:rStyle [^>]*/>', '', rpr)
        if sp.get('b') and not re.search(r'<w:b[ /]', rpr):
            rpr = rpr.replace('<w:rPr>', '<w:rPr><w:b/><w:bCs/>')
        if sp.get('u') and not re.search(r'<w:u ', rpr):
            rpr = rpr.replace('</w:rPr>', '<w:u w:val="single"/></w:rPr>')
        
        def_sz = sp.get('sz') or st.get('defaults', {}).get('sz') or 28
        if not re.search(r'<w:sz ', rpr):
            rpr = rpr.replace('</w:rPr>', f'<w:sz w:val="{def_sz}"/><w:szCs w:val="{def_sz}"/></w:rPr>')
        elif not re.search(r'<w:szCs ', rpr):
            rpr = re.sub(r'(<w:sz w:val="(\d+)"/>)', r'\1<w:szCs w:val="\2"/>', rpr)
        
        rpr = re.sub(r'<w:(sz|szCs) w:val="(\d+)"/>', lambda q: f'<w:{q.group(1)} w:val="{map_sz(int(q.group(2)))}"/>', rpr)
        rpr = rpr.replace('<w:rPr>', '<w:rPr>' + RF)
        
        # Order rPr
        props = {}
        for pm in re.finditer(r'<w:([A-Za-z0-9]+)([^>]*?)(?:/>|>([\s\S]*?)</w:\1>)', rpr):
            props[pm.group(1)] = pm.group(0)
        ordered = '<w:rPr>' + ''.join(props[k] for k in RPR_ORDER if k in props) + '</w:rPr>'
        return '<w:r>' + ordered

    return re.sub(r'<w:r>(<w:rPr>[\s\S]*?</w:rPr>)?', repl, x)

def merge_runs(x: str) -> str:
    run_seq_re = re.compile(r'(?:<w:r>(?:<w:rPr>(?:(?!</w:rPr>)[\s\S])*</w:rPr>)?<w:t(?: xml:space="preserve")?>[^<]*</w:t></w:r>){2,}')
    single_run_re = re.compile(r'<w:r>(<w:rPr>(?:(?!</w:rPr>)[\s\S])*</w:rPr>)?<w:t(?: xml:space="preserve")?>([^<]*)</w:t></w:r>')

    def repl(m: re.Match) -> str:
        seq = m.group(0)
        runs = list(single_run_re.finditer(seq))
        if not runs:
            return seq
        out = []
        cur_rpr = runs[0].group(1) or ''
        cur_t = runs[0].group(2)
        for r in runs[1:]:
            rpr = r.group(1) or ''
            t = r.group(2)
            if rpr == cur_rpr:
                cur_t += t
            else:
                sp = ' xml:space="preserve"' if (cur_t.startswith(' ') or cur_t.endswith(' ')) else ''
                out.append(f'<w:r>{cur_rpr}<w:t{sp}>{cur_t}</w:t></w:r>')
                cur_rpr = rpr
                cur_t = t
        sp = ' xml:space="preserve"' if (cur_t.startswith(' ') or cur_t.endswith(' ')) else ''
        out.append(f'<w:r>{cur_rpr}<w:t{sp}>{cur_t}</w:t></w:r>')
        return ''.join(out)

    return run_seq_re.sub(repl, x)

def split_numbers(x: str) -> str:
    def repl(m: re.Match) -> str:
        rpr = m.group(1) or ''
        t = m.group(2)
        if not re.search(r'<w:rtl[ /]', rpr) or not NUM_PATTERN.search(t):
            return m.group(0)
        rpr_no_rtl = re.sub(r'<w:rtl[ /][^>]*>', '', rpr)
        pieces = []
        last = 0
        for nm in NUM_PATTERN.finditer(t):
            st_pos = nm.start()
            if st_pos > last:
                chunk = t[last:st_pos]
                sp = ' xml:space="preserve"' if (chunk.startswith(' ') or chunk.endswith(' ')) else ''
                pieces.append(f'<w:r>{rpr}<w:t{sp}>{chunk}</w:t></w:r>')
            num_chunk = nm.group(0)
            pieces.append(f'<w:r>{rpr_no_rtl}<w:t xml:space="preserve">{num_chunk}</w:t></w:r>')
            last = nm.end()
        if last < len(t):
            chunk = t[last:]
            sp = ' xml:space="preserve"' if (chunk.startswith(' ') or chunk.endswith(' ')) else ''
            pieces.append(f'<w:r>{rpr}<w:t{sp}>{chunk}</w:t></w:r>')
        return ''.join(pieces)

    return re.sub(r'<w:r>(<w:rPr>[\s\S]*?</w:rPr>)?<w:t(?: xml:space="preserve")?>([^<]*)</w:t></w:r>', repl, x)

def want_dashes(text: str, jc: str) -> bool:
    if jc == 'center':
        return False
    t = text.strip()
    if not t or re.search(r'[\-_ـ]{4,}$', t):
        return False
    if len(t) < 60:
        return False
    return bool(re.search(r'(?:[.؟!:؛]|أدناه|أعلاه|التالي|كما يلي|ما يلي)[\s)]*$', t))

def add_dashes(p_xml: str, leader_pos: int) -> str:
    p_xml = re.sub(r'[\s\-_ـ]+(?=</w:t></w:r></w:p>)', '', p_xml)
    p_xml = re.sub(r'<w:r><w:t>[\s\-_ـ]+</w:t></w:r></w:p>$', '</w:p>', p_xml)
    p_xml = re.sub(r'<w:tab/>', '', p_xml)
    tabs_xml = f'<w:tabs><w:tab w:val="right" w:leader="hyphen" w:pos="{leader_pos}"/></w:tabs>'
    
    if '<w:tabs>' in p_xml:
        p_xml = re.sub(r'<w:tabs>[\s\S]*?</w:tabs>', tabs_xml, p_xml)
    elif '</w:pPr>' in p_xml:
        p_xml = p_xml.replace('</w:pPr>', tabs_xml + '</w:pPr>')
    else:
        p_xml = re.sub(r'<w:p[ >]', lambda m: m.group(0) + '<w:pPr>' + tabs_xml + '</w:pPr>', p_xml, count=1)
        
    dash_run = f'<w:r>{RF}<w:tab/></w:r>'
    return p_xml[:-6] + dash_run + '</w:p>'

def convert_paragraph(p_xml: str, st: Dict[str, Any], mode: str) -> str:
    p_xml = re.sub(r'<w:pStyle [^>]*/>', '', p_xml)
    p_xml = re.sub(r'<w:pBdr>[\s\S]*?</w:pBdr>', '', p_xml)
    p_xml = re.sub(r'<w:shd [^>]*/>', '', p_xml)
    p_xml = re.sub(r'<w:(ind|spacing|rFonts|jc) [^>]*/>', '', p_xml)

    t = plain_text(p_xml)
    jc_val = 'center' if re.search(r'^\s*-(?:-|\s)*[0-9]+\s*-(?:-|\s)*\s*$', t) else 'both'
    
    p_xml = fix_run_props(p_xml, st, {})
    p_xml = merge_runs(p_xml)
    p_xml = split_numbers(p_xml)

    ind_w = '14004' if mode == 's556' else '6698'
    p_xml = re.sub(
        r'<w:pPr>|<w:p>',
        f'<w:p><w:pPr><w:bidi/><w:jc w:val="{jc_val}"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:start="0" w:end="0"/>',
        p_xml,
        count=1
    )
    if want_dashes(t, jc_val):
        p_xml = add_dashes(p_xml, 6690 if mode != 's556' else 14000)
    return p_xml

def keep_table(x: str, W: int, st: Dict[str, Any]) -> str:
    grid_block_m = re.search(r'<w:tblGrid>[\s\S]*?</w:tblGrid>', x)
    grid_block = grid_block_m.group(0) if grid_block_m else ''
    grid = [int(m.group(1)) for m in re.finditer(r'<w:gridCol w:w="(\d+)"', grid_block)]

    tbl_inner = x[x.find('>')+1:x.rfind('</w:tbl>')]
    outer_rows = children(tbl_inner, 'tr')
    row_max = 0
    best_row = None
    for r in outer_rows:
        row_content = r[r.find('>')+1:r.rfind('</w:tr>')]
        cs = children(row_content, 'tc')
        ws = []
        for c in cs:
            wm = re.search(r'<w:tcW (?:w:w="(\d+)" w:type="dxa"|w:type="dxa" w:w="(\d+)")', c)
            val = int(wm.group(1) or wm.group(2)) if wm else 0
            ws.append(val)
        row_sum = sum(ws)
        if row_sum > row_max:
            row_max = row_sum
        if not best_row and len(cs) == len(grid) and all(w > 0 for w in ws) and not any('<w:gridSpan' in c for c in cs):
            best_row = ws

    gsum = sum(grid)
    if best_row and gsum > 0 and abs(sum(best_row) - gsum) > gsum * 0.05:
        grid = best_row
        new_grid_xml = '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{g}"/>' for g in grid) + '</w:tblGrid>'
        x = x.replace(grid_block, new_grid_xml)

    tbl_w_m = re.search(r'<w:tblW (?:w:w="(\d+)" w:type="dxa"|w:type="dxa" w:w="(\d+)")', x[:x.find('</w:tblPr>')])
    tbl_w_val = int(tbl_w_m.group(1) or tbl_w_m.group(2)) if tbl_w_m else 0
    tw = max(sum(grid), tbl_w_val, row_max)
    f = (W / tw) if tw > W else 1.0

    if f < 1.0:
        def sc(v: Any) -> int:
            return max(1, round(int(v) * f))
        x = re.sub(r'<w:gridCol w:w="(\d+)"', lambda m: f'<w:gridCol w:w="{sc(m.group(1))}"/>', x)
        x = re.sub(r'<w:(tcW|tblW|wBefore|wAfter) w:w="(\d+)" w:type="dxa"/>', lambda m: f'<w:{m.group(1)} w:w="{sc(m.group(2))}" w:type="dxa"/>', x)
        x = re.sub(r'<w:(tcW|tblW|wBefore|wAfter) w:type="dxa" w:w="(\d+)"/>', lambda m: f'<w:{m.group(1)} w:type="dxa" w:w="{sc(m.group(2))}"/>', x)
        x = re.sub(r'<w:(sz|szCs) w:val="(\d+)"/>', lambda m: f'<w:{m.group(1)} w:val="{max(14, round(int(m.group(2)) * f))}"/>', x)
        x = re.sub(r'(<w:(?:tcMar|tblCellMar)>[\s\S]*?</w:(?:tcMar|tblCellMar)>)', lambda m: re.sub(r'w:w="(\d+)"', lambda q: f'w:w="{sc(q.group(1))}"', m.group(1)), x)

    x = re.sub(r'<w:tblLayout [^>]*/>', '', x)
    def repl_fixed(m: re.Match) -> str:
        inner = m.group(1)
        after = [inner.find(f'<w:{n}') for n in ['tblCellMar', 'tblLook', 'tblCaption', 'tblDescription'] if inner.find(f'<w:{n}') >= 0]
        pos = min(after) if after else len(inner)
        return '<w:tblPr>' + inner[:pos] + '<w:tblLayout w:type="fixed"/>' + inner[pos:] + '</w:tblPr>'

    x = re.sub(r'<w:tblPr>([\s\S]*?)</w:tblPr>', repl_fixed, x, count=1)
    x = re.sub(r'<w:tblInd w:w="-?\d+" w:type="dxa"/>', '<w:tblInd w:w="0" w:type="dxa"/>', x)
    return x

def analyse_docx(zip_bytes: bytes) -> Dict[str, Any]:
    try:
        z = zipfile.ZipFile(io.BytesIO(zip_bytes), 'r')
    except Exception as e:
        raise ValueError(f"الملف ليس بصيغة Word (.docx) صالحة: {e}")

    with z:
        namelist = z.namelist()
        if 'word/document.xml' not in namelist:
            raise ValueError('الملف لا يحتوي على word/document.xml، تأكد من سلامة ملف Word.')
        
        doc = z.read('word/document.xml').decode('utf-8', errors='replace')
        styles_xml = z.read('word/styles.xml').decode('utf-8', errors='replace') if 'word/styles.xml' in namelist else ''
        theme_file = next((n for n in namelist if re.match(r'^word/theme/theme\d*\.xml$', n)), None)
        theme_xml = z.read(theme_file).decode('utf-8', errors='replace') if theme_file else ''

    b_idx = doc.find('<w:body>')
    if b_idx >= 0:
        b_start = b_idx + 8
    else:
        b_idx_tag = doc.find('<w:body')
        b_start = doc.find('>', b_idx_tag) + 1 if b_idx_tag >= 0 else 0

    sect_pos = doc.rfind('<w:sectPr')
    body_end_idx = doc.rfind('</w:body>')
    if sect_pos > b_start:
        body_end = sect_pos
    elif body_end_idx > b_start:
        body_end = body_end_idx
    else:
        body_end = len(doc)

    body = doc[b_start:body_end]
    boxed = False

    bh = list(re.finditer(r'<w:trHeight w:val="(2000|9300)"', doc))
    if bh and 'إطار مخصص للمحافظ' in strip_tags(doc):
        row_i = bh[0].start()
        tc1 = doc.find('<w:tc>', row_i)
        tc2 = doc.find('<w:tc>', tc1 + 5) if tc1 >= 0 else -1
        if tc2 > 0:
            tc_pr_end = doc.find('</w:tcPr>', tc2)
            s = tc_pr_end + 9 if tc_pr_end >= 0 else -1
            e = doc.rfind('</w:tc></w:tr></w:tbl>')
            if s > 0 and e > s:
                body = doc[s:e]
                boxed = True

    blocks = top_blocks(body)
    text = '\n'.join(plain_text(b['x']) for b in blocks)
    return {
        'doc': doc,
        'styles_xml': styles_xml,
        'theme_xml': theme_xml,
        'blocks': blocks,
        'boxed': boxed,
        'box': detect_box(text),
        'text': text
    }

def convert_ijraa(analysed: Dict[str, Any], template_dir: str, box_values: Dict[str, str]) -> Tuple[bytes, Dict[str, int]]:
    end_p = '<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
    content_xml, stats = '', {'paras': 0, 'tables': 0, 'rows': 0, 'dashes': 0}
    st = {}

    for b in analysed['blocks']:
        if b['t'] == 'tbl':
            content_xml += keep_table(b['x'], 6698, st)
            stats['tables'] += 1
            stats['rows'] += len(re.findall(r'<w:tr[ >]', b['x']))
            continue
        t = plain_text(b['x']).strip()
        if not t or re.match(r'^[\s\-_ـ.:]*$', t):
            continue
        p = convert_paragraph(b['x'], st, 'tpl')
        if 'w:leader="hyphen"' in p:
            stats['dashes'] += 1
        content_xml += p
        stats['paras'] += 1

    # Date and Fihris line at the beginning
    dt = (box_values.get('date') or '').strip() or '..................'
    fh = (box_values.get('fihris') or '').strip() or '..................'
    header_line = f"تاريخ تحرير العقد: {dt} ، الفهرس : {fh}."
    header_p = f'<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="60" w:line="276" w:lineRule="auto"/><w:ind w:start="0" w:end="0"/><w:jc w:val="start"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr><w:t xml:space="preserve">{esc_xml(header_line)}</w:t></w:r></w:p>'
    content_xml = header_p + content_xml

    doc_xml_path = os.path.join(template_dir, 'word', 'document.xml')
    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        tdoc = f.read()

    def v(s: Optional[str], d: str) -> str:
        s_clean = (s or '').strip()
        return esc_xml(s_clean) if s_clean else d

    tdoc = tdoc.replace('{{CONTENT}}', content_xml + end_p)
    tdoc = tdoc.replace('{{QISM}}', v(box_values.get('qism'), ''))
    tdoc = tdoc.replace('{{BALADIYA}}', v(box_values.get('baladiya'), ''))
    tdoc = tdoc.replace('{{MAJMOUA}}', v(box_values.get('majmoua'), ''))
    tdoc = tdoc.replace('{{HISSA}}', v(box_values.get('hissa'), ''))
    tdoc = tdoc.replace('{{WILAYA}}', v(box_values.get('wilaya'), '..............'))
    tdoc = tdoc.replace('{{MUHAFIZ}}', v(box_values.get('wilaya'), ''))
    tdoc = tdoc.replace('{{IDA_HAJM}}', v(box_values.get('idaHajm'), '..................'))
    tdoc = tdoc.replace('{{IDA_RAQM}}', v(box_values.get('idaRaqm'), '..................'))
    tdoc = tdoc.replace('{{FI}}', v(box_values.get('fi'), '..............................................'))
    tdoc = tdoc.replace('{{MUJALLAD}}', v(box_values.get('mujallad'), '...................'))
    tdoc = tdoc.replace('{{MUJ_RAQM}}', v(box_values.get('mujRaqm'), '...................'))
    tdoc = tdoc.replace('{{RASM}}', v(box_values.get('rasm'), ''))

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for root_dir, _, files in os.walk(template_dir):
            for file in files:
                abs_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(abs_path, template_dir).replace('\\', '/')
                if rel_path == 'word/document.xml':
                    zout.writestr(rel_path, tdoc.encode('utf-8'))
                else:
                    with open(abs_path, 'rb') as fp:
                        zout.writestr(rel_path, fp.read())

    return zip_buffer.getvalue(), stats


# ----------------- TKINTER MODERN V2 GUI -----------------
class NotaryAppV2(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("مكتب الأستاذة حاجي آسيا — موثقة — إجراء إشهار عقاري (V2)")
        self.geometry("860x720")
        self.minsize(740, 620)
        self.configure(bg="#f3f4f6")

        self.current_file_path = None
        self.analysed_data = None

        self.setup_styles()
        self.build_ui()

    def setup_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        
        self.c_band = "#16213a"
        self.c_accent = "#0e6b62"
        self.c_accent_hover = "#128277"
        self.c_paper = "#f3f4f6"
        self.c_sheet = "#ffffff"
        self.c_ink = "#151c2b"
        self.c_muted = "#5d6676"
        self.c_line = "#dfe3ea"

    def build_ui(self):
        # 1. Header Band
        band = tk.Frame(self, bg=self.c_band, height=76)
        band.pack(fill='x', side='top')
        band.pack_propagate(False)

        header_container = tk.Frame(band, bg=self.c_band)
        header_container.pack(expand=True, fill='both', padx=24, pady=10)

        seal_canvas = tk.Canvas(header_container, width=50, height=50, bg=self.c_band, highlightthickness=0)
        seal_canvas.pack(side='right', padx=(14, 0))
        seal_canvas.create_oval(3, 3, 47, 47, outline='#eef1f7', width=1.5)
        seal_canvas.create_oval(7, 7, 43, 43, outline='#eef1f7', width=0.8)
        seal_canvas.create_text(25, 26, text="ح.آ", fill='#eef1f7', font=('Traditional Arabic', 14, 'bold'))

        title_box = tk.Frame(header_container, bg=self.c_band)
        title_box.pack(side='right', anchor='center')
        lbl_title = tk.Label(title_box, text="مكتب الأستاذة حاجي آسيا — موثقة", bg=self.c_band, fg="#ffffff", font=('Segoe UI', 13, 'bold'))
        lbl_title.pack(anchor='e')
        lbl_sub = tk.Label(title_box, text="أداة تحضير العقود الرسمية: «إجراء إشهار عقاري» (V2)", bg=self.c_band, fg="#a9b3c7", font=('Segoe UI', 9))
        lbl_sub.pack(anchor='e')

        # 2. Main Scrollable Container
        main_canvas = tk.Canvas(self, bg=self.c_paper, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        scrollable_frame = tk.Frame(main_canvas, bg=self.c_paper)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=820)
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True, padx=16, pady=12)
        scrollbar.pack(side="right", fill="y")

        # Step 1: File selection
        f1 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=20, pady=16)
        f1.pack(fill='x', pady=6)

        lbl_s1 = tk.Label(f1, text="1. ملف العقد (Word)", bg=self.c_sheet, fg=self.c_accent, font=('Segoe UI', 11, 'bold'))
        lbl_s1.pack(anchor='e')

        self.btn_select = tk.Button(f1, text="📂 اختيار ملف العقد (.docx أو .doc)", bg=self.c_paper, fg=self.c_ink,
                                    font=('Segoe UI', 10, 'bold'), relief='solid', bd=1, padx=16, pady=10, cursor='hand2',
                                    command=self.browse_file)
        self.btn_select.pack(fill='x', pady=10)

        self.lbl_file = tk.Label(f1, text="لم يتم اختيار أي ملف بعد", bg=self.c_sheet, fg=self.c_muted, font=('Segoe UI', 9))
        self.lbl_file.pack(anchor='e')

        self.lbl_stats = tk.Label(f1, text="", bg=self.c_sheet, fg=self.c_accent, font=('Segoe UI', 9, 'bold'))
        self.lbl_stats.pack(anchor='e', pady=(4, 0))

        # Step 2: Contract info & Cadastral box
        f2 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=20, pady=16)
        f2.pack(fill='x', pady=6)

        lbl_s2 = tk.Label(f2, text="2. معلومات العقد ومراجع مسح الأراضي", bg=self.c_sheet, fg=self.c_accent, font=('Segoe UI', 11, 'bold'))
        lbl_s2.pack(anchor='e')

        lbl_s2_hint = tk.Label(f2, text="تُملأ تلقائياً من نص العقد. تأكد منها وصحّحها قبل الإنشاء.", bg=self.c_sheet, fg=self.c_muted, font=('Segoe UI', 9))
        lbl_s2_hint.pack(anchor='e')

        grid2 = tk.Frame(f2, bg=self.c_sheet)
        grid2.pack(fill='x', pady=10)

        self.fields = {}
        contract_cadastral_fields = [
            ("تاريخ تحرير العقد", "date", "مثال: 2026/09/02"),
            ("الفهرس", "fihris", "مثال: 2026/619"),
            ("قسم", "qism", "مثال: 27"),
            ("بلدية", "baladiya", "مثال: وادي العنب"),
            ("مجموعة ملكية رقم", "majmoua", "مثال: 155"),
            ("حصة رقم", "hissa", "فارغة إن لم توجد"),
            ("الولاية (الحفظ العقاري)", "wilaya", "مثال: عنابة")
        ]
        for idx, (label_txt, key, pl_holder) in enumerate(contract_cadastral_fields):
            r = idx // 2
            c = (idx % 2) * 2
            lbl = tk.Label(grid2, text=label_txt, bg=self.c_sheet, fg=self.c_muted, font=('Segoe UI', 9))
            lbl.grid(row=r*2, column=c+1, sticky='e', padx=8, pady=(4, 0))
            entry = tk.Entry(grid2, font=('Segoe UI', 10), justify='right', bd=1, relief='solid')
            entry.grid(row=r*2+1, column=c+1, sticky='ew', padx=8, pady=(0, 6))
            self.fields[key] = entry
        grid2.columnconfigure(1, weight=1)
        grid2.columnconfigure(3, weight=1)

        # Step 3: Deposit & Fees
        f3 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=20, pady=16)
        f3.pack(fill='x', pady=6)

        lbl_s3 = tk.Label(f3, text="3. الإيداع والرسم (اختياري)", bg=self.c_sheet, fg=self.c_accent, font=('Segoe UI', 11, 'bold'))
        lbl_s3.pack(anchor='e')

        lbl_s3_hint = tk.Label(f3, text="ما تتركه فارغاً يبقى نقاطاً «......» لتُكتب باليد في الملف النهائي.", bg=self.c_sheet, fg=self.c_muted, font=('Segoe UI', 8))
        lbl_s3_hint.pack(anchor='e')

        grid3 = tk.Frame(f3, bg=self.c_sheet)
        grid3.pack(fill='x', pady=10)

        deposit_fields = [
            ("إيداع — حجم", "idaHajm"),
            ("إيداع — رقم", "idaRaqm"),
            ("في", "fi"),
            ("مجلد", "mujallad"),
            ("رقم المجلد", "mujRaqm"),
            ("رسم", "rasm")
        ]
        for idx, (label_txt, key) in enumerate(deposit_fields):
            r = idx // 3
            c = (idx % 3) * 2
            lbl = tk.Label(grid3, text=label_txt, bg=self.c_sheet, fg=self.c_muted, font=('Segoe UI', 9))
            lbl.grid(row=r*2, column=c+1, sticky='e', padx=6, pady=(4, 0))
            entry = tk.Entry(grid3, font=('Segoe UI', 10), justify='right', bd=1, relief='solid')
            entry.grid(row=r*2+1, column=c+1, sticky='ew', padx=6, pady=(0, 6))
            self.fields[key] = entry
        grid3.columnconfigure(1, weight=1)
        grid3.columnconfigure(3, weight=1)
        grid3.columnconfigure(5, weight=1)

        # Step 4: Action & Generation
        f4 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=20, pady=18)
        f4.pack(fill='x', pady=6)

        self.btn_convert = tk.Button(f4, text="⚡ إنشاء ملف «إجراء إشهار عقاري» الجديد", bg="#1b6a3a", fg="#ffffff",
                                     font=('Segoe UI', 11, 'bold'), relief='flat', padx=18, pady=12, cursor='hand2',
                                     state='disabled', command=self.do_convert)
        self.btn_convert.pack(fill='x', pady=4)

        self.lbl_result = tk.Label(f4, text="", bg=self.c_sheet, fg="#1b6a3a", font=('Segoe UI', 9, 'bold'))
        self.lbl_result.pack(pady=4)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="اختيار ملف العقد Word",
            filetypes=[
                ("ملفات Word (*.docx;*.doc)", "*.docx;*.doc"),
                ("مستند Word (.docx)", "*.docx"),
                ("مستند Word قديم (.doc)", "*.doc"),
                ("كافة الملفات (*.*)", "*.*")
            ]
        )
        if not file_path:
            return

        try:
            if not os.path.exists(file_path):
                messagebox.showerror("خطأ", f"تعذر العثور على الملف المحدد:\n{file_path}")
                return

            if file_path.lower().endswith('.doc'):
                messagebox.showinfo(
                    "ملف .doc قديم",
                    "هذا ملف Word قديم (.doc).\nلقراءته بدقة ومراعاة التنسيق، يرجى فتحه في Word واختيار «حفظ باسم» بصيغة .docx، أو استخدام واجهة المتصفح المخصصة."
                )
                return

            with open(file_path, 'rb') as fp:
                raw_bytes = fp.read()

            if not raw_bytes:
                messagebox.showerror("خطأ", "الملف المختار فارغ (0 بايت).")
                return

            analysed = analyse_docx(raw_bytes)
            self.current_file_path = file_path
            self.analysed_data = analysed

            fname = os.path.basename(file_path)
            self.lbl_file.config(text=f"الملف المختار: {fname}", fg=self.c_ink)

            n_p = sum(1 for b in analysed['blocks'] if b['t'] == 'p' and b['x'].strip())
            n_tbl = sum(1 for b in analysed['blocks'] if b['t'] == 'tbl')
            boxed_txt = " | الملف داخل الإطار أصلاً" if analysed['boxed'] else ""
            self.lbl_stats.config(text=f"عدد الفقرات: {n_p}  |  عدد الجداول: {n_tbl}{boxed_txt}")

            # Fill extracted cadastral values
            box = analysed.get('box', {})
            for key in ['qism', 'baladiya', 'majmoua', 'wilaya', 'fihris']:
                if key in self.fields and key in box:
                    self.fields[key].delete(0, tk.END)
                    self.fields[key].insert(0, box[key])

            self.btn_convert.config(state='normal')
            self.lbl_result.config(text="")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ في فحص الملف", f"تعذر فتح وقراءة ملف Word:\n{str(e)}")

    def do_convert(self):
        if not self.analysed_data or not self.current_file_path:
            return

        if not os.path.exists(TEMPLATES_DIR):
            messagebox.showerror("خطأ في القالب", f"لم يتم العثور على مجلد القالب الرسمي في:\n{TEMPLATES_DIR}")
            return

        base_name = os.path.splitext(os.path.basename(self.current_file_path))[0]
        default_out = f"{base_name} - إجراء إشهار عقاري.docx"
        
        save_path = filedialog.asksaveasfilename(
            title="حفظ الملف المحول",
            initialfile=default_out,
            defaultextension=".docx",
            filetypes=[("Word Document (.docx)", "*.docx")]
        )
        if not save_path:
            return

        try:
            box_values = {key: entry.get().strip() for key, entry in self.fields.items()}
            out_bytes, stats = convert_ijraa(self.analysed_data, TEMPLATES_DIR, box_values)

            with open(save_path, 'wb') as f_out:
                f_out.write(out_bytes)

            msg = f"تم بنجاح! تم إنشاء الملف:\n{os.path.basename(save_path)}\n\nالإحصائيات:\n• فقرات: {stats['paras']}\n• جداول: {stats['tables']} ({stats['rows']} سطر)\n• فواصل: {stats['dashes']}"
            self.lbl_result.config(text="✓ تم إنشاء الملف بنجاح.")
            messagebox.showinfo("اكتمل التحويل", msg)
        except Exception as e:
            messagebox.showerror("خطأ أثناء التحويل", f"حدث خطأ أثناء إنشاء ملف Word:\n{str(e)}")


if __name__ == '__main__':
    app = NotaryAppV2()
    app.mainloop()
