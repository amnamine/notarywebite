"""
Tool A: Ijraa Ichhar converter.
Direct port of core.js:
- fixRunProps, mergeRuns, splitNumbers, wantDashes, addDashes, convertParagraph
- keepTable: scaling down to W=6698, grid re-matching, inline table styles borders/margins, fixed layout
- analyse, convert, and filling docx templates.
"""
import io
import os
import re
import zipfile
from typing import Dict, Any, List, Optional, Tuple

from core_engine.xml_helpers import (
    FONT, RF, NUM_PATTERN, LATIN_PATTERN,
    map_sz, strip_tags, decode_xml, esc_xml, plain_text,
    top_blocks, children, read_styles, style_props,
    clean_xml, order_rpr, read_fonts, explicit_fonts
)
from core_engine.detect_box import detect_box


def fix_run_props(x: str, st: Dict[str, Any], sp: Dict[str, Any]) -> str:
    """Normalise every run: Arial, mapped size, style bold/underline."""
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
        rpr = order_rpr(rpr)
        return '<w:r>' + rpr

    return re.sub(r'<w:r>(<w:rPr>[\s\S]*?</w:rPr>)?', repl, x)


def merge_runs(x: str) -> str:
    run_seq_re = re.compile(r'(?:<w:r>(?:<w:rPr>(?:(?!</w:rPr>)[\s\S])*</w:rPr>)?<w:t(?: xml:space="preserve")?>[^<]*</w:t></w:r>){2,}')
    single_run_re = re.compile(r'<w:r>(<w:rPr>(?:(?!</w:rPr>)[\s\S])*</w:rPr>)?<w:t(?: xml:space="preserve")?>([^<]*)</w:t></w:r>')

    def repl_seq(m: re.Match) -> str:
        seq = m.group(0)
        out = []
        cur = None
        
        def key(r: str) -> str:
            return re.sub(r'<w:lang [^>]*/>', '', r or '')

        for rm in single_run_re.finditer(seq):
            rpr = rm.group(1) or ''
            text = rm.group(2)
            if cur and key(cur[0]) == key(rpr):
                cur[1] += text
            else:
                if cur:
                    out.append(cur)
                cur = [rpr, text]
        if cur:
            out.append(cur)
        return ''.join(f'<w:r>{a}<w:t xml:space="preserve">{b}</w:t></w:r>' for a, b in out)

    return run_seq_re.sub(repl_seq, x)


def split_numbers(x: str) -> str:
    """Numbers like 91/148 or 2026/10/05 inside RTL runs -> own LTR run (shown as typed)."""
    single_run_re = re.compile(r'<w:r>(<w:rPr>(?:(?!</w:rPr>)[\s\S])*</w:rPr>)<w:t(?: xml:space="preserve")?>([^<]*)</w:t></w:r>')

    def repl_run(m: re.Match) -> str:
        rpr = m.group(1)
        t = m.group(2)
        if '<w:rtl/>' not in rpr or not NUM_PATTERN.search(t):
            return m.group(0)
        
        parts = re.split(f'({NUM_PATTERN.pattern})', t)
        out = []
        for p in parts:
            if not p:
                continue
            if NUM_PATTERN.fullmatch(p):
                ltr_rpr = rpr.replace('<w:rtl/>', '')
                out.append(f'<w:r>{ltr_rpr}<w:t xml:space="preserve">{p}</w:t></w:r>')
            else:
                out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{p}</w:t></w:r>')
        return ''.join(out)

    return single_run_re.sub(repl_run, x)


def want_dashes(text_trim: str, jc: str) -> bool:
    if not text_trim:
        return False
    if jc in ['center', 'start', 'end', 'right', 'left']:
        return False
    if re.search(r'-{3,}$', text_trim):
        return True
    if len(text_trim) < 60:
        return False
    return text_trim.endswith('.') or bool(re.search(r'(أدناه|ادناه|أدنـ+اه|يلي\s*:|التالية\s*:)$', text_trim))


def add_dashes(p: str, tab_pos: int) -> str:
    """Strip typed trailing dashes, then add a single tab with a hyphen leader to the end of the line."""
    runs = list(re.finditer(r'<w:r>(?:(?!<w:r>)[\s\S])*?</w:r>', p))
    last_rpr = ''
    
    for i in range(len(runs) - 1, -1, -1):
        rm = runs[i]
        r_text = rm.group(0)
        tm = re.search(r'(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)', r_text)
        if not tm:
            continue
        if not last_rpr:
            rp_match = re.search(r'<w:rPr>[\s\S]*?</w:rPr>', r_text)
            last_rpr = rp_match.group(0) if rp_match else ''
        
        if re.search(r'-\s*$', tm.group(2)):
            nt = re.sub(r'[\s\-]*-[\s\-]*$', '', tm.group(2))
            r2 = r_text[:tm.start()] + tm.group(1) + nt + tm.group(3) + r_text[tm.end():]
            p = p[:rm.start()] + r2 + p[rm.end():]
            if nt.strip():
                break
            continue
        break

    rp = last_rpr or '<w:rPr><w:rtl/></w:rPr>'
    if '<w:rtl/>' not in rp:
        rp = order_rpr(rp.replace('</w:rPr>', '<w:rtl/></w:rPr>'))
    
    last_p_close = p.rfind('</w:p>')
    p = p[:last_p_close] + f'<w:r>{rp}<w:tab/></w:r></w:p>'
    return p.replace('<w:pPr>', f'<w:pPr><w:tabs><w:tab w:val="left" w:leader="hyphen" w:pos="{tab_pos}"/></w:tabs>')


def convert_paragraph(x: str, st: Dict[str, Any], mode: str) -> str:
    x = clean_xml(x)
    pm = re.match(r'^<w:p(?: [^>]*)?>(<w:pPr>[\s\S]*?</w:pPr>)?', x)
    ppr = pm.group(1) or '' if pm else ''
    
    pstyle_m = re.search(r'<w:pStyle w:val="([^"]+)"', ppr)
    pstyle = pstyle_m.group(1) if pstyle_m else 'Normal'
    sp = style_props(st, pstyle)
    
    jc_m = re.search(r'<w:jc w:val="(\w+)"\/>', ppr)
    jc = jc_m.group(1) if jc_m else (sp.get('jc') or 'both')
    if re.search(r'kashida|distribute|thaiDistribute', jc, re.I):
        jc = 'both'
        
    mark_rpr_m = re.search(r'<w:rPr>[\s\S]*?</w:rPr>', ppr)
    mark_rpr = mark_rpr_m.group(0) if mark_rpr_m else ''
    if mark_rpr:
        mark_rpr = re.sub(r'<w:rFonts [^>]*/>|<w:rStyle [^>]*/>|<w:ins [^>]*/>|<w:del [^>]*/>', '', mark_rpr)
        mark_rpr = re.sub(r'<w:(sz|szCs) w:val="(\d+)"/>', lambda q: f'<w:{q.group(1)} w:val="{map_sz(int(q.group(2)))}"/>', mark_rpr)
        mark_rpr = order_rpr(mark_rpr)
        
    body = x[len(pm.group(0)):] if pm else x
    # dashes already added to this paragraph before: remove them
    body = re.sub(r'(?:<w:r>(?:<w:rPr>(?:(?!</w:rPr>)[\s\S])*</w:rPr>)?<w:tab/></w:r>\s*)+</w:p>$', '</w:p>', body)
    body = fix_run_props(body, st, sp)
    body = merge_runs(body)
    body = split_numbers(body)
    
    ind = 2552 if mode == 's556' else 0
    p = f'<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/><w:ind w:start="{ind}" w:end="0"/><w:jc w:val="{jc}"/>{mark_rpr}</w:pPr>' + body
    t = plain_text(p).rstrip()
    if want_dashes(t, jc):
        p = add_dashes(p, 13990 if mode == 's556' else 6690)
    return p


def keep_table(x: str, W: int, st: Dict[str, Any]) -> str:
    """Keep table exact layout, scale down only if wider than W (6698 twips)."""
    x = clean_xml(x)
    x = re.sub(r'<w:numPr>[\s\S]*?</w:numPr>', '', x)
    
    # borders/cell margins that came from the file's table style
    ts_m = re.search(r'<w:tblStyle w:val="([^"]+)"', x)
    if ts_m:
        id_ = ts_m.group(1)
        raw = ''
        g = 0
        while id_ and id_ in st['styles'] and g < 10:
            g += 1
            raw = st['styles'][id_]['raw'] + raw
            id_ = st['styles'][id_].get('based')
        
        bord_m = re.search(r'<w:tblBorders>[\s\S]*?</w:tblBorders>', raw)
        bord = bord_m.group(0) if bord_m else ''
        mar_m = re.search(r'<w:tblCellMar>[\s\S]*?</w:tblCellMar>', raw)
        mar = mar_m.group(0) if mar_m else ''

        def repl_tbl_pr(m: re.Match) -> str:
            inner = m.group(1)
            inner = re.sub(r'<w:tblStyle [^>]*/>', '', inner)
            ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
                     'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook']
            def add(name: str, el: str):
                nonlocal inner
                if not el or f'<w:{name}' in inner:
                    return
                k = ORDER.index(name)
                after = [inner.find(f'<w:{n}') for n in ORDER[k+1:] if inner.find(f'<w:{n}') >= 0]
                pos = min(after) if after else len(inner)
                inner = inner[:pos] + el + inner[pos:]
            add('tblBorders', bord)
            add('tblCellMar', mar)
            return '<w:tblPr>' + inner + '</w:tblPr>'

        x = re.sub(r'<w:tblPr>([\s\S]*?)</w:tblPr>', repl_tbl_pr, x, count=1)

    # paragraphs in cells: alignment / size that came from the file's styles
    def repl_cell_p(m: re.Match) -> str:
        ppr = m.group(1) or ''
        rest = m.group(2)
        style_m = re.search(r'<w:pStyle w:val="([^"]+)"', ppr)
        id_ = style_m.group(1) if style_m else 'Normal'
        sp = style_props(st, id_)
        p = ppr or '<w:pPr></w:pPr>'
        p = re.sub(r'<w:pStyle [^>]*/>', '', p)
        if sp.get('jc') and not re.search(r'<w:jc ', p):
            p = re.sub(r'(<w:rPr>|</w:pPr>)', f'<w:jc w:val="{sp["jc"]}"/>\\1', p, count=1)
        def_sz = sp.get('sz') or st.get('defaults', {}).get('sz')

        def repl_cell_r(q: re.Match) -> str:
            r = q.group(1) or '<w:rPr></w:rPr>'
            r = order_rpr(explicit_fonts(r, st['fonts']))
            if not def_sz:
                return '<w:r>' + r
            if not re.search(r'<w:sz ', r):
                r = order_rpr(r.replace('</w:rPr>', f'<w:sz w:val="{def_sz}"/><w:szCs w:val="{def_sz}"/></w:rPr>'))
            if sp.get('b') and not re.search(r'<w:b[ /]', r):
                r = order_rpr(r.replace('</w:rPr>', '<w:b/><w:bCs/></w:rPr>'))
            return '<w:r>' + r

        rest = re.sub(r'<w:r>(<w:rPr>[\s\S]*?</w:rPr>)?', repl_cell_r, rest)
        return '<w:p>' + p + rest + '</w:p>'

    x = re.sub(r'<w:p(?: [^>]*)?>(<w:pPr>[\s\S]*?</w:pPr>)?([\s\S]*?)</w:p>', repl_cell_p, x)

    # compute widths: grid sum, tblW and widest row
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
    # If grid doesn't match cells, take bestRow
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

    # Fixed tblLayout
    x = re.sub(r'<w:tblLayout [^>]*/>', '', x)
    def repl_fixed(m: re.Match) -> str:
        inner = m.group(1)
        after = [inner.find(f'<w:{n}') for n in ['tblCellMar', 'tblLook', 'tblCaption', 'tblDescription'] if inner.find(f'<w:{n}') >= 0]
        pos = min(after) if after else len(inner)
        return '<w:tblPr>' + inner[:pos] + '<w:tblLayout w:type="fixed"/>' + inner[pos:] + '</w:tblPr>'

    x = re.sub(r'<w:tblPr>([\s\S]*?)</w:tblPr>', repl_fixed, x, count=1)
    x = re.sub(r'<w:tblInd w:w="-?\d+" w:type="dxa"/>', '<w:tblInd w:w="0" w:type="dxa"/>', x)
    x = re.sub(r'<w:tblpPr [^>]*/>', '', x)
    return x


def analyse_docx(zip_bytes: bytes) -> Dict[str, Any]:
    with zipfile.ZipFile(io.BytesIO(zip_bytes), 'r') as z:
        if 'word/document.xml' not in z.namelist():
            raise ValueError('NOT_DOCX')
        doc = z.read('word/document.xml').decode('utf-8')
        styles_xml = z.read('word/styles.xml').decode('utf-8') if 'word/styles.xml' in z.namelist() else ''
        theme_file = next((n for n in z.namelist() if re.match(r'^word/theme/theme\d*\.xml$', n)), None)
        theme_xml = z.read(theme_file).decode('utf-8') if theme_file else ''

    b_start = doc.find('<w:body>') + 8
    sect_pos = doc.rfind('<w:sectPr')
    body_end = sect_pos if sect_pos > 0 else doc.rfind('</w:body>')
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


def build_content(a: Dict[str, Any], mode: str) -> Tuple[str, Dict[str, int]]:
    st = read_styles(a.get('styles_xml'))
    st['fonts'] = read_fonts(a.get('styles_xml'), a.get('theme_xml'))
    
    out = ''
    stats = {'paras': 0, 'tables': 0, 'rows': 0, 'dashes': 0}
    W = 14004 if mode == 's556' else 6698
    gap = '<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
    last_was_table = False

    for b in a['blocks']:
        if b['t'] == 'tbl':
            if last_was_table:
                out += gap
            out += keep_table(b['x'], W, st)
            stats['tables'] += 1
            stats['rows'] += len(re.findall(r'<w:tr[ >]', b['x']))
            last_was_table = True
            continue
        
        t = plain_text(b['x']).strip()
        if not t or re.match(r'^[\s\-_ـ.:]*$', t):
            continue
            
        p = convert_paragraph(b['x'], st, mode)
        if 'w:leader="hyphen"' in p:
            stats['dashes'] += 1
        out += p
        stats['paras'] += 1
        last_was_table = False

    return out, stats


def convert_ijraa(analysed: Dict[str, Any], template_dir: str, box_values: Dict[str, str]) -> Tuple[bytes, Dict[str, int]]:
    """Inject deed content and box values into templates/ijraa_ichhar_template."""
    end_p = '<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
    content_xml, stats = build_content(analysed, 'tpl')

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
    tdoc = tdoc.replace('{{IDA_HAJM}}', v(box_values.get('idaHajm'), '..................'))
    tdoc = tdoc.replace('{{IDA_RAQM}}', v(box_values.get('idaRaqm'), '..................'))
    tdoc = tdoc.replace('{{FI}}', v(box_values.get('fi'), '..............................................'))
    tdoc = tdoc.replace('{{MUJALLAD}}', v(box_values.get('mujallad'), '...................'))
    tdoc = tdoc.replace('{{MUJ_RAQM}}', v(box_values.get('mujRaqm'), '...................'))
    tdoc = tdoc.replace('{{RASM}}', v(box_values.get('rasm'), ''))

    # Zip files back together
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
