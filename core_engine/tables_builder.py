"""
Tool B: Tables 556 builder.
Direct port of core556.js:
- toGrid (rowspan/colspan occupancy grid)
- check (subtotal/grand total comparisons, yellow uncertain cells '?' warning)
- widths (content-based column width calculation)
- tableXml (bidiVisual, headers, F2F2F2 totals, FFF2CC yellow, m2)
- textPara (title paragraphs)
- build and buildFromPages (A4 landscape 16838x11906, margins 1417)
"""
import io
import math
import re
import zipfile
from typing import Dict, Any, List, Optional, Tuple, Set

from core_engine.xml_helpers import esc_xml, strip_tags, top_blocks

NUM_PATTERN_556 = re.compile(r'[\d?]*\d[\d?]*(?:[.,][\d?]+)*(?:\s?[\/\-]\s?[\d?]+(?:[.,][\d?]+)*)*')
LATIN_PATTERN_556 = re.compile(r'[A-Za-z][A-Za-z0-9.]*(?:\s?-\s?[0-9]+)?')


def to_grid(rows: List[List[Dict[str, Any]]]) -> Dict[str, Any]:
    """Build an occupancy grid from table rows with HTML-like rowspan and colspan."""
    grid = []
    ncols = 0
    for r, row in enumerate(rows or []):
        while len(grid) <= r:
            grid.append([])
        c = 0
        for cell in (row or []):
            while c < len(grid[r]) and grid[r][c] is not None:
                c += 1
            
            cs = max(1, min(30, int((cell.get('cs') if cell else 1) or 1)))
            rs = max(1, min(200, int((cell.get('rs') if cell else 1) or 1)))
            text = str((cell.get('t') if cell and 't' in cell else cell.get('text', '')) or '')
            is_bold = bool(cell.get('b') if cell else False)
            
            ref = {'t': text, 'b': is_bold, 'cs': cs, 'rs': rs, 'r': r, 'c': c}
            
            for i in range(rs):
                for j in range(cs):
                    row_idx = r + i
                    while len(grid) <= row_idx:
                        grid.append([])
                    while len(grid[row_idx]) <= c + j:
                        grid[row_idx].append(None)
                    grid[row_idx][c + j] = {
                        'ref': ref,
                        'origin': (i == 0 and j == 0),
                        'vcont': (i > 0 and j == 0),
                        'hcont': (j > 0)
                    }
            c += cs
            ncols = max(ncols, c)
            
    # Clip grid to rows length (rowspan cannot reach past the last row)
    grid = grid[:max(len(rows or []), 0)]
    return {'grid': grid, 'ncols': ncols}


def parse_num(s: Any) -> Optional[float]:
    t = str(s or '').replace('م²', '').replace('م2', '').replace('م', '').replace(' ', '').replace(',', '.')
    if re.match(r'^-?\d+(\.\d+)?$', t):
        try:
            return float(t)
        except ValueError:
            return None
    return None


def is_total_row(row: List[Optional[Dict[str, Any]]]) -> bool:
    return any(s and s['origin'] and re.search(r'المجموع|مجموع', s['ref']['t']) for s in row)


def check_table(tb: Dict[str, Any]) -> List[str]:
    """Check totals (subtotals vs rows above, grand total vs all data rows) and count '?' cells."""
    issues = []
    grid = tb['grid']
    ncols = tb['ncols']
    header_rows = tb.get('headerRows', 0)
    
    qs = sum(1 for row in grid for s in row if s and s['origin'] and '?' in s['ref']['t'])
    if qs:
        issues.append(f"{qs} خانة غير واضحة (مُعلَّمة بالأصفر)")

    seg = [0.0] * ncols
    all_totals = [0.0] * ncols
    seg_n = [0] * ncols
    all_n = [0] * ncols
    subtotals = 0

    for r, row in enumerate(grid):
        if r < header_rows:
            continue
        if is_total_row(row):
            label = ' '.join(s['ref']['t'] for s in row if s and s['origin'] and parse_num(s['ref']['t']) is None)
            grand = bool(re.search(r'الكلي|العام|الإجمالي', label)) or (subtotals > 0 and r == len(grid) - 1 and not re.search(r'الجزئي', label))
            for c in range(ncols):
                s = row[c]
                if not s or not s['origin']:
                    continue
                t = parse_num(s['ref']['t'])
                if t is None:
                    continue
                s_sum = all_totals[c] if grand else seg[c]
                n = all_n[c] if grand else seg_n[c]
                if n < 2:
                    continue
                int_col = (t == math.floor(t))
                diff = abs(s_sum - t)
                tol = 0.5 if int_col else max(0.15, abs(t) * 0.0005)
                if diff > tol:
                    headers = [hr[c]['ref']['t'].split('\n')[0] for hr in grid[:header_rows] if hr[c]]
                    h = headers[-1] if headers else f'عمود {c + 1}'
                    formatted_sum = f"{round(s_sum, 2):.2f}".rstrip('0').rstrip('.').replace('.', ',')
                    cleaned_label = re.sub(r'\s+', ' ', label).strip()
                    issues.append(f"«{cleaned_label}» في «{h}»: مجموع الأسطر {formatted_sum} والمكتوب {s['ref']['t']}")
            if not grand:
                subtotals += 1
                seg = [0.0] * ncols
                seg_n = [0] * ncols
            continue

        for c in range(ncols):
            s = row[c]
            if not s or not s['origin']:
                continue
            v = parse_num(s['ref']['t'])
            if v is None:
                continue
            seg[c] += v
            all_totals[c] += v
            seg_n[c] += 1
            all_n[c] += 1

    return issues


def runs_556(t: str, sz: int, bold: bool, underline: bool = False) -> str:
    u_tag = '<w:u w:val="single"/>' if underline else ''
    b_tag = '<w:b/><w:bCs/>' if bold else ''
    base = (f'<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            f'{b_tag}'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            f'{u_tag}')
    
    re_split = re.compile(f'({NUM_PATTERN_556.pattern}|{LATIN_PATTERN_556.pattern})')
    parts = re_split.split(str(t))
    out = []
    for s in parts:
        if not s:
            continue
        is_num_or_latin = bool(NUM_PATTERN_556.fullmatch(s) or LATIN_PATTERN_556.fullmatch(s))
        rtl_tag = '' if is_num_or_latin else '<w:rtl/>'
        out.append(f'<w:r><w:rPr>{base}{rtl_tag}</w:rPr><w:t xml:space="preserve">{esc_xml(s)}</w:t></w:r>')
    return ''.join(out)


def cell_para(t: str, bold: bool) -> str:
    content = runs_556(t, 22, bold) if t else ''
    return f'<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>{content}</w:p>'


def text_para(b: Dict[str, Any]) -> str:
    align = b.get('align')
    jc = 'center' if align == 'center' else ('end' if align == 'left' else 'start')
    bold = bool(b.get('bold'))
    underline = bool(b.get('underline'))
    big = (jc == 'center' and bold)
    text = b.get('text', '')
    before = 240 if big else 60
    after = 160 if big else 60
    sz = 30 if big else 26
    return f'<w:p><w:pPr><w:keepNext/><w:bidi/><w:spacing w:before="{before}" w:after="{after}"/><w:jc w:val="{jc}"/></w:pPr>{runs_556(text, sz, bold, underline)}</w:p>'


def compute_widths(tb: Dict[str, Any], TW: int) -> List[int]:
    grid = tb['grid']
    ncols = tb['ncols']
    header_rows = tb.get('headerRows', 0)
    
    col_sum = [0] * ncols
    col_cnt = [0] * ncols
    col_max = [0] * ncols
    col_hw = [0] * ncols

    for r, row in enumerate(grid):
        for c, s in enumerate(row):
            if not s or not s['origin'] or s['ref']['cs'] != 1 or not s['ref']['t']:
                continue
            lines = s['ref']['t'].split('\n')
            if r < header_rows:
                max_w = max((len(w) for l in lines for w in re.split(r'\s+', l) if w), default=0)
                col_hw[c] = max(col_hw[c], max_w)
                continue
            L = max((len(l) for l in lines), default=0)
            col_sum[c] += L
            col_cnt[c] += 1
            col_max[c] = max(col_max[c], L)

    need = [max(col_hw[c] * 150, (col_max[c] + 2) * 120 if col_max[c] <= 16 else 0, 520) + 260 for c in range(ncols)]
    long_cols = [c for c in range(ncols) if col_max[c] > 16]
    w = list(need)
    for c in long_cols:
        avg_len = col_sum[c] / max(col_cnt[c], 1)
        w[c] = max(need[c], min(40, avg_len) * 120 + 230)

    total = sum(w)
    if total > TW and long_cols:
        fixed = sum(w[c] for c in range(ncols) if c not in long_cols)
        rest = max(TW - fixed, len(long_cols) * 2600)
        lsum = sum(w[c] for c in long_cols)
        if lsum > 0:
            for c in long_cols:
                w[c] = (w[c] * rest) / lsum

    t2 = sum(w)
    if t2 > TW and t2 > 0:
        w = [(x * TW) / t2 for x in w]
    elif t2 < TW * 0.75 and t2 > 0:
        w = [(x * TW * 0.75) / t2 for x in w]

    return [int(x) for x in w]


def table_xml(tb: Dict[str, Any], TW: int) -> str:
    grid = tb['grid']
    ncols = tb['ncols']
    header_rows = tb.get('headerRows', 0)
    w = compute_widths(tb, TW)
    tw = sum(w)

    area_cols = set()
    for row in grid[:header_rows]:
        for c, s in enumerate(row):
            if s and s['origin'] and re.search(r'مساحة|المساحة', s['ref']['t']):
                for j in range(s['ref']['cs']):
                    area_cols.add(c + j)

    M = '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:start w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:end w:w="80" w:type="dxa"/></w:tcMar>'
    borders_str = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for s in ['top', 'start', 'bottom', 'end', 'insideH', 'insideV'])
    
    grid_cols_xml = ''.join(f'<w:gridCol w:w="{v}"/>' for v in w)
    x = (f'<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="{tw}" w:type="dxa"/><w:jc w:val="center"/>'
         f'<w:tblBorders>{borders_str}</w:tblBorders>'
         f'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:start w:w="10" w:type="dxa"/><w:end w:w="10" w:type="dxa"/></w:tblCellMar></w:tblPr>'
         f'<w:tblGrid>{grid_cols_xml}</w:tblGrid>')

    for r, row in enumerate(grid):
        head = (r < header_rows)
        tot = (not head and is_total_row(row))
        fill = 'D9D9D9' if head else ('F2F2F2' if tot else '')
        x += '<w:tr><w:trPr><w:cantSplit/><w:jc w:val="center"/></w:trPr>'
        
        c = 0
        while c < ncols:
            s = row[c]
            if s and s.get('hcont'):
                c += 1
                continue
            cs = s['ref']['cs'] if s else 1
            width = sum(w[c:c+cs])
            pr = f'<w:tcW w:w="{width}" w:type="dxa"/>' + (f'<w:gridSpan w:val="{cs}"/>' if cs > 1 else '')
            text = ''
            if s and s.get('vcont'):
                pr += '<w:vMerge/>'
            elif s and s.get('origin'):
                if s['ref']['rs'] > 1:
                    pr += '<w:vMerge w:val="restart"/>'
                text = s['ref']['t']

            q = '?' in text
            shade = 'FFF2CC' if q else fill
            if shade:
                pr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
            pr += M + '<w:vAlign w:val="center"/>'

            t = text
            if not head and not tot and c in area_cols and parse_num(t) is not None and 'م' not in t:
                t += ' م²'
            bold = head or tot or (s and s['ref']['b'])
            paras = ''.join(cell_para(line, bold) for line in t.split('\n')) if t else cell_para('', False)
            x += f'<w:tc><w:tcPr>{pr}</w:tcPr>{paras}</w:tc>'
            c += cs

        x += '</w:tr>'

    x += '</w:tbl>'
    return x


def build_landscape_doc(template_dir: str, results: Dict[int, List[Dict[str, Any]]], num_pages: int) -> Tuple[bytes, List[Dict[str, Any]]]:
    """Rebuild landscape doc with real Word tables from AI blocks."""
    doc_xml_path = os.path.join(template_dir, 'word', 'document.xml')
    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        shell_doc = f.read()

    TW = 16838 - 2 * 1417  # 14004
    segs = []
    
    def last_real():
        for i in range(len(segs) - 1, -1, -1):
            if segs[i]['t'] != 'gap':
                return segs[i]
        return None

    for p in range(num_pages):
        page_blocks = results.get(p)
        if not page_blocks:
            # Failed page note
            fail_p = f'<w:p><w:pPr><w:bidi/><w:jc w:val="center"/></w:pPr>{runs_556("⚠ لم تُقرأ الصفحة " + str(p + 1) + " من ملف PDF — أعيدي قراءتها أو اكتبيها يدوياً.", 24, True)}</w:p>'
            segs.append({'t': 'xml', 'x': fail_p})
            continue

        first_on_page = True
        for bl in page_blocks:
            b_type = bl.get('type')
            if b_type == 'text':
                if str(bl.get('text', '')).strip():
                    segs.append({'t': 'text', 'b': bl})
                first_on_page = False
                continue
            if b_type != 'table' or not isinstance(bl.get('rows'), list) or not bl.get('rows'):
                continue
                
            g = to_grid(bl['rows'])
            hr = max(0, min(int(bl.get('header_rows') or 0), len(g['grid'])))
            prev = last_real()
            cont = (first_on_page and hr == 0 and prev and prev['t'] == 'table' and prev['ncols'] == g['ncols'] and (bl.get('continues_previous', True)))
            if cont:
                while segs and segs[-1]['t'] == 'gap':
                    segs.pop()
                prev['grid'].extend(g['grid'])
            else:
                segs.append({'t': 'table', 'grid': g['grid'], 'ncols': g['ncols'], 'headerRows': hr})
            first_on_page = False

    body = ''
    report = []
    k = 0
    gap = '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>'

    for i, s in enumerate(segs):
        if s['t'] in ['xml', 'gap']:
            body += s['x']
            continue
        if s['t'] == 'text':
            body += text_para(s['b'])
            continue
        k += 1
        body += table_xml(s, TW)
        nxt = segs[i + 1] if i + 1 < len(segs) else None
        if not nxt or nxt['t'] == 'table':
            body += gap
        report.append({'n': k, 'rows': len(s['grid']), 'issues': check_table(s)})

    final_doc = shell_doc.replace('{{BODY}}', body)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for root_dir, _, files in os.walk(template_dir):
            for file in files:
                abs_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(abs_path, template_dir).replace('\\', '/')
                if rel_path == 'word/document.xml':
                    zout.writestr(rel_path, final_doc.encode('utf-8'))
                else:
                    with open(abs_path, 'rb') as fp:
                        zout.writestr(rel_path, fp.read())

    return zip_buffer.getvalue(), report
