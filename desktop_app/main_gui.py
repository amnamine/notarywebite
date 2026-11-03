"""
تطبيق مكتبي (Tkinter) لمكتب الأستاذ بوخاري مريم – موثق
إجراء إشهار عقاري: معالجة ملفات Word ووضعها داخل الإطار الرسمي تلقائياً دون الحاجة لإنترنت.
"""
import sys
import os
import io
import re
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Any, List, Optional, Tuple

# Resolve application base path (handles both running from source and compiled PyInstaller .exe)
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    APP_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = BASE_DIR

# Check templates directory
TEMPLATES_DIR = os.path.join(BASE_DIR, 'docx_templates', 'ijraa_ichhar_template')
if not os.path.exists(TEMPLATES_DIR):
    # Fallback to parent docx_templates if running from source in desktop_app
    alt_tpl = os.path.join(os.path.dirname(BASE_DIR), 'docx_templates', 'ijraa_ichhar_template')
    if os.path.exists(alt_tpl):
        TEMPLATES_DIR = alt_tpl

# ----------------- XML & DOCX ENGINE -----------------
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
    re_child = re.compile(rf'<w:tbl>|</w:tbl>|<w:{tag}(?=[ >])[^>]*>|</w:{tag}>')
    out = []
    tdepth = 0
    st = -1
    for m in re_child.finditer(x):
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

def read_styles(styles_xml: Optional[str]) -> Dict[str, Any]:
    styles = {}
    defaults = {'sz': None}
    if not styles_xml:
        return {'styles': styles, 'defaults': defaults}
    
    dd_m = re.search(r'<w:rPrDefault>[\s\S]*?</w:rPrDefault>', styles_xml)
    if dd_m:
        sz_m = re.search(r'<w:sz w:val="(\d+)"', dd_m.group(0))
        if sz_m:
            defaults['sz'] = int(sz_m.group(1))

    re_style = re.compile(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)</w:style>')
    for m in re_style.finditer(styles_xml):
        id_ = m.group(1)
        x = m.group(2)
        based_m = re.search(r'<w:basedOn w:val="([^"]+)"', x)
        jc_m = re.search(r'<w:pPr>[\s\S]*?<w:jc w:val="(\w+)"', x)
        rpr_m = re.search(r'<w:rPr>[\s\S]*?</w:rPr>', x)
        rpr = rpr_m.group(0) if rpr_m else ''
        
        b = bool(re.search(r'<w:b/>|<w:b w:val="(1|true|on)"', rpr))
        u = bool(re.search(r'<w:u w:val="(?!none)', rpr))
        sz_val_m = re.search(r'<w:sz w:val="(\d+)"', rpr)
        sz = int(sz_val_m.group(1)) if sz_val_m else None
        
        styles[id_] = {
            'raw': x,
            'based': based_m.group(1) if based_m else None,
            'jc': jc_m.group(1) if jc_m else None,
            'b': b,
            'u': u,
            'sz': sz
        }
    return {'styles': styles, 'defaults': defaults}

def style_props(st: Dict[str, Any], id_: Optional[str]) -> Dict[str, Any]:
    out = {'jc': None, 'b': False, 'u': False, 'sz': None}
    guard = 0
    curr_id = id_
    while curr_id and curr_id in st['styles'] and guard < 10:
        guard += 1
        s = st['styles'][curr_id]
        if not out['jc'] and s['jc']:
            out['jc'] = s['jc']
        if s['b']:
            out['b'] = True
        if s['u']:
            out['u'] = True
        if not out['sz'] and s['sz']:
            out['sz'] = s['sz']
        curr_id = s['based']
    return out

def clean_xml(x: str) -> str:
    x = re.sub(r' w:rsid\w*="[^"]*"| w14:(paraId|textId)="[^"]*"', '', x)
    x = re.sub(r'<w:proofErr [^>]*/>|<w:bookmark(Start|End) [^>]*/>|<w:lastRenderedPageBreak/>|<w:noProof/>', '', x)
    x = re.sub(r'<w:r>(?:(?!</w:r>)[\s\S])*?<w:(drawing|pict|object|footnoteReference|endnoteReference|commentReference)[\s\S]*?</w:r>', '', x)
    x = re.sub(r'<w:br w:type="(page|column)"/>', '', x)
    x = re.sub(r'<w:commentRange(Start|End) [^>]*/>', '', x)
    x = re.sub(r'<w:ins [^>]*>|</w:ins>', '', x)
    x = re.sub(r'<w:del [^>]*>[\s\S]*?</w:del>', '', x)
    x = re.sub(r'<w:hyperlink[^>]*>|</w:hyperlink>', '', x)
    x = re.sub(r'<w:smartTag[^>]*>|</w:smartTag>|<w:smartTagPr>[\s\S]*?</w:smartTagPr>', '', x)
    return x

def order_rpr(r: str) -> str:
    if not r.startswith('<w:rPr>') or not r.endswith('</w:rPr>'):
        return r
    inner = r[7:-8]
    re_tag = re.compile(r'<w:(\w+)(?:\s[^>]*)?/>|<w:(\w+)(?:\s[^>]*)?>[\s\S]*?</w:\2>')
    els = []
    seen = set()
    for m in re_tag.finditer(inner):
        e = m.group(0)
        n = re.search(r'<w:(\w+)', e).group(1)
        if n not in seen:
            seen.add(n)
            k = RPR_ORDER.index(n) if n in RPR_ORDER else 99
            els.append({'n': n, 'e': e, 'k': k})
    els.sort(key=lambda o: o['k'])
    return '<w:rPr>' + ''.join(o['e'] for o in els) + '</w:rPr>'

def read_fonts(styles_xml: Optional[str], theme_xml: Optional[str]) -> Dict[str, Any]:
    theme = {}
    theme_xml = theme_xml or ''
    for k in ['major', 'minor']:
        blk_m = re.search(rf'<a:{k}Font>[\s\S]*?</a:{k}Font>', theme_xml)
        blk = blk_m.group(0) if blk_m else ''
        latin_m = re.search(r'<a:latin typeface="([^"]*)"', blk)
        cs_m = re.search(r'<a:cs typeface="([^"]*)"', blk)
        theme[k + 'Latin'] = latin_m.group(1) if latin_m else ''
        theme[k + 'Cs'] = cs_m.group(1) if cs_m else ''

    def res(name: Optional[str]) -> str:
        if not name:
            return ''
        k = 'major' if name.startswith('major') else 'minor'
        if 'Bidi' in name:
            return theme[k + 'Cs'] or theme[k + 'Latin']
        return theme[k + 'Latin']

    styles_xml = styles_xml or ''
    dd_m = re.search(r'<w:rPrDefault>[\s\S]*?</w:rPrDefault>', styles_xml)
    dd = dd_m.group(0) if dd_m else ''
    rf_m = re.search(r'<w:rFonts [^>]*/>', dd)
    rf = rf_m.group(0) if rf_m else ''

    def at(a: str) -> Optional[str]:
        m = re.search(rf'w:{a}="([^"]*)"', rf)
        return m.group(1) if m else None

    ascii_font = at('ascii') or res(at('asciiTheme')) or 'Times New Roman'
    cs_font = at('cs') or res(at('cstheme')) or ascii_font
    return {'ascii': ascii_font, 'cs': cs_font, 'res': res}

def explicit_fonts(r: str, fonts: Dict[str, Any]) -> str:
    m = re.search(r'<w:rFonts [^>]*/>', r)
    def at(a: str) -> Optional[str]:
        if not m:
            return None
        res_m = re.search(rf'w:{a}="([^"]*)"', m.group(0))
        return res_m.group(1) if res_m else None

    ascii_f = at('ascii') or fonts['res'](at('asciiTheme')) or fonts['ascii']
    hansi_f = at('hAnsi') or fonts['res'](at('hAnsiTheme')) or ascii_f
    cs_f = at('cs') or fonts['res'](at('cstheme')) or fonts['cs']
    el = f'<w:rFonts w:ascii="{esc_xml(ascii_f)}" w:hAnsi="{esc_xml(hansi_f)}" w:cs="{esc_xml(cs_f)}"/>'
    return r.replace(m.group(0), el) if m else r.replace('<w:rPr>', '<w:rPr>' + el)

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
    x = clean_xml(x)
    x = re.sub(r'<w:numPr>[\s\S]*?</w:numPr>', '', x)
    
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
    x = re.sub(r'<w:tblpPr [^>]*/>', '', x)
    return x

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

    return {
        'qism': ' و '.join(qism[:6]),
        'majmoua': '، '.join(maj[:20]),
        'baladiya': bal,
        'wilaya': wil,
    }

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


# ----------------- TKINTER GUI INTERFACE -----------------
class NotaryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("مكتب الأستاذ بوخاري مريم — إجراء إشهار عقاري")
        self.geometry("820x680")
        self.minsize(720, 600)
        self.configure(bg="#f3f4f6")

        self.current_file_path = None
        self.analysed_data = None

        self.setup_styles()
        self.build_ui()

    def setup_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        
        # Primary colors
        self.c_band = "#16213a"
        self.c_accent = "#0e6b62"
        self.c_accent_hover = "#128277"
        self.c_paper = "#f3f4f6"
        self.c_sheet = "#ffffff"
        self.c_ink = "#151c2b"
        self.c_muted = "#5d6676"
        self.c_line = "#dfe3ea"

        self.style.configure('TFrame', background=self.c_paper)
        self.style.configure('Sheet.TFrame', background=self.c_sheet, relief='flat')
        self.style.configure('TLabel', background=self.c_sheet, foreground=self.c_ink, font=('Segoe UI', 10))
        self.style.configure('Muted.TLabel', background=self.c_sheet, foreground=self.c_muted, font=('Segoe UI', 9))
        self.style.configure('Header.TLabel', background=self.c_band, foreground="#ffffff", font=('Segoe UI', 14, 'bold'))
        self.style.configure('SubHeader.TLabel', background=self.c_band, foreground="#a9b3c7", font=('Segoe UI', 9))
        
        self.style.configure('Action.TButton', background=self.c_accent, foreground="#ffffff", font=('Segoe UI', 10, 'bold'), borderwidth=0, padding=8)
        self.style.map('Action.TButton', background=[('active', self.c_accent_hover), ('disabled', '#cccccc')])

    def build_ui(self):
        # 1. Top ink header band
        band = tk.Frame(self, bg=self.c_band, height=72)
        band.pack(fill='x', side='top')
        band.pack_propagate(False)

        header_container = tk.Frame(band, bg=self.c_band)
        header_container.pack(expand=True, fill='both', padx=24, pady=10)

        # Seal
        seal_canvas = tk.Canvas(header_container, width=48, height=48, bg=self.c_band, highlightthickness=0)
        seal_canvas.pack(side='right', padx=(12, 0))
        seal_canvas.create_oval(3, 3, 45, 45, outline='#eef1f7', width=1.5)
        seal_canvas.create_oval(7, 7, 41, 41, outline='#eef1f7', width=1)
        seal_canvas.create_text(24, 25, text="ب.م", fill='#eef1f7', font=('Traditional Arabic', 13, 'bold'))

        # Title
        title_box = tk.Frame(header_container, bg=self.c_band)
        title_box.pack(side='right', anchor='center')
        lbl_title = tk.Label(title_box, text="مكتب الأستاذ بوخاري مريم", bg=self.c_band, fg="#ffffff", font=('Segoe UI', 13, 'bold'))
        lbl_title.pack(anchor='e')
        lbl_sub = tk.Label(title_box, text="موثق — الجزائر  |  إجراء إشهار عقاري", bg=self.c_band, fg="#a9b3c7", font=('Segoe UI', 9))
        lbl_sub.pack(anchor='e')

        # 2. Main content area (Scrollable)
        main_canvas = tk.Canvas(self, bg=self.c_paper, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas, style='TFrame')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=780)
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True, padx=16, pady=12)
        scrollbar.pack(side="right", fill="y")

        # Step 1: File selection
        f1 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=18, pady=14)
        f1.pack(fill='x', pady=6)
        
        lbl_s1 = tk.Label(f1, text="1. ملف العقد (Word)", bg=self.c_sheet, fg=self.c_accent, font=('Segoe UI', 11, 'bold'))
        lbl_s1.pack(anchor='e')

        drop_box = tk.Frame(f1, bg=self.c_paper, bd=1, relief='groove', padx=12, pady=14)
        drop_box.pack(fill='x', pady=8)

        btn_browse = tk.Button(drop_box, text="📁 اختيار ملف Word (.docx)", bg=self.c_accent, fg="#ffffff",
                               font=('Segoe UI', 10, 'bold'), relief='flat', padx=14, pady=6, cursor='hand2',
                               command=self.browse_file)
        btn_browse.pack(pady=4)

        self.lbl_file = tk.Label(drop_box, text="لم يتم اختيار أي ملف بعد", bg=self.c_paper, fg=self.c_muted, font=('Segoe UI', 9))
        self.lbl_file.pack()

        self.lbl_stats = tk.Label(drop_box, text="", bg=self.c_paper, fg=self.c_accent, font=('Segoe UI', 9, 'bold'))
        self.lbl_stats.pack(pady=2)

        # Step 2: Cadastral box
        f2 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=18, pady=14)
        f2.pack(fill='x', pady=6)

        lbl_s2 = tk.Label(f2, text="2. مراجع مسح الأراضي (تُملأ تلقائياً من نص الملف)", bg=self.c_sheet, fg=self.c_accent, font=('Segoe UI', 11, 'bold'))
        lbl_s2.pack(anchor='e')

        grid2 = tk.Frame(f2, bg=self.c_sheet)
        grid2.pack(fill='x', pady=8)

        self.fields = {}
        row_fields = [
            ("قسم", "qism"),
            ("بلدية", "baladiya"),
            ("مجموعة ملكية رقم", "majmoua"),
            ("حصة رقم", "hissa"),
            ("الحفظ العقاري لولاية", "wilaya")
        ]
        for idx, (label_txt, key) in enumerate(row_fields):
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
        f3 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=18, pady=14)
        f3.pack(fill='x', pady=6)

        lbl_s3 = tk.Label(f3, text="3. الإيداع والرسم (اختياري)", bg=self.c_sheet, fg=self.c_accent, font=('Segoe UI', 11, 'bold'))
        lbl_s3.pack(anchor='e')

        lbl_s3_hint = tk.Label(f3, text="ما تتركه فارغاً يبقى نقاطاً «......» لتُكتب باليد.", bg=self.c_sheet, fg=self.c_muted, font=('Segoe UI', 8))
        lbl_s3_hint.pack(anchor='e')

        grid3 = tk.Frame(f3, bg=self.c_sheet)
        grid3.pack(fill='x', pady=8)

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

        # Step 4: Conversion Action
        f4 = tk.Frame(scrollable_frame, bg=self.c_sheet, bd=1, relief='solid', padx=18, pady=16)
        f4.pack(fill='x', pady=6)

        self.btn_convert = tk.Button(f4, text="⚡ إنشاء ملف «إجراء إشهار عقاري» الجديد", bg="#1b6a3a", fg="#ffffff",
                                     font=('Segoe UI', 11, 'bold'), relief='flat', padx=18, pady=10, cursor='hand2',
                                     state='disabled', command=self.do_convert)
        self.btn_convert.pack(fill='x', pady=4)

        self.lbl_result = tk.Label(f4, text="", bg=self.c_sheet, fg="#1b6a3a", font=('Segoe UI', 9, 'bold'))
        self.lbl_result.pack(pady=4)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="اختيار ملف العقد Word",
            filetypes=[("Word Document (.docx)", "*.docx"), ("Old Word (.doc)", "*.doc")]
        )
        if not file_path:
            return

        if file_path.lower().endswith('.doc'):
            messagebox.showwarning("تنبيه صيغة الملف", "هذا ملف .doc قديم. افتحه في Word ثم «حفظ باسم» واختر .docx، وضعه هنا من جديد.")
            return

        if not file_path.lower().endswith('.docx'):
            messagebox.showerror("خطأ", "الملف يجب أن يكون Word بصيغة .docx.")
            return

        try:
            with open(file_path, 'rb') as fp:
                raw_bytes = fp.read()
            
            analysed = analyse_docx(raw_bytes)
            self.current_file_path = file_path
            self.analysed_data = analysed

            fname = os.path.basename(file_path)
            self.lbl_file.config(text=f"الملف الحالي: {fname}", fg=self.c_ink)

            n_p = sum(1 for b in analysed['blocks'] if b['t'] == 'p' and b['x'].strip())
            n_tbl = sum(1 for b in analysed['blocks'] if b['t'] == 'tbl')
            boxed_txt = " | الملف داخل الإطار أصلاً" if analysed['boxed'] else ""
            self.lbl_stats.config(text=f"عدد الفقرات: {n_p}  |  عدد الجداول: {n_tbl}{boxed_txt}")

            # Fill extracted cadastral values
            box = analysed.get('box', {})
            for key in ['qism', 'baladiya', 'majmoua', 'wilaya']:
                if key in self.fields and key in box:
                    self.fields[key].delete(0, tk.END)
                    self.fields[key].insert(0, box[key])

            self.btn_convert.config(state='normal')
            self.lbl_result.config(text="")
        except Exception as e:
            messagebox.showerror("خطأ في فحص الملف", f"تعذر فتح وقراءة الملف:\n{str(e)}")

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
    app = NotaryApp()
    app.mainloop()
