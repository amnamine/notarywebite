"""
XML helpers for DOCX manipulation.
Ports the cleanXml, orderRPr, plain, esc, decode, topBlocks, children, and style readers from core.js.
"""
import re
from typing import List, Dict, Any, Optional

FONT = 'Arial'
RF = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
SZ = {20: 20, 22: 22, 24: 24, 26: 24, 28: 26, 32: 28, 36: 30, 40: 32, 44: 34, 48: 36}

NUM_PATTERN = re.compile(r'\d+(?:\s?[\/\-]\s?\d+)+')
LATIN_PATTERN = re.compile(r'[A-Za-z][A-Za-z0-9.]*(?:\s?-\s?[0-9]+)?')

RPR_ORDER = [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
    'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
    'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
    'specVanish', 'oMath'
]

def map_sz(v: int) -> int:
    if v in SZ:
        return SZ[v]
    return 36 if v > 48 else v

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
    """Split XML into top-level blocks (<w:p>, <w:tbl>) of the given string."""
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
    """Find children <w:TAG> at tbl-depth 0 of an xml fragment (rows of table, cells of row)."""
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
