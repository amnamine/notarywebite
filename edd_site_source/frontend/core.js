/* EDD converter core: one Word deed in -> two Word files out
   (boxed "إجراء إشهار عقاري" template + landscape 556 style).
   Works in the browser and in node (needs JSZip). */
(function (root) {
  'use strict';

  // ---------- small helpers ----------
  const stripTags = (x) => x.replace(/<[^>]+>/g, '');
  const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const plain = (x) => decode(stripTags(x.replace(/<w:tab\/>/g, ' ').replace(/<w:br[^>]*\/>/g, ' ')));
  const FONT = 'Arial';
  const RF = `<w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>`;
  const SZ = { 20: 20, 22: 22, 24: 24, 26: 24, 28: 26, 32: 28, 36: 30, 40: 32, 44: 34, 48: 36 };
  const mapSz = (v) => (SZ[v] !== undefined ? SZ[v] : (v > 48 ? 36 : v));
  const NUM = /\d+(?:\s?[\/\-]\s?\d+)+/;
  const LATIN = /[A-Za-z][A-Za-z0-9.]*(?:\s?-\s?[0-9]+)?/;

  // split xml into top-level blocks (<w:p>, <w:tbl>) of the given string
  function topBlocks(body) {
    const out = [];
    const re = /<(\/?)w:(tbl|p)(?=[ >\/])[^>]*?(\/?)>/g;
    let depth = 0, st = 0, m;
    while ((m = re.exec(body))) {
      const close = m[1], tag = m[2], self = m[3];
      if (tag === 'tbl') {
        if (!close) { if (depth === 0) st = m.index; depth++; }
        else { depth--; if (depth === 0) out.push({ t: 'tbl', x: body.slice(st, re.lastIndex) }); }
      } else if (depth === 0) {
        if (!close && self) out.push({ t: 'p', x: m[0] });
        else if (!close) st = m.index;
        else out.push({ t: 'p', x: body.slice(st, re.lastIndex) });
      }
    }
    return out;
  }
  // children <w:TAG> at tbl-depth 0 of an xml fragment (rows of a table, cells of a row)
  function children(x, tag) {
    const re = new RegExp(`<w:tbl>|</w:tbl>|<w:${tag}(?=[ >])[^>]*>|</w:${tag}>`, 'g');
    const out = []; let tdepth = 0, st = -1, m;
    while ((m = re.exec(x))) {
      const s = m[0];
      if (s === '<w:tbl>') tdepth++;
      else if (s === '</w:tbl>') tdepth--;
      else if (tdepth === 0) {
        if (s.startsWith('</')) { if (st >= 0) out.push(x.slice(st, re.lastIndex)); st = -1; }
        else st = m.index;
      }
    }
    return out;
  }

  // ---------- source styles: resolve paragraph-style jc and run props ----------
  function readStyles(stylesXml) {
    const styles = {}; const defaults = { sz: null };
    if (!stylesXml) return { styles, defaults };
    const dd = stylesXml.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/);
    if (dd) { const s = dd[0].match(/<w:sz w:val="(\d+)"/); if (s) defaults.sz = +s[1]; }
    const re = /<w:style [^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g; let m;
    while ((m = re.exec(stylesXml))) {
      const id = m[1], x = m[2];
      const based = (x.match(/<w:basedOn w:val="([^"]+)"/) || [])[1];
      const jc = (x.match(/<w:pPr>[\s\S]*?<w:jc w:val="(\w+)"/) || [])[1];
      const rpr = (x.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
      styles[id] = { raw: x, based, jc, b: /<w:b\/>|<w:b w:val="(1|true|on)"/.test(rpr), u: /<w:u w:val="(?!none)/.test(rpr), sz: +((rpr.match(/<w:sz w:val="(\d+)"/) || [])[1] || 0) || null };
    }
    return { styles, defaults };
  }
  function styleProps(st, id) {
    const out = { jc: null, b: false, u: false, sz: null }; let guard = 0;
    while (id && st.styles[id] && guard++ < 10) {
      const s = st.styles[id];
      if (!out.jc && s.jc) out.jc = s.jc;
      if (s.b) out.b = true; if (s.u) out.u = true;
      if (!out.sz && s.sz) out.sz = s.sz;
      id = s.based;
    }
    return out;
  }

  // ---------- paragraphs ----------
  function cleanXml(x) {
    return x
      .replace(/ w:rsid\w*="[^"]*"| w14:(paraId|textId)="[^"]*"/g, '')
      .replace(/<w:proofErr [^>]*\/>|<w:bookmark(Start|End) [^>]*\/>|<w:lastRenderedPageBreak\/>|<w:noProof\/>/g, '')
      .replace(/<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:(drawing|pict|object|footnoteReference|endnoteReference|commentReference)[\s\S]*?<\/w:r>/g, '')
      .replace(/<w:br w:type="(page|column)"\/>/g, '')
      .replace(/<w:commentRange(Start|End) [^>]*\/>/g, '')
      .replace(/<w:ins [^>]*>|<\/w:ins>/g, '')
      .replace(/<w:del [^>]*>[\s\S]*?<\/w:del>/g, '')
      .replace(/<w:hyperlink[^>]*>|<\/w:hyperlink>/g, '')
      .replace(/<w:smartTag[^>]*>|<\/w:smartTag>|<w:smartTagPr>[\s\S]*?<\/w:smartTagPr>/g, '');
  }
  function fixRunProps(x, st, sp) {
    // normalise every run: Arial, mapped size, style bold/underline
    return x.replace(/<w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?/g, (m, rpr) => {
      let r = rpr || '<w:rPr></w:rPr>';
      r = r.replace(/<w:rFonts [^>]*\/>/g, '').replace(/<w:rStyle [^>]*\/>/g, '');
      if (sp.b && !/<w:b[ \/]/.test(r)) r = r.replace('<w:rPr>', '<w:rPr><w:b/><w:bCs/>');
      if (sp.u && !/<w:u /.test(r)) r = r.replace('</w:rPr>', '<w:u w:val="single"/></w:rPr>');
      const def = sp.sz || st.defaults.sz || 28;
      if (!/<w:sz /.test(r)) r = r.replace('</w:rPr>', `<w:sz w:val="${def}"/><w:szCs w:val="${def}"/></w:rPr>`);
      else if (!/<w:szCs /.test(r)) r = r.replace(/(<w:sz w:val="(\d+)"\/>)/, '$1<w:szCs w:val="$2"/>');
      r = r.replace(/<w:(sz|szCs) w:val="(\d+)"\/>/g, (q, a, v) => `<w:${a} w:val="${mapSz(+v)}"/>`);
      r = r.replace('<w:rPr>', '<w:rPr>' + RF);
      // keep schema order: rFonts, b, bCs ... ; move rtl after sz/u if present (Word tolerant, validator strict)
      r = orderRPr(r);
      return '<w:r>' + r;
    });
  }
  const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
  function orderRPr(r) {
    const inner = r.slice(7, -8);
    const els = inner.match(/<w:(\w+)(?:\s[^>]*)?\/>|<w:(\w+)(?:\s[^>]*)?>[\s\S]*?<\/w:\2>/g) || [];
    const seen = {};
    const keyed = els.map((e) => { const n = e.match(/<w:(\w+)/)[1]; return { n, e, k: RPR_ORDER.indexOf(n) < 0 ? 99 : RPR_ORDER.indexOf(n) }; })
      .filter((o) => (seen[o.n] ? false : (seen[o.n] = true)));
    keyed.sort((a, b) => a.k - b.k);
    return '<w:rPr>' + keyed.map((o) => o.e).join('') + '</w:rPr>';
  }
  function mergeRuns(x) {
    const RUN = /<w:r>(<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:t(?: xml:space="preserve")?>([^<]*)<\/w:t><\/w:r>/g;
    return x.replace(/(?:<w:r>(?:<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:t(?: xml:space="preserve")?>[^<]*<\/w:t><\/w:r>){2,}/g, (seq) => {
      const out = []; let cur = null, m; RUN.lastIndex = 0;
      const key = (r) => r.replace(/<w:lang [^>]*\/>/g, '');
      while ((m = RUN.exec(seq))) {
        const rpr = m[1] || '';
        if (cur && key(cur[0]) === key(rpr)) cur[1] += m[2];
        else { if (cur) out.push(cur); cur = [rpr, m[2]]; }
      }
      if (cur) out.push(cur);
      return out.map(([a, b]) => `<w:r>${a}<w:t xml:space="preserve">${b}</w:t></w:r>`).join('');
    });
  }
  function splitNumbers(x) {
    // numbers like 91/148 or 2026/10/05 inside right-to-left runs -> own left-to-right run (shown as typed)
    return x.replace(/<w:r>(<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)<w:t(?: xml:space="preserve")?>([^<]*)<\/w:t><\/w:r>/g, (m, rpr, t) => {
      if (!rpr.includes('<w:rtl/>') || !NUM.test(t)) return m;
      const parts = t.split(new RegExp('(' + NUM.source + ')'));
      return parts.filter((p) => p).map((p) => NUM.test(p) && new RegExp('^' + NUM.source + '$').test(p)
        ? `<w:r>${rpr.replace('<w:rtl/>', '')}<w:t xml:space="preserve">${p}</w:t></w:r>`
        : `<w:r>${rpr}<w:t xml:space="preserve">${p}</w:t></w:r>`).join('');
    });
  }
  function wantDashes(textTrim, jc) {
    if (!textTrim) return false;
    if (['center', 'start', 'end', 'right', 'left'].includes(jc)) return false;
    if (/-{3,}$/.test(textTrim)) return true;
    if (textTrim.length < 60) return false;
    return textTrim.endsWith('.') || /(أدناه|ادناه|أدنـ+اه|يلي\s*:|التالية\s*:)$/.test(textTrim);
  }
  function addDashes(p, tabPos) {
    // strip typed trailing dashes, then a tab with a "-----" leader to the end of the line
    const runs = [...p.matchAll(/<w:r>(?:(?!<w:r>)[\s\S])*?<\/w:r>/g)];
    let lastRpr = '';
    for (let i = runs.length - 1; i >= 0; i--) {
      const r = runs[i][0];
      const tm = r.match(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/);
      if (!tm) continue;
      if (!lastRpr) { const rp = r.match(/<w:rPr>[\s\S]*?<\/w:rPr>/); lastRpr = rp ? rp[0] : ''; }
      if (/-\s*$/.test(tm[2])) {
        const nt = tm[2].replace(/[\s\-]*-[\s\-]*$/, '');
        const r2 = r.replace(tm[0], tm[1] + nt + tm[3]);
        p = p.slice(0, runs[i].index) + r2 + p.slice(runs[i].index + r.length);
        if (nt.trim()) break;
        continue;
      }
      break;
    }
    let rp = lastRpr || '<w:rPr><w:rtl/></w:rPr>';
    if (!rp.includes('<w:rtl/>')) rp = orderRPr(rp.replace('</w:rPr>', '<w:rtl/></w:rPr>'));
    p = p.slice(0, p.lastIndexOf('</w:p>')) + `<w:r>${rp}<w:tab/></w:r></w:p>`;
    return p.replace('<w:pPr>', `<w:pPr><w:tabs><w:tab w:val="left" w:leader="hyphen" w:pos="${tabPos}"/></w:tabs>`);
  }
  function convertParagraph(x, st, mode) {
    x = cleanXml(x);
    const pm = x.match(/^<w:p(?: [^>]*)?>(<w:pPr>[\s\S]*?<\/w:pPr>)?/);
    const ppr = pm[1] || '';
    const pstyle = (ppr.match(/<w:pStyle w:val="([^"]+)"/) || [])[1];
    const sp = styleProps(st, pstyle || 'Normal');
    let jc = (ppr.match(/<w:jc w:val="(\w+)"\/>/) || [])[1] || sp.jc || 'both';
    if (/kashida|distribute|thaiDistribute/i.test(jc)) jc = 'both';
    let markRpr = (ppr.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
    if (markRpr) markRpr = orderRPr(markRpr.replace(/<w:rFonts [^>]*\/>|<w:rStyle [^>]*\/>|<w:ins [^>]*\/>|<w:del [^>]*\/>/g, '').replace(/<w:(sz|szCs) w:val="(\d+)"\/>/g, (q, a, v) => `<w:${a} w:val="${mapSz(+v)}"/>`));
    let body = x.slice(pm[0].length);
    // dashes already added to this paragraph before (a tab at its end): remove them, they are added again once below
    body = body.replace(/(?:<w:r>(?:<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:tab\/><\/w:r>\s*)+<\/w:p>$/, '</w:p>');
    body = fixRunProps(body, st, sp);
    body = mergeRuns(body);
    body = splitNumbers(body);
    const ind = mode === 's556' ? 2552 : 0;
    let p = `<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/><w:ind w:start="${ind}" w:end="0"/><w:jc w:val="${jc}"/>${markRpr}</w:pPr>` + body;
    const t = plain(p).replace(/\s+$/, '');
    if (wantDashes(t, jc)) p = addDashes(p, mode === 's556' ? 13990 : 6690);
    return p;
  }

  // ---------- tables -> logical rows ----------
  function tableRows(tblXml, items) {
    const bidi = /<w:tblPr>[\s\S]*?<w:bidiVisual\/>[\s\S]*?<\/w:tblPr>/.test(tblXml.slice(0, 2000)) && tblXml.slice(0, 2000).indexOf('<w:bidiVisual/>') < tblXml.slice(0, 2000).indexOf('</w:tblPr>');
    const grid = [...(tblXml.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => +m[1]);
    const gx = [0]; grid.forEach((g) => gx.push(gx[gx.length - 1] + g));
    const total = gx[gx.length - 1] || 1;
    const inner = tblXml.slice(tblXml.indexOf('>') + 1, tblXml.lastIndexOf('</w:tbl>'));
    for (const r of children(inner, 'tr')) {
      const cellsX = children(r.slice(r.indexOf('>') + 1, r.lastIndexOf('</w:tr>')), 'tc');
      if (/<w:tbl>/.test(r)) {
        // container row: its own text first (titles), then nested tables
        for (const c of cellsX) {
          for (const b of topBlocks(c.slice(c.indexOf('</w:tcPr>') >= 0 ? c.indexOf('</w:tcPr>') + 9 : 6, c.lastIndexOf('</w:tc>')))) {
            if (b.t === 'tbl') tableRows(b.x, items);
            else { const t = plain(b.x).trim(); if (t) items.push({ t: 'title', text: t }); }
          }
        }
        continue;
      }
      let col = +((r.match(/<w:gridBefore w:val="(\d+)"/) || [])[1] || 0);
      const cells = [];
      for (const c of cellsX) {
        const span = +((c.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1);
        const vm = c.match(/<w:vMerge(?: w:val="(\w+)")?\/>/);
        const paras = c.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
        const text = paras.map((p) => plain(p).replace(/\s+/g, ' ').trim()).filter((s) => s).join(' ');
        let a = gx[Math.min(col, grid.length)] ?? total, b = gx[Math.min(col + span, grid.length)] ?? total;
        if (!bidi) { const a2 = total - b, b2 = total - a; a = a2; b = b2; }   // measure from the right edge
        cells.push({ text, a: a / total, b: b / total, cont: !!(vm && vm[1] !== 'restart') });
        col += span;
      }
      cells.sort((p, q) => p.a - q.a);
      items.push({ t: 'row', cells });
    }
  }
  const isNum = (s) => /^[\d\s.,\/\-]+$/.test(s) && /\d/.test(s);
  const rowKind = (cells) => {
    const ne = cells.filter((c) => c.text);
    const all = ne.map((c) => c.text).join(' ');
    if (/المجموع|مجموع/.test(all) && ne.some((c) => isNum(c.text))) return 'total';
    if (ne.length >= 3 && ne.filter((c) => isNum(c.text)).length === 0) return 'head';
    if (ne.length <= 2 && !ne.some((c) => isNum(c.text) && ne.length === 1)) return ne.length ? 'title' : 'empty';
    return 'data';
  };
  function buildHeader(headRows) {
    // columns = cells of the last header row; a wider cell above adds its text on top
    const last = headRows[headRows.length - 1];
    const cols = last.cells.map((c) => ({ a: c.a, b: c.b, text: c.text }));
    for (let k = headRows.length - 2; k >= 0; k--) {
      for (const up of headRows[k].cells) {
        if (!up.text) continue;
        const covered = cols.filter((c) => c.a >= up.a - 0.005 && c.b <= up.b + 0.005);
        if (covered.length === 1 && !covered[0].text) covered[0].text = up.text;
        else if (covered.length > 1) covered.forEach((c) => { if (!c.text.includes(up.text)) c.text = up.text + ' ' + c.text; });
      }
    }
    // columns that are only spacing (no text, very narrow) are dropped
    return cols.filter((c) => c.text || c.b - c.a > 0.03);
  }
  function mapRow(cells, cols, prev) {
    return cols.map((col, i) => {
      let best = null, bo = 0;
      for (const c of cells) {
        const o = Math.min(c.b, col.b) - Math.max(c.a, col.a);
        if (o > bo + 1e-6) { bo = o; best = c; }
      }
      if (!best) return '';
      if (best.cont && !best.text && prev) return prev[i];
      return best.text;
    });
  }

  // ---------- table output ----------
  function runsFor(t, sz, bold) {
    const base = RF + (bold ? '<w:b/><w:bCs/>' : '') + `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
    const re = new RegExp('(' + NUM.source + '|' + LATIN.source + ')');
    return t.split(re).filter((s) => s).map((s) => {
      const ltr = new RegExp('^(?:' + NUM.source + '|' + LATIN.source + ')$').test(s);
      return `<w:r><w:rPr>${base}${ltr ? '' : '<w:rtl/>'}</w:rPr><w:t xml:space="preserve">${esc(s)}</w:t></w:r>`;
    }).join('');
  }
  function titlePara(text, mode) {
    const ind = mode === 's556' ? 2552 : 0;
    return `<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/><w:ind w:start="${ind}" w:end="0"/><w:jc w:val="center"/></w:pPr>${runsFor(text, 26, true)}</w:p>`;
  }
  function colWidths(cols, W) {
    const wts = cols.map((c) => /طبيعة|التعيين|وصف|الوصف|البيان|تتكون|تشتمل/.test(c.text) ? 4 : /المساحة|مساحة/.test(c.text) ? 1.45 : /الطابق|المستوى/.test(c.text) ? 1.4 : /المشترك|الأجزاء|الاجزاء|نسبة/.test(c.text) ? 1.15 : 1);
    const s = wts.reduce((a, b) => a + b, 0);
    const w = wts.map((x) => Math.floor((W * x) / s));
    w[w.length - 1] += W - w.reduce((a, b) => a + b, 0);
    return w;
  }
  function makeTable(head, rows, mode, showHead) {
    const W = mode === 's556' ? 13000 : 6600;
    const cols = head.cols; const w = colWidths(cols, W);
    const areaIdx = cols.map((c, i) => (/المساحة|مساحة/.test(c.text) ? i : -1)).filter((i) => i >= 0);
    const descIdx = cols.findIndex((c) => /طبيعة|التعيين|وصف|الوصف|البيان/.test(c.text));
    const S = mode === 's556'
      ? { body: 24, head: 22, mar: '<w:tcMar><w:top w:w="150" w:type="dxa"/><w:start w:w="80" w:type="dxa"/><w:bottom w:w="150" w:type="dxa"/><w:end w:w="80" w:type="dxa"/></w:tcMar>', shade: true }
      : { body: 15, head: 15, mar: '', shade: false };
    const BORD = '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:start w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:end w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>';
    const cell = (width, text, opt = {}) => {
      const pr = `<w:tcW w:w="${width}" w:type="dxa"/>` + (opt.span > 1 ? `<w:gridSpan w:val="${opt.span}"/>` : '') + BORD +
        (opt.shade && S.shade ? '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' : '') + S.mar + '<w:vAlign w:val="center"/>';
      const sz = opt.bold ? S.head : S.body;
      return `<w:tc><w:tcPr>${pr}</w:tcPr><w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:start="0" w:end="0"/><w:jc w:val="${opt.jc || 'center'}"/></w:pPr>${text ? runsFor(text, sz, opt.bold) : ''}</w:p></w:tc>`;
    };
    const TR = '<w:tr><w:trPr><w:cantSplit/><w:jc w:val="center"/></w:trPr>';
    let out = `<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="${W}" w:type="dxa"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:start w:w="30" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:end w:w="30" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${w.map((x) => `<w:gridCol w:w="${x}"/>`).join('')}</w:tblGrid>`;
    if (showHead) out += TR + cols.map((c, i) => cell(w[i], c.text, { bold: true, shade: true })).join('') + '</w:tr>';
    const addM2 = (v) => (v && isNum(v) && !/م/.test(v) ? v + ' م²' : v);
    for (const r of rows) {
      if (r.kind === 'total') {
        const vals = r.vals;
        let first = vals.findIndex((v) => isNum(v)); if (first < 1) first = Math.max(1, cols.length - 4);
        const label = vals.slice(0, first).filter((v) => v && !isNum(v)).join(' ') || 'المجموع';
        out += TR + cell(w.slice(0, first).reduce((a, b) => a + b, 0), label, { span: first, bold: true, shade: true }) +
          vals.slice(first).map((v, k) => cell(w[first + k], areaIdx.includes(first + k) ? addM2(v) : v, { bold: true, shade: true })).join('') + '</w:tr>';
      } else {
        out += TR + r.vals.map((v, i) => cell(w[i], areaIdx.includes(i) ? addM2(v) : v, { jc: i === descIdx ? 'both' : 'center' })).join('') + '</w:tr>';
      }
    }
    return out + '</w:tbl>';
  }
  // turn the logical items (titles / header rows / data rows) into output xml
  function renderTableItems(items, mode, state) {
    let out = ''; let rows = []; let headRows = []; let showHead = false;
    const flush = () => {
      if (rows.length && state.head) out += makeTable(state.head, rows, mode, showHead);
      else if (rows.length) {  // no header known: plain table with the row cells as they are
        const n = Math.max(...rows.map((r) => r.vals.length));
        const cols = Array.from({ length: n }, () => ({ text: '' }));
        out += makeTable({ cols }, rows.map((r) => ({ ...r, vals: [...r.vals, ...Array(n - r.vals.length).fill('')] })), mode, false);
      }
      rows = []; showHead = false;
    };
    for (const it of items) {
      if (it.t === 'title') { flush(); out += titlePara(it.text, mode); continue; }
      const kind = rowKind(it.cells);
      if (kind === 'empty') continue;
      if (kind === 'head') {
        if (rows.length) flush();
        headRows.push(it); state.head = { cols: buildHeader(headRows) }; showHead = true; state.prev = null; continue;
      }
      headRows = [];
      if (kind === 'title') {
        flush();
        const t = it.cells.filter((c) => c.text).map((c) => c.text).join(' ');
        // "القسم 27 مجموعة الملكية 155العمارة 92" -> two title lines
        const m = t.match(/^(.*?\d)\s*(العمارة\s*\d.*|المدخل\s*\d.*|البرج\s*\d.*)$/);
        if (m && m[1].length > 3) { out += titlePara(m[1].trim(), mode) + titlePara(m[2].trim(), mode); } else out += titlePara(t, mode);
        continue;
      }
      const vals = state.head ? mapRow(it.cells, state.head.cols, state.prev) : it.cells.map((c) => c.text);
      if (kind === 'data') state.prev = vals;
      rows.push({ kind, vals });
      if (kind === 'total') { flush(); state.prev = null; }
    }
    flush();
    return out;
  }

  // ---------- detection of the cadastral box values ----------
  function detectBox(text) {
    const uniq = (arr) => [...new Set(arr)];
    const qism = uniq([...text.matchAll(/(?:^|[\s(])(?:ال)?قسم\s*(?:رقم)?\s*[:：]?\s*(\d{1,4})(?!\d)/g)].map((m) => +m[1] + '')).map((v) => v.padStart(2, '0'));
    const maj = uniq([...text.matchAll(/مجموع[ةه]\s+(?:ال)?ملكي[ةه]\s*(?:رقم)?\s*[:：]?\s*(\d{1,5})/g)].map((m) => m[1]));
    const freq = (re, stop) => {
      const c = {};
      for (const m of text.matchAll(re)) {
        const words = []; for (const w of m[1].split(/\s+/)) { if (!w || stop.test(w)) break; words.push(w); if (words.length >= 2) break; }
        const k = words.join(' ').replace(/^و(?=ال)/, '');
        const near = /قسم|مجموع[ةه] (?:ال)?ملكي|مسح|المسمى|واقع/.test(text.slice(Math.max(0, m.index - 250), m.index + 250));
        if (k) c[k] = (c[k] || 0) + (near ? 5 : 1);
      }
      return Object.entries(c).sort((x, y) => y[1] - x[1])[0]?.[0] || '';
    };
    const STOP = /^(بتاري|دائرة|ولاي|وولاي|الكائن|رقم|قسم|حي|مسجل|المؤرخ|الموافق)/;
    const bal = freq(/بلدي[ةه]\s*[:：]?\s*([^،,.:()\-0-9\n]{2,60})/g, STOP);
    const wil = freq(/ولاي[ةه]\s*[:：]?\s*([^،,.:()\-0-9\n]{2,40})/g, /^(بتاري|دائرة|بلدي|الكائن|رقم|قسم|حي|المؤرخ|تحت|المسجل)/);
    return {
      qism: qism.slice(0, 6).join(' و '),
      majmoua: maj.slice(0, 20).join('، '),
      baladiya: bal, wilaya: wil,
    };
  }

  // ---------- main ----------
  async function analyse(JSZip, buf) {
    const zip = await JSZip.loadAsync(buf);
    const f = zip.file('word/document.xml');
    if (!f) throw new Error('NOT_DOCX');
    let doc = await f.async('string');
    const stylesXml = zip.file('word/styles.xml') ? await zip.file('word/styles.xml').async('string') : '';
    const themeF = Object.keys(zip.files).find((n) => /^word\/theme\/theme\d*\.xml$/.test(n));
    const themeXml = themeF ? await zip.file(themeF).async('string') : '';
    let body = doc.slice(doc.indexOf('<w:body>') + 8, doc.lastIndexOf('<w:sectPr') > 0 ? doc.lastIndexOf('<w:sectPr') : doc.lastIndexOf('</w:body>'));
    let boxed = false;
    // a file already in the box template: take only what is inside the big box
    const bh = [...doc.matchAll(/<w:trHeight w:val="(2000|9300)"/g)];
    if (bh.length && /إطار مخصص للمحافظ/.test(stripTags(doc))) {
      const rowI = bh[0].index;
      const tc1 = doc.indexOf('<w:tc>', rowI), tc2 = doc.indexOf('<w:tc>', tc1 + 5);
      const s = doc.indexOf('</w:tcPr>', tc2) + 9, e = doc.lastIndexOf('</w:tc></w:tr></w:tbl>');
      if (tc2 > 0 && e > s) { body = doc.slice(s, e); boxed = true; }
    }
    const blocks = topBlocks(body);
    const text = blocks.map((b) => plain(b.x)).join('\n');
    return { zip, doc, stylesXml, themeXml, blocks, boxed, box: detectBox(text), text };
  }


  // the file's own fonts (theme fonts written out by name), used for runs in kept tables
  function readFonts(stylesXml, themeXml) {
    const theme = {};
    for (const k of ['major', 'minor']) {
      const blk = (themeXml.match(new RegExp(`<a:${k}Font>[\\s\\S]*?</a:${k}Font>`)) || [''])[0];
      theme[k + 'Latin'] = (blk.match(/<a:latin typeface="([^"]*)"/) || [])[1] || '';
      theme[k + 'Cs'] = (blk.match(/<a:cs typeface="([^"]*)"/) || [])[1] || '';
    }
    const res = (name) => {
      if (!name) return '';
      const k = name.startsWith('major') ? 'major' : 'minor';
      return /Bidi/.test(name) ? (theme[k + 'Cs'] || theme[k + 'Latin']) : theme[k + 'Latin'];
    };
    const dd = ((stylesXml || '').match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/) || [''])[0];
    const rf = (dd.match(/<w:rFonts [^>]*\/>/) || [''])[0];
    const at = (a) => (rf.match(new RegExp(`w:${a}="([^"]*)"`)) || [])[1];
    const ascii = at('ascii') || res(at('asciiTheme')) || 'Times New Roman';
    const cs = at('cs') || res(at('cstheme')) || ascii;
    return { ascii, cs, res };
  }
  function explicitFonts(r, fonts) {
    // <w:rFonts> with names only (no theme references), so the table looks the same in the new file
    const m = r.match(/<w:rFonts [^>]*\/>/);
    const at = (a) => m && (m[0].match(new RegExp(`w:${a}="([^"]*)"`)) || [])[1];
    const ascii = at('ascii') || fonts.res(at('asciiTheme')) || fonts.ascii;
    const hAnsi = at('hAnsi') || fonts.res(at('hAnsiTheme')) || ascii;
    const cs = at('cs') || fonts.res(at('cstheme')) || fonts.cs;
    const el = `<w:rFonts w:ascii="${esc(ascii)}" w:hAnsi="${esc(hAnsi)}" w:cs="${esc(cs)}"/>`;
    return m ? r.replace(m[0], el) : r.replace('<w:rPr>', '<w:rPr>' + el);
  }

  // ---------- tables kept exactly as in the file; only made narrower when wider than the page/box ----------
  function keepTable(x, W, st) {
    x = cleanXml(x).replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, '');
    // borders/cell margins that came from the file's table style are written into the table itself
    const ts = (x.match(/<w:tblStyle w:val="([^"]+)"/) || [])[1];
    if (ts) {
      let id = ts, raw = '', g = 0;
      while (id && st.styles[id] && g++ < 10) { raw = st.styles[id].raw + raw; id = st.styles[id].based; }
      const bord = (raw.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/) || [])[0];
      const mar = (raw.match(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/) || [])[0];
      x = x.replace(/<w:tblPr>([\s\S]*?)<\/w:tblPr>/, (m, inner) => {
        inner = inner.replace(/<w:tblStyle [^>]*\/>/, '');
        const ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook'];
        const add = (name, el) => {
          if (!el || inner.includes('<w:' + name)) return;
          const k = ORDER.indexOf(name);
          const after = ORDER.slice(k + 1).map((n) => inner.indexOf('<w:' + n)).filter((i) => i >= 0);
          const pos = after.length ? Math.min(...after) : inner.length;
          inner = inner.slice(0, pos) + el + inner.slice(pos);
        };
        add('tblBorders', bord); add('tblCellMar', mar);
        return '<w:tblPr>' + inner + '</w:tblPr>';
      });
    }
    // paragraphs in cells: alignment / size that came from the file's styles
    x = x.replace(/<w:p(?: [^>]*)?>(<w:pPr>[\s\S]*?<\/w:pPr>)?([\s\S]*?)<\/w:p>/g, (m, ppr, rest) => {
      const id = ((ppr || '').match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || 'Normal';
      const sp = styleProps(st, id);
      let p = ppr || '<w:pPr></w:pPr>';
      p = p.replace(/<w:pStyle [^>]*\/>/, '');
      if (sp.jc && !/<w:jc /.test(p)) p = p.replace(/(<w:rPr>|<\/w:pPr>)/, `<w:jc w:val="${sp.jc}"/>$1`);
      const def = sp.sz || st.defaults.sz;
      rest = rest.replace(/<w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?/g, (q, r) => {
        r = r || '<w:rPr></w:rPr>';
        r = orderRPr(explicitFonts(r, st.fonts));
        if (!def) return '<w:r>' + r;
        if (!/<w:sz /.test(r)) r = orderRPr(r.replace('</w:rPr>', `<w:sz w:val="${def}"/><w:szCs w:val="${def}"/></w:rPr>`));
        if (sp.b && !/<w:b[ \/]/.test(r)) r = orderRPr(r.replace('</w:rPr>', '<w:b/><w:bCs/></w:rPr>'));
        return '<w:r>' + r;
      });
      return '<w:p>' + p + rest + '</w:p>';
    });
    // width: the real width is the largest of the grid, the table width and the widest row of cells
    const dxa = (re) => [...x.matchAll(re)].map((m) => +m[1]);
    const gridBlock = (x.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0];
    let grid = [...gridBlock.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => +m[1]);
    const outerRows = children(x.slice(x.indexOf('>') + 1, x.lastIndexOf('</w:tbl>')), 'tr');
    let rowMax = 0, bestRow = null;
    for (const r of outerRows) {
      const cs = children(r.slice(r.indexOf('>') + 1, r.lastIndexOf('</w:tr>')), 'tc');
      const ws = cs.map((c) => +((c.match(/<w:tcW (?:w:w="(\d+)" w:type="dxa"|w:type="dxa" w:w="(\d+)")/) || []).slice(1).find((v) => v) || 0));
      const sum = ws.reduce((p, q) => p + q, 0);
      if (sum > rowMax) rowMax = sum;
      if (!bestRow && cs.length === grid.length && ws.every((w) => w > 0) && !cs.some((c) => /<w:gridSpan/.test(c))) bestRow = ws;
    }
    const gsum = grid.reduce((p, q) => p + q, 0);
    // a grid that does not match the cells (files saved by other programs): take the widths from the cells
    if (bestRow && Math.abs(bestRow.reduce((p, q) => p + q, 0) - gsum) > gsum * 0.05) {
      grid = bestRow;
      x = x.replace(gridBlock, '<w:tblGrid>' + grid.map((g) => `<w:gridCol w:w="${g}"/>`).join('') + '</w:tblGrid>');
    }
    const tblW = (x.slice(0, x.indexOf('</w:tblPr>')).match(/<w:tblW (?:w:w="(\d+)" w:type="dxa"|w:type="dxa" w:w="(\d+)")/) || []).slice(1).find((v) => v);
    const tw = Math.max(grid.reduce((p, q) => p + q, 0), +(tblW || 0), rowMax);
    const f = tw > W ? W / tw : 1;
    if (f < 1) {
      const sc = (v) => Math.max(1, Math.round(+v * f));
      x = x.replace(/<w:gridCol w:w="(\d+)"/g, (m, v) => `<w:gridCol w:w="${sc(v)}"`)
        .replace(/<w:(tcW|tblW|wBefore|wAfter) w:w="(\d+)" w:type="dxa"\/>/g, (m, n, v) => `<w:${n} w:w="${sc(v)}" w:type="dxa"/>`)
        .replace(/<w:(tcW|tblW|wBefore|wAfter) w:type="dxa" w:w="(\d+)"\/>/g, (m, n, v) => `<w:${n} w:type="dxa" w:w="${sc(v)}"/>`)
        // the same table, just smaller: text and cell padding shrink with the columns so words do not break
        .replace(/<w:(sz|szCs) w:val="(\d+)"\/>/g, (m, n, v) => `<w:${n} w:val="${Math.max(14, Math.round(+v * f))}"/>`)
        .replace(/(<w:(?:tcMar|tblCellMar)>[\s\S]*?<\/w:(?:tcMar|tblCellMar)>)/g, (m) => m.replace(/w:w="(\d+)"/g, (q, v) => `w:w="${sc(v)}"`));
    }
    // widths are fixed so Word keeps the columns as they are instead of re-fitting them
    x = x.replace(/<w:tblLayout [^>]*\/>/, '');
    x = x.replace(/<w:tblPr>([\s\S]*?)<\/w:tblPr>/, (m, inner) => {
      const after = ['tblCellMar', 'tblLook', 'tblCaption', 'tblDescription'].map((n) => inner.indexOf('<w:' + n)).filter((i) => i >= 0);
      const pos = after.length ? Math.min(...after) : inner.length;
      return '<w:tblPr>' + inner.slice(0, pos) + '<w:tblLayout w:type="fixed"/>' + inner.slice(pos) + '</w:tblPr>';
    });
    x = x.replace(/<w:tblInd w:w="-?\d+" w:type="dxa"\/>/, '<w:tblInd w:w="0" w:type="dxa"/>');
    x = x.replace(/<w:tblpPr [^>]*\/>/g, '');   // floating tables stay in the flow
    return x;
  }

  function buildContent(a, mode) {
    const st = readStyles(a.stylesXml); st.fonts = readFonts(a.stylesXml, a.themeXml);
    let out = ''; const stats = { paras: 0, tables: 0, rows: 0, dashes: 0 };
    const state = { head: null, prev: null };
    let pendingItems = [];
    const flushItems = () => { if (pendingItems.length) { out += renderTableItems(pendingItems, mode, state); pendingItems = []; } };
    const W = mode === 's556' ? 14004 : 6698;
    const gap = '<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
    let lastWasTable = false;
    for (const b of a.blocks) {
      if (b.t === 'tbl') {
        if (lastWasTable) out += gap;       // two tables never stick together
        out += keepTable(b.x, W, st);
        stats.tables++; stats.rows += (b.x.match(/<w:tr[ >]/g) || []).length;
        lastWasTable = true;
        continue;
      }
      const t = plain(b.x).trim();
      if (!t || /^[\s\-_ـ.:]*$/.test(t)) continue;   // no empty lines, and no lines made only of dashes
      // a short centred title between two tables keeps the table context (header shown once per building)
      flushItems();
      const p = convertParagraph(b.x, st, mode);
      if (p.includes('w:leader="hyphen"')) stats.dashes++;
      out += p; stats.paras++; lastWasTable = false;
    }
    flushItems();
    return { xml: out, stats };
  }

  async function zipFrom(JSZip, files, replaceDoc) {
    const z = new JSZip();
    for (const [name, content] of Object.entries(files)) z.file(name, name === 'word/document.xml' ? replaceDoc : content);
    return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function convert(JSZip, a, shells, box) {
    const endP = '<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
    const t = buildContent(a, 'tpl');
    const v = (s, d) => esc((s || '').trim()) || d;
    let tdoc = shells.tpl['word/document.xml']
      .replace('{{CONTENT}}', t.xml + endP)
      .replace('{{QISM}}', v(box.qism, '')).replace('{{BALADIYA}}', v(box.baladiya, '')).replace('{{MAJMOUA}}', v(box.majmoua, ''))
      .replace('{{HISSA}}', v(box.hissa, '')).replace('{{WILAYA}}', v(box.wilaya, '..............'))
      // إيداع / في / مجلد / رسم: the value when typed, otherwise the dotted line of the form
      .replace('{{IDA_HAJM}}', v(box.idaHajm, '..................')).replace('{{IDA_RAQM}}', v(box.idaRaqm, '..................'))
      .replace('{{FI}}', v(box.fi, '..............................................'))
      .replace('{{MUJALLAD}}', v(box.mujallad, '...................')).replace('{{MUJ_RAQM}}', v(box.mujRaqm, '...................'))
      .replace('{{RASM}}', v(box.rasm, ''));
    const s = shells.s556 ? buildContent(a, 's556') : null;
    const sdoc = s ? shells.s556['word/document.xml'].replace('{{BODY}}', s.xml + endP) : null;
    return {
      template: await zipFrom(JSZip, shells.tpl, tdoc),
      s556: s ? await zipFrom(JSZip, shells.s556, sdoc) : null,
      stats: t.stats,
    };
  }

  root.EDD = { analyse, convert, detectBox };
})(typeof window !== 'undefined' ? window : globalThis);
