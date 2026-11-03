/* Page splitter for «إجراء إشهار عقاري» (text written on the page inside the drawn box).
   Measures every paragraph and table with the browser (canvas, Arial), decides where each page ends,
   writes «الصـفحـــة الأولى .../...» as normal text at the end of each page and starts the next page;
   the last page gets «ـ الصـفحـــة … والأخـيــرة ـ» before «الموثقة». */
(function (root) {
  'use strict';
  const G = {
    width: (9706 - 2882 - 120) / 20,      // text column, pt
    firstTop: 309.2 + 2, bottom: 841.9 - 130, otherTop: 35,
    safety: 1.05,                            // measured heights are increased a little: a page is never overfilled
    labelH: 13 * 1.15 * 1.149 + 4,           // the page line itself
    font: 'Arial, "Traditional Arabic", sans-serif',   // the font Word uses for the deed
  };
  const IND = '<w:ind w:start="2882" w:end="120"/>';
  const TOP = { s: 2882 / 20, e: 6 };
  let ctx = null;
  const textW = (s, pt, bold) => {
    if (!ctx) ctx = document.createElement('canvas').getContext('2d');
    ctx.font = `${bold ? 'bold ' : ''}${pt}px ${G.font}`;
    return ctx.measureText(s).width;
  };
  const dec = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  function topBlocks(body) {
    const out = []; const re = /<(\/?)w:(tbl|p)(?=[ >\/])[^>]*?(\/?)>/g; let depth = 0, st = 0, m;
    while ((m = re.exec(body))) {
      const close = m[1], tag = m[2], self = m[3];
      if (tag === 'tbl') { if (!close) { if (depth === 0) st = m.index; depth++; } else { depth--; if (depth === 0) out.push({ t: 'tbl', x: body.slice(st, re.lastIndex) }); } }
      else if (depth === 0) { if (!close && self) out.push({ t: 'p', x: m[0] }); else if (!close) st = m.index; else out.push({ t: 'p', x: body.slice(st, re.lastIndex) }); }
    }
    return out;
  }
  function children(x, tag) {
    const re = new RegExp(`<w:tbl>|</w:tbl>|<w:${tag}(?=[ >])[^>]*>|</w:${tag}>`, 'g');
    const out = []; let td = 0, st = -1, m;
    while ((m = re.exec(x))) {
      const s = m[0];
      if (s === '<w:tbl>') td++; else if (s === '</w:tbl>') td--;
      else if (td === 0) { if (s.startsWith('</')) { if (st >= 0) out.push(x.slice(st, re.lastIndex)); st = -1; } else st = m.index; }
    }
    return out;
  }
  function ordinalAr(n, compound) {
    const U = ['', 'الأولى', 'الثانية', 'الثالثة', 'الرابعة', 'الخامسة', 'السادسة', 'السابعة', 'الثامنة', 'التاسعة', 'العاشرة'];
    const UC = ['', 'الحادية', 'الثانية', 'الثالثة', 'الرابعة', 'الخامسة', 'السادسة', 'السابعة', 'الثامنة', 'التاسعة'];
    const T = { 2: 'العشرون', 3: 'الثلاثون', 4: 'الأربعون', 5: 'الخمسون', 6: 'الستون', 7: 'السبعون', 8: 'الثمانون', 9: 'التسعون' };
    if (n <= 10) return compound && n === 1 ? UC[1] : U[n];
    if (n < 20) return UC[n - 10] + ' عشرة';
    if (n < 100) { const t = Math.floor(n / 10), u = n % 10; return u ? UC[u] + ' و' + T[t] : T[t]; }
    if (n === 100) return 'المائة';
    if (n < 200) return ordinalAr(n - 100, true) + ' بعد المائة';
    if (n === 200) return 'المائتان';
    return String(n);
  }

  // ---------- paragraph: runs, wrapping into lines ----------
  function paraInfo(p, width, base) {
    base = base || { s: 0, e: 0 };
    const ppr = (p.match(/^<w:p(?: [^>]*)?>(<w:pPr>[\s\S]*?<\/w:pPr>)?/) || [])[1] || '';
    const sp = ppr.match(/<w:spacing ([^>]*)\/>/); const a = (k) => { const m = sp && sp[1].match(new RegExp(`w:${k}="(-?\\d+)"`)); return m ? +m[1] : null; };
    const before = (a('before') || 0) / 20, after = (a('after') || 0) / 20;
    const line = a('line'), rule = sp && (sp[1].match(/w:lineRule="(\w+)"/) || [])[1] || 'auto';
    const ind = ppr.match(/<w:ind ([^>]*)\/>/);
    const iv = (k) => { const m = ind && ind[1].match(new RegExp(`w:${k}="(-?\\d+)"`)); return m ? +m[1] / 20 : 0; };
    const avail = Math.max(40, width - Math.max(0, iv('start') - base.s) - Math.max(0, iv('end') - base.e));
    // tokens with their run font
    const toks = []; let maxSz = 0, offset = 0;
    const body = p.slice((p.match(/^<w:p(?: [^>]*)?>(<w:pPr>[\s\S]*?<\/w:pPr>)?/) || [''])[0].length);
    for (const r of body.match(/<w:r>[\s\S]*?<\/w:r>|<w:r [^>]*>[\s\S]*?<\/w:r>/g) || []) {
      const sz = +((r.match(/<w:sz w:val="(\d+)"/) || [])[1] || 24) / 2; const bold = /<w:b\/>|<w:b w:val="(1|true)"/.test(r);
      if (/<w:br\/>|<w:br w:type="textWrapping"\/>/.test(r)) toks.push({ br: true, off: offset });
      for (const m of r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)) {
        const s = dec(m[1]); if (!s) continue; maxSz = Math.max(maxSz, sz);
        for (const w of s.split(/(?<= )/)) { if (w) { toks.push({ w, sz, bold, off: offset }); offset += w.length; } }
      }
    }
    if (!maxSz) maxSz = 12;
    const lh = rule === 'exact' ? line / 20 : rule === 'atLeast' ? Math.max(line / 20, maxSz * 1.149) : ((line || 240) / 240) * maxSz * 1.149;
    // greedy wrapping
    const lines = []; let cur = 0, start = 0;
    for (const t of toks) {
      if (t.br) { lines.push({ end: t.off, w: cur }); cur = 0; start = t.off; continue; }
      const w = textW(t.w, t.sz, t.bold);
      if (cur > 0 && cur + textW(t.w.trimEnd(), t.sz, t.bold) > avail) { lines.push({ end: t.off, w: cur }); cur = 0; start = t.off; }
      cur += w;
    }
    lines.push({ end: offset, w: cur });
    let n = lines.length;
    if (p.includes('w:leader="hyphen"') && lines[n - 1].w > avail * 0.92) n++;   // the dash leader needs room at the end
    return { lines: n, lineEnds: lines.map((l) => l.end), lh, before, after, height: before + after + n * lh };
  }
  // split a paragraph after `chars` characters of its text (between two words)
  function splitPara(p, chars) {
    const head = (p.match(/^<w:p(?: [^>]*)?>(<w:pPr>[\s\S]*?<\/w:pPr>)?/) || [''])[0];
    const body = p.slice(head.length, p.lastIndexOf('</w:p>'));
    const runs = body.match(/<w:r>[\s\S]*?<\/w:r>|<w:r [^>]*>[\s\S]*?<\/w:r>|<w:[a-zA-Z]+[^>]*\/>|<w:(?!r[ >])([a-zA-Z]+)[^>]*>[\s\S]*?<\/w:\1>/g) || [];
    let a = '', b = '', count = 0, done = false;
    for (const r of runs) {
      if (done) { b += r; continue; }
      const texts = [...r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)];
      const len = texts.reduce((s, m) => s + dec(m[1]).length, 0);
      if (count + len <= chars) { a += r; count += len; if (count === chars) done = true; continue; }
      // split inside this run (one <w:t> per run after cleaning)
      const m = texts[0]; const s = dec(m[1]); const k = chars - count;
      const enc = (x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      a += r.replace(m[0], `<w:t xml:space="preserve">${enc(s.slice(0, k))}</w:t>`);
      b += r.replace(m[0], `<w:t xml:space="preserve">${enc(s.slice(k))}</w:t>`);
      done = true;
    }
    // first part: no dash leader (it continues on the next page)
    const headA = head.replace(/<w:tabs>[\s\S]*?<\/w:tabs>/, '').replace('<w:jc w:val="both"/>', '<w:jc w:val="both"/>');
    return [headA + a + '</w:p>', head + b + '</w:p>'];
  }

  // ---------- tables: row heights ----------
  function tableRows(t) {
    const inner = t.slice(t.indexOf('>') + 1, t.lastIndexOf('</w:tbl>'));
    const pre = inner.slice(0, inner.search(/<w:tr[ >]/));
    const rows = children(inner, 'tr');
    const tblMar = (pre.match(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/) || [''])[0];
    const mv = (blk, side, d) => { const m = blk.match(new RegExp(`<w:${side} w:w="(\\d+)"`)); return m ? +m[1] / 20 : d; };
    const dL = mv(tblMar, 'left', mv(tblMar, 'start', 5.4)), dR = mv(tblMar, 'right', mv(tblMar, 'end', 5.4));
    const dT = mv(tblMar, 'top', 0), dB = mv(tblMar, 'bottom', 0);
    return {
      pre,
      rows: rows.map((r) => {
        const trh = r.match(/<w:trHeight w:val="(\d+)"(?: w:hRule="(\w+)")?/);
        let h = 0;
        for (const c of children(r.slice(r.indexOf('>') + 1, r.lastIndexOf('</w:tr>')), 'tc')) {
          if (/<w:vMerge\/>|<w:vMerge w:val="continue"\/>/.test(c)) continue;
          const w = +((c.match(/<w:tcW w:w="(\d+)"/) || [])[1] || 1000) / 20;
          const tcm = (c.match(/<w:tcMar>[\s\S]*?<\/w:tcMar>/) || [''])[0];
          const cw = w - mv(tcm, 'left', mv(tcm, 'start', dL)) - mv(tcm, 'right', mv(tcm, 'end', dR));
          let ch = mv(tcm, 'top', dT) + mv(tcm, 'bottom', dB);
          for (const p of c.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) || []) ch += paraInfo(p, Math.max(20, cw)).height;
          h = Math.max(h, ch);
        }
        if (trh) { const v = +trh[1] / 20; h = trh[2] === 'exact' ? v : Math.max(h, v); }
        return { x: r, h: h + 1 };
      }),
    };
  }
  const restartMerges = (rowXml) => rowXml.replace(/<w:vMerge\/>|<w:vMerge w:val="continue"\/>/g, '<w:vMerge w:val="restart"/>');

  // ---------- page lines and breaks ----------
  const labelXml = (text) => `<w:p><w:pPr><w:keepLines/><w:bidi/><w:spacing w:before="60" w:after="0" w:line="276" w:lineRule="auto"/>${IND}<w:jc w:val="end"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/><w:rtl/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const breakPara = '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  const withBreak = (p) => p.replace(/^(<w:p(?: [^>]*)?>)(<w:pPr>)?/, (m, a, b) => a + '<w:pPr><w:pageBreakBefore/>' + (b ? '' : '</w:pPr>'));

  function run(xml) {
    const blocks = topBlocks(xml);
    // where the last-page line goes: the old automatic line (with NUMPAGES), or before «الموثقة», or at the end
    let lastAt = blocks.findIndex((b) => b.x.includes('NUMPAGES'));
    if (lastAt >= 0) blocks.splice(lastAt, 1);
    if (lastAt < 0) {
      const sig = blocks.map((b, i) => (b.t === 'p' && /^ال[ـ\s]*م[ـ\s]*و[ـ\s]*ث[ـ\s]*ق[ـ\s]*(ة|ـ*ة|)[ـ\s./]*$/.test(dec(b.x.replace(/<[^>]+>/g, '')).trim()) ? i : -1)).filter((i) => i >= 0).pop();
      lastAt = sig !== undefined && sig >= blocks.length - 4 ? sig : blocks.length;
    }
    const out = []; let page = 1, used = 0, cap = G.bottom - G.firstTop, breakNext = false;
    const usable = () => cap - G.labelH;
    const endPage = () => { out.push(labelXml(`الصـفحـــة ${ordinalAr(page)} .../...`)); page++; used = 0; cap = G.bottom - G.otherTop; breakNext = true; };
    const put = (x, h, isTable) => {
      if (breakNext) { out.push(isTable ? breakPara + x : (x.startsWith('<w:p') ? withBreak(x) : breakPara + x)); breakNext = false; }
      else out.push(x);
      used += h;
    };
    const tailH = (from) => blocks.slice(from).reduce((s, b) => s + (b.t === 'p' ? paraInfo(b.x, G.width, TOP).height * G.safety : tableRows(b.x).rows.reduce((q, r) => q + r.h, 0) * G.safety), 0);
    for (let i = 0; i <= blocks.length; i++) {
      if (i === lastAt) {
        // the last-page line and what follows it («الموثقة») stay together on the last page
        const need = G.labelH + tailH(i);
        if (used > 0 && used + need > cap) endPage();
        out.push('LAST_PAGE_LINE');
        used += G.labelH;
        if (i === blocks.length) break;
      }
      if (i === blocks.length) break;
      const b = blocks[i];
      if (b.t === 'p') {
        const info = paraInfo(b.x, G.width, TOP); const h = info.height * G.safety;
        if (used + h <= usable() || used === 0) { put(b.x, h, false); continue; }
        // split a long paragraph between two lines: at least 2 lines on each side
        const room = usable() - used - info.before * G.safety;
        const k = Math.floor(room / (info.lh * G.safety));
        if (info.lines >= 4 && k >= 2 && info.lines - k >= 2) {
          const [pa, pb] = splitPara(b.x, info.lineEnds[k - 1]);
          put(pa, (info.before + k * info.lh) * G.safety, false);
          endPage();
          const ib = paraInfo(pb, G.width, TOP);
          put(pb, ib.height * G.safety, false);
          continue;
        }
        endPage(); put(b.x, h, false);
        continue;
      }
      // table: rows stay whole; a table can continue on the next page
      const tr = tableRows(b.x);
      const head = b.x.slice(0, b.x.indexOf('>') + 1) + tr.pre;
      let chunk = [], chunkH = 0, first = true;
      const flush = () => { if (!chunk.length) return; put(head + chunk.join('') + '</w:tbl>', chunkH, true); chunk = []; chunkH = 0; };
      for (const r of tr.rows) {
        const h = r.h * G.safety;
        if (used + chunkH + h > usable() && (used + chunkH) > 0) { flush(); endPage(); first = false; chunk.push(restartMerges(r.x)); chunkH = h; continue; }
        chunk.push(first || chunk.length ? r.x : restartMerges(r.x)); chunkH += h;
      }
      flush();
    }
    const total = page;
    return out.join('').replace('LAST_PAGE_LINE', labelXml(`ـ الصـفحـــة ${ordinalAr(total)} والأخـيــرة ـ`)) ;
  }

  root.EDDPaginate = { run, G, paraInfo, TOP };
})(typeof window !== 'undefined' ? window : globalThis);
