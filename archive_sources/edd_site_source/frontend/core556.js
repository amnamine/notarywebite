/* 556 version: Word file whose tables are scanned pictures -> the same Word file with the
   pictures replaced by real typed tables (EDD_Ilot556 table style), landscape page.
   Reading the pictures is done by the page (Claude via the sample capability); this file
   finds the pictures, prepares the request and builds the Word file from the answers. */
(function (root) {
  'use strict';
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const stripTags = (x) => x.replace(/<[^>]+>/g, '');
  const NUM = /[\d?]*\d[\d?]*(?:[.,][\d?]+)*(?:\s?[\/\-]\s?[\d?]+(?:[.,][\d?]+)*)*/;
  const LATIN = /[A-Za-z][A-Za-z0-9.]*(?:\s?-\s?[0-9]+)?/;

  function topBlocks(body) {
    const out = []; const re = /<(\/?)w:(tbl|p)(?=[ >\/])[^>]*?(\/?)>/g; let depth = 0, st = 0, m;
    while ((m = re.exec(body))) {
      const close = m[1], tag = m[2], self = m[3];
      if (tag === 'tbl') {
        if (!close) { if (depth === 0) st = m.index; depth++; } else { depth--; if (depth === 0) out.push({ t: 'tbl', x: body.slice(st, re.lastIndex) }); }
      } else if (depth === 0) {
        if (!close && self) out.push({ t: 'p', x: m[0] });
        else if (!close) st = m.index;
        else out.push({ t: 'p', x: body.slice(st, re.lastIndex) });
      }
    }
    return out;
  }

  // ---------- 1. find the scanned pictures ----------
  async function analyse(JSZip, buf) {
    const zip = await JSZip.loadAsync(buf);
    const f = zip.file('word/document.xml'); if (!f) throw new Error('NOT_DOCX');
    const doc = await f.async('string');
    const relsXml = zip.file('word/_rels/document.xml.rels') ? await zip.file('word/_rels/document.xml.rels').async('string') : '';
    const rels = {};
    for (const m of relsXml.matchAll(/<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/?>/g)) rels[m[1]] = m[2];
    for (const m of relsXml.matchAll(/<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"[^>]*?\/?>/g)) rels[m[2]] = rels[m[2]] || m[1];
    const bStart = doc.indexOf('<w:body>') + 8;
    const lastSect = doc.lastIndexOf('<w:sectPr');
    const body = doc.slice(bStart, lastSect > 0 ? lastSect : doc.lastIndexOf('</w:body>'));
    const blocks = topBlocks(body);
    const images = [];
    blocks.forEach((b, i) => {
      if (b.t !== 'p') return;
      const ids = [...b.x.matchAll(/<a:blip [^>]*r:embed="([^"]+)"|<v:imagedata [^>]*r:id="([^"]+)"/g)].map((m) => m[1] || m[2]);
      for (const id of ids) {
        const target = rels[id]; if (!target) continue;
        const path = target.startsWith('/') ? target.slice(1) : 'word/' + target.replace(/^\.\//, '');
        const ext = (path.split('.').pop() || '').toLowerCase();
        images.push({ block: i, rid: id, path, ext, ok: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'].includes(ext) });
      }
    });
    return { zip, doc, body, bStart, blocks, images };
  }

  // ---------- 2. the request sent with each picture ----------
  function prompt(prev, nImages, pdfText) {
    return [
      'You are copying a scanned page from an Algerian notarial deed (Arabic, right-to-left) into structured data, so it can be rebuilt as an editable Word page that looks exactly like the scan.',
      nImages > 1
        ? `Image 1 is the WHOLE page: use it for the layout (which cells are merged, where each table starts and ends). Images 2 to ${nImages} are zoomed horizontal bands of the same page, top to bottom, overlapping a little: use them to read small text and digits exactly. Every row must appear ONCE.`
        : 'The image is the whole page.',
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
      '- "continues_previous": true only when this table has no heading rows because it continues a table from the previous page' + (prev ? ` (the previous table had ${prev.cols} columns; its headings were ${JSON.stringify(prev.head)}).` : '.'),
      '',
      pdfText ? 'The PDF page also carries this machine text layer (exact characters, but not in table order). When it contains a value you see in the image, use its exact characters:\n' + pdfText : '',
      'Reply with ONLY this JSON: {"blocks":[...],"unreadable":false}. Set "unreadable": true only if nothing on the page can be read.',
    ].filter((l) => l !== '').join('\n');
  }

  async function slices(blob, maxCount) {
    // Claude sees each image at about 1.2 megapixels: send the whole page (layout) and zoomed bands (small digits)
    let bmp;
    try { bmp = await createImageBitmap(blob); } catch (e) { return [blob]; }
    const w = bmp.width, h = bmp.height;
    const toPng = (y0, y1) => {
      const c = document.createElement('canvas'); c.width = w; c.height = y1 - y0;
      const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, w, c.height); g.drawImage(bmp, 0, y0, w, c.height, 0, 0, w, c.height);
      return new Promise((r) => c.toBlob(r, 'image/jpeg', 0.92));
    };
    const whole = /image\/(png|jpeg|webp|gif)/.test(blob.type) && blob.size < 15e6 ? blob : await toPng(0, h);
    let n = Math.ceil((w * h) / 1.3e6);
    n = Math.max(0, Math.min(n, 4, (maxCount || 1) - 1));
    if (n < 2) return [whole];
    const ov = Math.round(h * 0.04), step = Math.ceil(h / n);
    const out = [whole];
    for (let i = 0; i < n; i++) out.push(await toPng(Math.max(0, i * step - ov), Math.min(h, (i + 1) * step + ov)));
    return out;
  }

  // ---------- 3. a table as a grid (merged cells kept) ----------
  function toGrid(rows) {
    const grid = []; let ncols = 0;
    rows.forEach((row, r) => {
      grid[r] = grid[r] || [];
      let c = 0;
      for (const cell of row || []) {
        while (grid[r][c]) c++;
        const cs = Math.max(1, Math.min(30, parseInt(cell && cell.cs, 10) || 1)), rs = Math.max(1, Math.min(200, parseInt(cell && cell.rs, 10) || 1));
        const ref = { t: String((cell && (cell.t ?? cell.text)) ?? ''), b: !!(cell && cell.b), cs, rs, r, c };
        for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) {
          grid[r + i] = grid[r + i] || [];
          grid[r + i][c + j] = { ref, origin: i === 0 && j === 0, vcont: i > 0 && j === 0, hcont: j > 0 };
        }
        c += cs; ncols = Math.max(ncols, c);
      }
    });
    grid.length = Math.max(rows.length, 0);   // a rowspan cannot reach past the last row
    return { grid, ncols };
  }
  const num = (s) => {
    const t = String(s || '').replace(/م²|م2|م/g, '').replace(/\s/g, '').replace(',', '.');
    return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : null;
  };
  const isTotalRow = (row) => row.some((s) => s && s.origin && /المجموع|مجموع/.test(s.ref.t));
  function check(tb) {
    const issues = []; const { grid, ncols } = tb;
    let qs = 0; grid.forEach((row) => row.forEach((s) => { if (s && s.origin && s.ref.t.includes('?')) qs++; }));
    if (qs) issues.push(`${qs} خانة غير واضحة (مُعلَّمة بالأصفر)`);
    const seg = Array(ncols).fill(0), all = Array(ncols).fill(0), segN = Array(ncols).fill(0), allN = Array(ncols).fill(0);
    let subtotals = 0;
    grid.forEach((row, r) => {
      if (r < tb.headerRows) return;
      if (isTotalRow(row)) {
        const label = row.filter((s) => s && s.origin && num(s.ref.t) === null).map((s) => s.ref.t).join(' ');
        const grand = /الكلي|العام|الإجمالي/.test(label) || (subtotals > 0 && r === grid.length - 1 && !/الجزئي/.test(label));
        for (let c = 0; c < ncols; c++) {
          const s = row[c]; if (!s || !s.origin) continue;
          const t = num(s.ref.t); if (t === null) continue;
          const sum = grand ? all[c] : seg[c], n = grand ? allN[c] : segN[c];
          if (n < 2) continue;
          const intCol = Number.isInteger(t);
          if (Math.abs(sum - t) > (intCol ? 0.5 : Math.max(0.15, Math.abs(t) * 0.0005))) {
            const h = tb.grid.slice(0, tb.headerRows).map((hr) => hr[c] && hr[c].ref.t.split('\n')[0]).filter(Boolean).pop() || 'عمود ' + (c + 1);
            issues.push(`«${label.replace(/\s+/g, ' ').trim()}» في «${h}»: مجموع الأسطر ${String(+sum.toFixed(2)).replace('.', ',')} والمكتوب ${s.ref.t}`);
          }
        }
        if (!grand) { subtotals++; seg.fill(0); segN.fill(0); }
        return;
      }
      for (let c = 0; c < ncols; c++) {
        const s = row[c]; if (!s || !s.origin) continue;
        const v = num(s.ref.t); if (v === null) continue;
        seg[c] += v; all[c] += v; segN[c]++; allN[c]++;
      }
    });
    return issues;
  }

  // ---------- 4. Word xml in the EDD_Ilot556 table style ----------
  function runs(t, sz, bold, underline) {
    const base = `<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>` + (bold ? '<w:b/><w:bCs/>' : '') + `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` + (underline ? '<w:u w:val="single"/>' : '');
    const re = new RegExp('(' + NUM.source + '|' + LATIN.source + ')');
    const whole = new RegExp('^(?:' + NUM.source + '|' + LATIN.source + ')$');
    return String(t).split(re).filter((s) => s).map((s) => `<w:r><w:rPr>${base}${whole.test(s) ? '' : '<w:rtl/>'}</w:rPr><w:t xml:space="preserve">${esc(s)}</w:t></w:r>`).join('');
  }
  const cellPara = (t, bold) => `<w:p><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>${t ? runs(t, 22, bold) : ''}</w:p>`;
  function textPara(b) {
    const jc = b.align === 'center' ? 'center' : b.align === 'left' ? 'end' : 'start';
    const big = jc === 'center' && b.bold;
    return `<w:p><w:pPr><w:keepNext/><w:bidi/><w:spacing w:before="${big ? 240 : 60}" w:after="${big ? 160 : 60}"/><w:jc w:val="${jc}"/></w:pPr>${runs(b.text, big ? 30 : 26, !!b.bold, !!b.underline)}</w:p>`;
  }
  function widths(tb, TW) {
    const { grid, ncols } = tb;
    const sum = Array(ncols).fill(0), cnt = Array(ncols).fill(0), mx = Array(ncols).fill(0), hw = Array(ncols).fill(0);
    grid.forEach((row, r) => row.forEach((s, c) => {
      if (!s || !s.origin || s.ref.cs !== 1 || !s.ref.t) return;
      const lines = s.ref.t.split('\n');
      if (r < tb.headerRows) { hw[c] = Math.max(hw[c], ...lines.flatMap((l) => l.split(/\s+/)).map((w) => w.length)); return; }
      const L = Math.max(...lines.map((l) => l.length));
      sum[c] += L; cnt[c]++; mx[c] = Math.max(mx[c], L);
    }));
    const need = mx.map((m, c) => Math.max(hw[c] * 150, m <= 16 ? (m + 2) * 120 : 0, 520) + 260);
    const longCols = mx.map((m, c) => (m > 16 ? c : -1)).filter((c) => c >= 0);
    let w = need.slice();
    longCols.forEach((c) => { w[c] = Math.max(need[c], Math.min(40, sum[c] / Math.max(cnt[c], 1)) * 120 + 230); });
    const total = w.reduce((p, q) => p + q, 0);
    if (total > TW && longCols.length) {
      const fixed = w.reduce((p, q, c) => p + (longCols.includes(c) ? 0 : q), 0);
      const rest = Math.max(TW - fixed, longCols.length * 2600);
      const lsum = longCols.reduce((p, c) => p + w[c], 0);
      longCols.forEach((c) => { w[c] = (w[c] * rest) / lsum; });
    }
    // a table never wider than the page; narrow tables keep a natural width (centred), like the scan
    const t2 = w.reduce((p, q) => p + q, 0);
    if (t2 > TW) w = w.map((x) => (x * TW) / t2);
    else if (t2 < TW * 0.75) w = w.map((x) => (x * TW * 0.75) / t2);
    return w.map((x) => Math.floor(x));
  }
  function tableXml(tb, TW) {
    const { grid, ncols } = tb;
    const w = widths(tb, TW); const tw = w.reduce((p, q) => p + q, 0);
    const areaCols = new Set();
    grid.slice(0, tb.headerRows).forEach((row) => row.forEach((s, c) => { if (s && s.origin && /مساحة|المساحة/.test(s.ref.t)) for (let j = 0; j < s.ref.cs; j++) areaCols.add(c + j); }));
    const M = '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:start w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:end w:w="80" w:type="dxa"/></w:tcMar>';
    let x = `<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="${tw}" w:type="dxa"/><w:jc w:val="center"/>` +
      '<w:tblBorders>' + ['top', 'start', 'bottom', 'end', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>' +
      '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:start w:w="10" w:type="dxa"/><w:end w:w="10" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      `<w:tblGrid>${w.map((v) => `<w:gridCol w:w="${v}"/>`).join('')}</w:tblGrid>`;
    grid.forEach((row, r) => {
      const head = r < tb.headerRows, tot = !head && isTotalRow(row);
      const fill = head ? 'D9D9D9' : tot ? 'F2F2F2' : '';
      x += `<w:tr><w:trPr><w:cantSplit/><w:jc w:val="center"/></w:trPr>`;
      for (let c = 0; c < ncols; c++) {
        const s = row[c];
        if (s && s.hcont) continue;
        const cs = s ? s.ref.cs : 1;
        const width = w.slice(c, c + cs).reduce((p, q) => p + q, 0);
        let pr = `<w:tcW w:w="${width}" w:type="dxa"/>` + (cs > 1 ? `<w:gridSpan w:val="${cs}"/>` : '');
        let text = '';
        if (s && s.vcont) pr += '<w:vMerge/>';
        else if (s && s.origin) { if (s.ref.rs > 1) pr += '<w:vMerge w:val="restart"/>'; text = s.ref.t; }
        const q = text.includes('?');
        const shade = q ? 'FFF2CC' : fill;
        if (shade) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${shade}"/>`;
        pr += M + '<w:vAlign w:val="center"/>';
        let t = text;
        if (!head && !tot && areaCols.has(c) && num(t) !== null && !/م/.test(t)) t += ' م²';
        const bold = head || tot || (s && s.ref.b);
        x += `<w:tc><w:tcPr>${pr}</w:tcPr>` + (t ? t.split('\n').map((l) => cellPara(l, bold)).join('') : cellPara('', false)) + '</w:tc>';
        c += cs - 1;
      }
      x += '</w:tr>';
    });
    return x + '</w:tbl>';
  }

  // ---------- 5. assemble: pictures replaced by the page content, the rest of the file untouched ----------
  // results: {blockIndex: [block, ...]} where block = {type:'text',...} | {type:'table', rows, header_rows, continues_previous}
  async function build(JSZip, a, results) {
    // A4 landscape like the EDD_Ilot556 model, whatever page size the scan came with
    const sect = (a.doc.slice(a.doc.lastIndexOf('<w:sectPr')).match(/<w:sectPr[\s\S]*?<\/w:sectPr>/) || [''])[0];
    const TW = 16838 - 2 * 1417;
    let newSect = (sect || '<w:sectPr></w:sectPr>').replace(/<w:pgSz [^>]*\/>/, '').replace(/<w:pgMar [^>]*\/>/, '').replace(/<w:cols [^>]*\/>/, '');
    const pg = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>';
    newSect = newSect.replace(/(<w:sectPr[^>]*>)((?:<w:(?:headerReference|footerReference|endnotePr|footnotePr|type)\b[^>]*\/>)*)/, `$1$2${pg}`);

    const segs = [];
    const lastReal = () => { for (let i = segs.length - 1; i >= 0; i--) if (segs[i].t !== 'gap') return segs[i]; return null; };
    a.blocks.forEach((b, i) => {
      const res = results[i];
      if (!res) {
        const empty = b.t === 'p' && !stripTags(b.x).trim() && !/<w:drawing|<w:pict/.test(b.x);
        segs.push({ t: empty ? 'gap' : 'xml', x: b.x }); return;
      }
      let firstOnPage = true;
      for (const bl of res) {
        if (bl.type === 'text') { if (String(bl.text || '').trim()) segs.push({ t: 'text', b: bl }); firstOnPage = false; continue; }
        if (bl.type !== 'table' || !Array.isArray(bl.rows) || !bl.rows.length) continue;
        const g = toGrid(bl.rows);
        const hr = Math.max(0, Math.min(parseInt(bl.header_rows, 10) || 0, g.grid.length));
        const prev = lastReal();
        const cont = firstOnPage && hr === 0 && prev && prev.t === 'table' && prev.ncols === g.ncols && (bl.continues_previous || true);
        if (cont) {
          while (segs.length && segs[segs.length - 1].t === 'gap') segs.pop();
          prev.grid.push(...g.grid);
        } else segs.push({ t: 'table', grid: g.grid, ncols: g.ncols, headerRows: hr });
        firstOnPage = false;
      }
    });
    let body = ''; const report = []; let k = 0;
    const gap = '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';
    segs.forEach((s, i) => {
      if (s.t === 'xml' || s.t === 'gap') { body += s.x; return; }
      if (s.t === 'text') { body += textPara(s.b); return; }
      k++;
      body += tableXml(s, TW);
      const nxt = segs[i + 1];
      if (!nxt || nxt.t === 'table') body += gap;
      report.push({ n: k, rows: s.grid.length, issues: check(s) });
    });
    const doc = a.doc.slice(0, a.bStart) + body + newSect + '</w:body></w:document>';
    const z = new JSZip();
    const usedIds = new Set([...doc.matchAll(/r:(?:embed|id|link)="([^"]+)"/g)].map((m) => m[1]));
    const replaced = new Set(a.images.filter((im) => results[im.block]).map((im) => im.rid));
    for (const [name, file] of Object.entries(a.zip.files)) {
      if (file.dir) continue;
      if (name === 'word/document.xml') { z.file(name, doc); continue; }
      if (name === 'word/_rels/document.xml.rels') {
        let r = await file.async('string');
        r = r.replace(/<Relationship [^>]*\/>/g, (m) => { const id = (m.match(/Id="([^"]+)"/) || [])[1]; return replaced.has(id) && !usedIds.has(id) ? '' : m; });
        z.file(name, r); continue;
      }
      if (a.images.some((im) => im.path === name && replaced.has(im.rid)) && !a.images.some((im) => im.path === name && usedIds.has(im.rid))) continue;
      z.file(name, await file.async('uint8array'));
    }
    return { bytes: await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), report };
  }


  // a PDF has no Word file around it: start from an empty landscape Word file (shell) and put each page in order
  async function buildFromPages(JSZip, shellFiles, results, nPages) {
    const zip = new JSZip();
    for (const [n, c] of Object.entries(shellFiles)) zip.file(n, c);
    const shellDoc = shellFiles['word/document.xml'];
    const bStart = shellDoc.indexOf('{{BODY}}');
    const doc = shellDoc.replace('{{BODY}}', '');
    const blocks = []; const res = {};
    for (let p = 0; p < nPages; p++) {
      if (results[p]) { blocks.push({ t: 'p', x: '' }); res[blocks.length - 1] = results[p]; }
      else blocks.push({ t: 'p', x: `<w:p><w:pPr><w:bidi/><w:jc w:val="center"/></w:pPr>${runs('⚠ لم تُقرأ الصفحة ' + (p + 1) + ' من ملف PDF — أعيدي قراءتها أو اكتبيها يدوياً.', 24, true)}</w:p>` });
    }
    // reuse the Word-file builder: the shell plays the role of the original file
    await zip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
    return build(JSZip, { zip, doc, bStart, blocks, images: [] }, res);
  }

  root.EDD556 = { analyse, prompt, slices, build, buildFromPages, check, toGrid };
})(typeof window !== 'undefined' ? window : globalThis);
