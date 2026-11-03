/* Old Word files (.doc, Word 97-2003) -> a simple .docx, in the browser.
   Reads the compound file, the piece table, paragraph and character properties (FKPs), the style sheet
   and the table rows (sprmTDefTable), then writes paragraphs, runs and tables as WordprocessingML.
   Kept: text, paragraphs, tables (cell widths, merged cells), bold, italic, underline, size, alignment,
   right-to-left. Not kept: pictures, text boxes, headers/footers, footnotes, exact fonts. */
(function (root) {
  'use strict';

  // ---------- compound file (CFB) ----------
  function cfbRead(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    if (!sig.every((b, i) => u8[i] === b)) throw new Error('NOT_DOC');
    const ss = 1 << dv.getUint16(0x1E, true), mss = 1 << dv.getUint16(0x20, true);
    const nFat = dv.getUint32(0x2C, true), dirStart = dv.getUint32(0x30, true);
    const cutoff = dv.getUint32(0x38, true), miniFatStart = dv.getUint32(0x3C, true), nMiniFat = dv.getUint32(0x40, true);
    let difStart = dv.getUint32(0x44, true); const nDif = dv.getUint32(0x48, true);
    const sectOff = (s) => (s + 1) * ss;
    const fatSects = [];
    for (let i = 0; i < 109 && fatSects.length < nFat; i++) fatSects.push(dv.getUint32(0x4C + i * 4, true));
    for (let k = 0; k < nDif && fatSects.length < nFat; k++) {
      const o = sectOff(difStart);
      for (let i = 0; i < ss / 4 - 1 && fatSects.length < nFat; i++) fatSects.push(dv.getUint32(o + i * 4, true));
      difStart = dv.getUint32(o + ss - 4, true);
    }
    const fat = new Uint32Array(fatSects.length * ss / 4);
    fatSects.forEach((s, k) => { const o = sectOff(s); for (let i = 0; i < ss / 4; i++) fat[k * ss / 4 + i] = dv.getUint32(o + i * 4, true); });
    const chain = (start, table) => { const out = []; let s = start, guard = 0; while (s < 0xFFFFFFFA && guard++ < 10000000) { out.push(s); s = table[s]; } return out; };
    const readBig = (start, size) => {
      const out = new Uint8Array(size); let p = 0;
      for (const s of chain(start, fat)) { const n = Math.min(ss, size - p); if (n <= 0) break; out.set(u8.subarray(sectOff(s), sectOff(s) + n), p); p += n; }
      return out;
    };
    // directory
    const dirSects = chain(dirStart, fat);
    const entries = [];
    for (const s of dirSects) for (let i = 0; i < ss / 128; i++) {
      const o = sectOff(s) + i * 128;
      const nl = dv.getUint16(o + 0x40, true);
      let name = ''; for (let j = 0; j < Math.max(0, nl / 2 - 1); j++) name += String.fromCharCode(dv.getUint16(o + j * 2, true));
      entries.push({ name, type: u8[o + 0x42], start: dv.getUint32(o + 0x74, true), size: dv.getUint32(o + 0x78, true) });
    }
    const rootE = entries[0];
    const ministream = rootE && rootE.size ? readBig(rootE.start, rootE.size) : new Uint8Array(0);
    let minifat = new Uint32Array(0);
    if (nMiniFat) { const mf = readBig(miniFatStart, nMiniFat * ss); minifat = new Uint32Array(mf.buffer, mf.byteOffset, mf.byteLength / 4); }
    const readMini = (start, size) => {
      const out = new Uint8Array(size); let p = 0;
      for (const s of chain(start, minifat)) { const n = Math.min(mss, size - p); if (n <= 0) break; out.set(ministream.subarray(s * mss, s * mss + n), p); p += n; }
      return out;
    };
    return {
      stream(name) {
        const e = entries.find((x) => x.type === 2 && x.name === name);
        if (!e) return null;
        return e.size < cutoff ? readMini(e.start, e.size) : readBig(e.start, e.size);
      },
    };
  }

  // ---------- sprms ----------
  function sprmSize(sprm, u8, p) {
    switch ((sprm >> 13) & 7) {
      case 0: case 1: return 1;
      case 2: case 4: case 5: return 2;
      case 3: return 4;
      case 7: return 3;
      case 6:
        if (sprm === 0xD608 || sprm === 0xD606) return (u8[p] | (u8[p + 1] << 8)) + 1;   // sprmTDefTable: 2-byte length, value is cb-1 bytes
        if (sprm === 0xC615) { const cb = u8[p]; return cb === 255 ? 1 + u8[p + 1] * 4 + 1 + u8[p + 2 + u8[p + 1] * 4] * 3 + 1 : cb + 1; }
        return u8[p] + 1;
    }
    return 0;
  }
  function parseGrpprl(u8, start, end) {
    const out = [];
    let p = start;
    while (p + 2 <= end) {
      const sprm = u8[p] | (u8[p + 1] << 8); p += 2;
      const n = sprmSize(sprm, u8, p);
      if (n <= 0 || p + n > end + 1) break;
      out.push({ sprm, p, n });
      p += n;
    }
    return out;
  }

  // ---------- main ----------
  function readDoc(buf) {
    const cfb = cfbRead(buf);
    const wd = cfb.stream('WordDocument'); if (!wd) throw new Error('NOT_DOC');
    const W = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    if (W.getUint16(0, true) !== 0xA5EC) throw new Error('NOT_DOC');
    const flags = W.getUint16(0x0A, true);
    if (flags & 0x0100) throw new Error('ENCRYPTED');
    const tbl = cfb.stream(flags & 0x0200 ? '1Table' : '0Table'); if (!tbl) throw new Error('NOT_DOC');
    const T = new DataView(tbl.buffer, tbl.byteOffset, tbl.byteLength);
    const csw = W.getUint16(32, true);
    const lwOff = 34 + csw * 2;
    const cslw = W.getUint16(lwOff, true);
    const lw = (k) => W.getInt32(lwOff + 2 + k * 4, true);
    const ccpText = lw(3), ccpFtn = lw(4), ccpHdd = lw(5), ccpAtn = lw(7), ccpEdn = lw(8), ccpTxbx = lw(9);
    const fcOff = lwOff + 2 + cslw * 4 + 2;
    const fcLcb = (i) => [W.getUint32(fcOff + i * 8, true), W.getUint32(fcOff + i * 8 + 4, true)];

    // ----- piece table (Clx) -----
    const [fcClx, lcbClx] = fcLcb(33);
    let p = fcClx;
    while (p < fcClx + lcbClx && tbl[p] === 0x01) p += 3 + T.getInt16(p + 1, true);
    if (tbl[p] !== 0x02) throw new Error('NO_PIECES');
    const lcbPlc = T.getUint32(p + 1, true); p += 5;
    const nP = (lcbPlc - 4) / 12;
    const pieces = [];
    for (let i = 0; i < nP; i++) {
      const cpS = T.getUint32(p + i * 4, true), cpE = T.getUint32(p + (i + 1) * 4, true);
      const raw = T.getUint32(p + (nP + 1) * 4 + i * 8 + 2, true);
      const comp = !!(raw & 0x40000000);
      const fc = comp ? (raw & 0x3FFFFFFF) / 2 : (raw & 0x3FFFFFFF);
      pieces.push({ cpS, cpE, fc, comp });
    }
    const CP1252 = { 0x80: 0x20AC, 0x82: 0x201A, 0x83: 0x0192, 0x84: 0x201E, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02C6, 0x89: 0x2030, 0x8A: 0x0160, 0x8B: 0x2039, 0x8C: 0x0152, 0x8E: 0x017D, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201C, 0x94: 0x201D, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02DC, 0x99: 0x2122, 0x9A: 0x0161, 0x9B: 0x203A, 0x9C: 0x0153, 0x9E: 0x017E, 0x9F: 0x0178 };
    // characters of the main text with their file position (fc)
    const chars = []; const fcs = [];
    for (const pc of pieces) {
      for (let cp = pc.cpS; cp < pc.cpE; cp++) {
        if (pc.comp) { const fc = pc.fc + (cp - pc.cpS); const b = wd[fc]; chars.push(CP1252[b] || b); fcs.push(fc); }
        else { const fc = pc.fc + (cp - pc.cpS) * 2; chars.push(W.getUint16(fc, true)); fcs.push(fc); }
      }
    }

    // ----- FKP lookups (paragraph and character properties by fc) -----
    function plcBte(i) {
      const [fc, lcb] = fcLcb(i); if (!lcb) return [];
      const n = (lcb - 4) / 8; const out = [];
      for (let k = 0; k < n; k++) out.push({ fcS: T.getUint32(fc + k * 4, true), fcE: T.getUint32(fc + (k + 1) * 4, true), pn: T.getUint32(fc + (n + 1) * 4 + k * 4, true) & 0x3FFFFF });
      return out;
    }
    const papRuns = []; // {fcS, fcE, istd, sprms}
    for (const b of plcBte(13)) {
      const base = b.pn * 512; const crun = wd[base + 511];
      for (let k = 0; k < crun; k++) {
        const fcS = W.getUint32(base + k * 4, true), fcE = W.getUint32(base + (k + 1) * 4, true);
        const bOff = wd[base + (crun + 1) * 4 + k * 13] * 2;
        if (!bOff) { papRuns.push({ fcS, fcE, istd: 0, g: [] }); continue; }
        let q = base + bOff; let cb = wd[q++]; let len;
        if (cb === 0) { cb = wd[q++]; len = cb * 2; } else len = cb * 2 - 1;
        const istd = W.getUint16(q, true);
        papRuns.push({ fcS, fcE, istd, g: parseGrpprl(wd, q + 2, q + len) });
      }
    }
    const chpRuns = [];
    for (const b of plcBte(12)) {
      const base = b.pn * 512; const crun = wd[base + 511];
      for (let k = 0; k < crun; k++) {
        const fcS = W.getUint32(base + k * 4, true), fcE = W.getUint32(base + (k + 1) * 4, true);
        const off = wd[base + (crun + 1) * 4 + k] * 2;
        if (!off) { chpRuns.push({ fcS, fcE, g: [] }); continue; }
        const cb = wd[base + off];
        chpRuns.push({ fcS, fcE, g: parseGrpprl(wd, base + off + 1, base + off + 1 + cb) });
      }
    }
    const sortF = (a, b) => a.fcS - b.fcS; papRuns.sort(sortF); chpRuns.sort(sortF);
    const find = (runs, fc) => { let lo = 0, hi = runs.length - 1; while (lo <= hi) { const m = (lo + hi) >> 1; if (fc < runs[m].fcS) hi = m - 1; else if (fc >= runs[m].fcE) lo = m + 1; else return runs[m]; } return null; };

    // ----- styles (STSH): paragraph alignment / bidi and character props by style -----
    const styles = [];
    try {
      const [fcSt, lcbSt] = fcLcb(1);
      if (lcbSt) {
        const cbStshi = T.getUint16(fcSt, true);
        const cstd = T.getUint16(fcSt + 2, true), cbBase = T.getUint16(fcSt + 4, true);
        let q = fcSt + 2 + cbStshi;
        for (let i = 0; i < cstd; i++) {
          const cbStd = T.getUint16(q, true); const s0 = q + 2; q = s0 + cbStd;
          if (!cbStd) { styles.push(null); continue; }
          const stk = T.getUint16(s0 + 2, true) & 0xF, istdBase = T.getUint16(s0 + 2, true) >> 4, cupx = T.getUint16(s0 + 4, true) & 0xF;
          let r = s0 + cbBase;
          const cch = T.getUint16(r, true); r += 2 + cch * 2 + 2;
          const st = { stk, base: istdBase, p: [], c: [] };
          for (let u = 0; u < cupx && r < q; u++) {
            const cbUpx = T.getUint16(r, true); r += 2;
            if (stk === 1 && u === 0) st.p = parseGrpprl(tbl, r + 2, r + cbUpx);
            else if ((stk === 1 && u === 1) || (stk === 2 && u === 0)) st.c = parseGrpprl(tbl, r, r + cbUpx);
            r += cbUpx + (cbUpx & 1);
          }
          styles.push(st);
        }
      }
    } catch (e) { /* no style sheet: defaults */ }
    const styleChain = (istd) => { const out = []; let g = 0; while (istd !== 0x0FFF && styles[istd] && g++ < 12) { out.unshift(styles[istd]); istd = styles[istd].base; } return out; };

    const u16 = (u8, at) => u8[at] | (u8[at + 1] << 8);
    function papProps(run) {
      const pr = { inTable: false, ttp: false, jc: null, bidi: false, itap: 0, tdef: null, tbidi: false, cs: {} };
      const apply = (g, u8) => {
        for (const s of g) {
          const v = u8[s.p];
          switch (s.sprm) {
            case 0x2416: pr.inTable = !!v; break;
            case 0x2417: pr.ttp = !!v; break;
            case 0x6649: pr.itap = u8[s.p] | (u8[s.p + 1] << 8); if (pr.itap) pr.inTable = true; break;
            case 0x244B: pr.ttp = !!v; break;                       // sprmPFInnerTtp
            case 0x2403: if (pr.jc === null) pr.jc = v; break;
            case 0x2461: pr.jc = v; break;
            case 0x2441: pr.bidi = !!v; break;
            case 0x560B: pr.tbidi = !!(v | u8[s.p + 1]); break;
            case 0xD608: pr.tdef = { u8, at: s.p + 2, len: s.n - 2 }; break;
          }
        }
      };
      if (run) { for (const st of styleChain(run.istd)) apply(st.p, tbl); apply(run.g, wd); }
      return pr;
    }
    function chpProps(run, istd) {
      const c = { b: false, bBi: false, i: false, iBi: false, u: false, sz: null, rtl: false };
      const apply = (g, u8) => {
        for (const s of g) {
          const v = u8[s.p];
          const tog = (cur) => (v === 0 ? false : v === 1 ? true : v === 0x81 ? !cur : cur);
          switch (s.sprm) {
            case 0x0835: c.b = tog(c.b); break;
            case 0x085C: c.bBi = tog(c.bBi); break;
            case 0x0836: c.i = tog(c.i); break;
            case 0x085D: c.iBi = tog(c.iBi); break;
            case 0x2A3E: c.u = v !== 0; break;
            case 0x4A43: c.sz = u16(u8, s.p); break;
            case 0x4A61: c.szBi = u16(u8, s.p); break;
            case 0x085A: c.rtl = tog(c.rtl); break;
          }
        }
      };
      for (const st of styleChain(istd)) apply(st.c, tbl);
      if (run) apply(run.g, wd);
      return c;
    }

    // ----- walk the text: paragraphs, cells, rows -----
    const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const AR = /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/;
    const runXml = (r) => {
      const c = r.key;
      const sz = c.szBi || c.sz || 24;
      const rtl = c.rtl || AR.test(r.text);
      const parts = r.text.split('\n');
      const rPr = `<w:rPr>${(c.b || c.bBi) ? '<w:b/><w:bCs/>' : ''}${(c.i || c.iBi) ? '<w:i/><w:iCs/>' : ''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${c.u ? '<w:u w:val="single"/>' : ''}${rtl ? '<w:rtl/>' : ''}</w:rPr>`;
      return parts.map((t, k) => (k ? `<w:r>${rPr}<w:br/></w:r>` : '') + (t ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : '')).join('');
    };
    // text boxes: anchor cp (0x08 in the text) -> shape id -> text-box story
    const spaAt = {};
    try {
      const [fcSpa, lcbSpa] = fcLcb(40);
      if (lcbSpa) { const n = (lcbSpa - 4) / 30; for (let k = 0; k < n; k++) spaAt[T.getUint32(fcSpa + k * 4, true)] = T.getInt32(fcSpa + (n + 1) * 4 + k * 26, true); }
    } catch (e) {}
    const storyBySpid = {};
    try {
      const [fcTx, lcbTx] = fcLcb(56);
      if (lcbTx && ccpTxbx) {
        const n = (lcbTx - 4) / 26; const base = ccpText + ccpFtn + ccpHdd + ccpAtn + ccpEdn;
        for (let k = 0; k < n - 1; k++) {
          const s0 = T.getUint32(fcTx + k * 4, true), s1 = T.getUint32(fcTx + (k + 1) * 4, true);
          const f = fcTx + (n + 1) * 4 + k * 22;
          const reusable = T.getUint16(f + 8, true), lid = T.getInt32(f + 14, true);
          if (!reusable) storyBySpid[lid] = [base + s0, base + s1];
        }
      }
    } catch (e) {}
    let anchors = [];
    const anchorXml = () => {
      const out = [];
      for (const cp of anchors.splice(0)) {
        const story = storyBySpid[spaAt[cp]];
        if (!story) continue;
        const sub = walk(story[0], story[1], false);
        for (const b of sub) if (b.t === 'p' && b.xml.replace(/<[^>]+>/g, '').trim()) out.push(b.xml);
      }
      return out;
    };
    function walk(cpA, cpB, onAnchor) {
      const blocks = [];
      let curRuns = [], curRunKey = null, curText = '';
      let fieldDepth = 0, fieldHide = [];
      let cellParas = [], rowCells = [];
      const flushRun = () => { if (curText) curRuns.push({ key: curRunKey, text: curText }); curText = ''; };
      const paraXml = (pp) => {
        flushRun();
        const text = curRuns.map((r) => r.text).join('');
        const bidi = pp.bidi || AR.test(text);
        const jcMap = { 0: 'start', 1: 'center', 2: 'end', 3: 'both', 4: 'both', 5: 'both' };
        const jc = pp.jc === null ? null : jcMap[pp.jc] || null;
        const xml = `<w:p><w:pPr>${bidi ? '<w:bidi/>' : ''}${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr>${curRuns.map(runXml).join('')}</w:p>`;
        curRuns = []; curRunKey = null;
        return xml;
      };
    for (let i = cpA; i < cpB && i < chars.length; i++) {
        let ch = chars[i]; const fc = fcs[i];
        // fields: show only the result part
        if (ch === 0x13) { flushRun(); fieldDepth++; fieldHide.push(true); continue; }
        if (ch === 0x14) { flushRun(); if (fieldHide.length) fieldHide[fieldHide.length - 1] = false; continue; }
        if (ch === 0x15) { flushRun(); fieldDepth = Math.max(0, fieldDepth - 1); fieldHide.pop(); continue; }
        if (fieldHide.length && fieldHide[fieldHide.length - 1] && ch !== 0x0D && ch !== 0x07) continue;
        if (ch === 0x0D || ch === 0x07) {
          const prun = find(papRuns, fc); const pp = papProps(prun);
          const istd = prun ? prun.istd : 0;
          const xml = paraXml(pp);
          if (ch === 0x0D) {
            if (pp.inTable) cellParas.push(xml, ...anchorXml()); else { blocks.push({ t: 'p', xml }); for (const x of anchorXml()) blocks.push({ t: 'p', x2: true, xml: x }); }
          } else if (pp.ttp) {               // end of a table row
            blocks.push({ t: 'row', cells: rowCells, tdef: pp.tdef, bidi: pp.tbidi });
            rowCells = []; cellParas = [];
          } else {                            // end of a cell
            cellParas.push(xml, ...anchorXml()); rowCells.push(cellParas); cellParas = [];
          }
          void istd;
          continue;
        }
        if (ch === 0x08) { if (onAnchor) anchors.push(i); continue; }
        if (ch === 0x0C || ch === 0x01 || ch === 0x02 || ch === 0x05 || ch === 0x03 || ch === 0x04) continue;   // page break, pictures, notes marks
        if (ch === 0x0B) ch = 0x0A;            // line break inside the paragraph
        if (ch === 0x1E) ch = 0x2D;            // non-breaking hyphen
        if (ch === 0x1F) continue;             // optional hyphen
        if (ch === 0x09) ch = 0x09;
        const prun = find(papRuns, fc);
        const key = chpProps(find(chpRuns, fc), prun ? prun.istd : 0);
        const k = JSON.stringify(key);
        if (curRunKey && JSON.stringify(curRunKey) !== k) flushRun();
        curRunKey = key;
        curText += ch === 0x09 ? '\t' : String.fromCharCode(ch);
      }
      if (curRuns.length || curText) blocks.push({ t: 'p', xml: paraXml({ jc: null, bidi: false }) });
      return blocks;
    }
    const blocks = walk(0, ccpText, true);

    // ----- tables: consecutive rows -> one table, grid from all cell boundaries -----
    const out = [];
    let tableRows = [];
    const flushTable = () => {
      if (!tableRows.length) return;
      const rowsInfo = tableRows.map((r) => {
        let bounds = null, merges = [];
        if (r.tdef) {
          const u8 = r.tdef.u8, at = r.tdef.at; const itc = u8[at];
          bounds = []; for (let k = 0; k <= itc; k++) { let v = u8[at + 1 + k * 2] | (u8[at + 2 + k * 2] << 8); if (v > 32767) v -= 65536; bounds.push(v); }
          const tcAt = at + 1 + (itc + 1) * 2;
          for (let k = 0; k < itc; k++) {
            const f = (tcAt + k * 20 + 1 < at + r.tdef.len) ? u8[tcAt + k * 20] | (u8[tcAt + k * 20 + 1] << 8) : 0;
            merges.push({ hFirst: !!(f & 1), hMerged: !!(f & 2), vMerge: !!(f & 0x20), vRestart: !!(f & 0x40) });
          }
        }
        return { cells: r.cells, bounds, merges, bidi: r.bidi };
      });
      const all = new Set();
      rowsInfo.forEach((ri) => { if (ri.bounds && ri.bounds.length === ri.cells.length + 1) ri.bounds.forEach((b) => all.add(b)); });
      let grid = [...all].sort((a, b) => a - b);
      const maxCells = Math.max(...rowsInfo.map((ri) => ri.cells.length));
      if (grid.length < 2) { grid = []; for (let k = 0; k <= maxCells; k++) grid.push(k * Math.round(9000 / maxCells)); }
      const gw = grid.slice(1).map((g, k) => Math.max(1, g - grid[k]));
      const bidi = rowsInfo.some((ri) => ri.bidi) || rowsInfo.some((ri) => ri.cells.some((c) => c.join('').match(AR)));
      let x = `<w:tbl><w:tblPr>${bidi ? '<w:bidiVisual/>' : ''}<w:tblW w:w="${gw.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:tblBorders>` +
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
        `</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${gw.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
      for (const ri of rowsInfo) {
        x += '<w:tr>';
        const okB = ri.bounds && ri.bounds.length === ri.cells.length + 1;
        let skip = 0;
        ri.cells.forEach((paras, k) => {
          const m = ri.merges[k] || {};
          if (m.hMerged && !m.hFirst) return;        // old-style horizontal merge: covered cell
          let span = 1, a = 0, b = 1;
          if (okB) {
            let endK = k + 1;
            if (m.hFirst) while (endK < ri.cells.length && (ri.merges[endK] || {}).hMerged && !(ri.merges[endK] || {}).hFirst) endK++;
            a = grid.indexOf(ri.bounds[k]); b = grid.indexOf(ri.bounds[endK]); span = Math.max(1, b - a);
          }
          const w = okB ? gw.slice(a, a + span).reduce((p, q) => p + q, 0) : gw[Math.min(k, gw.length - 1)];
          const vm = m.vMerge ? (m.vRestart ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') : '';
          x += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${vm}<w:vAlign w:val="center"/></w:tcPr>${paras.length ? paras.join('') : '<w:p/>'}</w:tc>`;
          void skip;
        });
        x += '</w:tr>';
      }
      x += '</w:tbl>';
      out.push(x);
      tableRows = [];
    };
    for (const b of blocks) {
      if (b.t === 'row') { tableRows.push(b); continue; }
      flushTable(); out.push(b.xml);
    }
    flushTable();

    // base size of the Normal style (for runs that have none)
    let baseSz = 24;
    if (styles[0]) for (const s of styles[0].c) if (s.sprm === 0x4A43 || s.sprm === 0x4A61) baseSz = u16(tbl, s.p);
    return { body: out.join(''), baseSz };
  }

  async function docToDocx(JSZip, buf) {
    const { body } = readDoc(buf);
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/><w:bidi/></w:sectPr></w:body></w:document>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
    const z = new JSZip();
    z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    z.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    z.file('word/document.xml', doc);
    z.file('word/styles.xml', styles);
    return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  root.DOC2DOCX = { readDoc, docToDocx };
})(typeof window !== 'undefined' ? window : globalThis);
