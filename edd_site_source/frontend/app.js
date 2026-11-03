(function () {
  const $ = (id) => document.getElementById(id);
  const SHELLS = JSON.parse(document.getElementById('shells').textContent);
  let analysed = null, result = null, baseName = 'EDD', a556 = null, pdf556 = null, name556 = 'EDD', out556 = null, ctl556 = null;
  const showErr = (el, t) => { el.textContent = t; el.hidden = !t; };

  async function load(file) {
    showErr($('err'), ''); $('done').hidden = true; $('dl').hidden = true; result = null;
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      showErr($('err'), /\.doc$/i.test(file.name)
        ? 'هذا ملف .doc قديم. افتحيه في Word ثم «حفظ باسم» واختاري .docx، وضعيه هنا من جديد.'
        : 'الملف يجب أن يكون Word بصيغة .docx.');
      return;
    }
    const buf = await file.arrayBuffer();
    try {
      analysed = await EDD.analyse(JSZip, buf);
    } catch (e) {
      analysed = null; $('go').disabled = true;
      showErr($('err'), 'تعذّر فتح الملف. تأكدي أنه ملف Word سليم بصيغة .docx.');
      return;
    }
    baseName = file.name.replace(/\.docx$/i, '');
    $('fname').textContent = file.name;
    const nP = analysed.blocks.filter((b) => b.t === 'p' && b.x.replace(/<[^>]+>/g, '').trim()).length;
    const nT = analysed.blocks.filter((b) => b.t === 'tbl').length;
    $('stats').innerHTML = '';
    [[nP + ' فقرة'], [nT + ' جدول'], analysed.boxed ? ['الملف داخل الإطار أصلاً'] : null].filter(Boolean).forEach(([t]) => {
      const s = document.createElement('span'); s.className = 'chip'; s.textContent = t; $('stats').appendChild(s);
    });
    $('loaded').hidden = false;
    const b = analysed.box;
    $('fQism').value = b.qism || ''; $('fBal').value = b.baladiya || ''; $('fMaj').value = b.majmoua || '';
    $('fHissa').value = ''; $('fWil').value = '';
    $('detHint').textContent = 'مُلئت الخانات من نص الملف. تأكدي منها وصحّحيها إن لزم، خاصة البلدية.';
    $('go').disabled = false;
  }

  $('inp').addEventListener('change', (e) => load(e.target.files[0]));
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => load(e.dataTransfer.files[0]));

  $('go').addEventListener('click', async () => {
    if (!analysed) return;
    $('go').disabled = true; $('go').textContent = 'جارٍ الإنشاء…';
    try {
      result = await EDD.convert(JSZip, analysed, { tpl: SHELLS.tpl }, {
        qism: $('fQism').value, baladiya: $('fBal').value, majmoua: $('fMaj').value,
        hissa: $('fHissa').value, wilaya: $('fWil').value,
        idaHajm: $('fIdaHajm').value, idaRaqm: $('fIdaRaqm').value, fi: $('fFi').value,
        mujallad: $('fMujallad').value, mujRaqm: $('fMujRaqm').value, rasm: $('fRasm').value,
      });
      const s = result.stats;
      $('done').textContent = `تمّ: ${s.paras} فقرة، ${s.tables} جدول (${s.rows} سطر)، و«-----» في آخر ${s.dashes} فقرة. حمّلي الملف وافتحيه في Word للتأكد.`;
      $('done').hidden = false; $('dl').hidden = false;
    } catch (e) {
      showErr($('dlErr'), 'حدث خطأ أثناء التحويل. أرسلي هذا الملف إلى Claude ليصلح المحوّل.');
    }
    $('go').disabled = false; $('go').textContent = 'إنشاء الملف من جديد';
  });

  async function save(bytes, name) {
    showErr($('dlErr'), '');
    const dl = window.claude && await claude.use('downloads');
    if (!dl) { showErr($('dlErr'), 'التحميل غير متاح في هذه النافذة. افتحي الصفحة في claude.ai.'); return; }
    try { await dl.save({ filename: name, data: new Blob([bytes]) }); }
    catch (e) {
      if (e && e.code === 'declined') return;
      showErr($('dlErr'), e && e.code === 'rate_limited' ? 'نافذة تحميل مفتوحة بالفعل. أكمليها ثم أعيدي المحاولة.' : 'لم يتم التحميل. أعيدي المحاولة.');
    }
  }
  $('dlTpl').addEventListener('click', () => result && save(result.template, baseName + ' - إجراء إشهار عقاري.docx'));

  // ---------- 556: scanned tables ----------
  const chip = (t, warn) => { const s = document.createElement('span'); s.className = 'chip'; if (warn) s.style.cssText = 'background:var(--warn-soft);color:var(--warn)'; s.textContent = t; return s; };
  // pdf.js runs its worker on the page itself (the frame allows scripts from the CDN only)
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  async function load556(file) {
    a556 = null; pdf556 = null; out556 = null; $('dl556wrap').hidden = true; $('rep556').hidden = true; $('prog556').hidden = true; showErr($('err556'), '');
    $('imgStats').innerHTML = ''; $('go556').disabled = true;
    if (!file) return;
    name556 = file.name.replace(/\.(pdf|docx)$/i, '');
    $('fname556').textContent = file.name; $('fname556').hidden = false;
    const buf = await file.arrayBuffer();
    if (/\.pdf$/i.test(file.name)) {
      if (!window.pdfjsLib) { showErr($('err556'), 'تعذّر تحميل قارئ PDF. أعيدي فتح الصفحة.'); return; }
      try { pdf556 = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise; }
      catch (e) { $('imgStats').appendChild(chip('تعذّر فتح ملف PDF', true)); return; }
      $('imgStats').appendChild(chip(pdf556.numPages + ' صفحة في الملف'));
      $('go556').disabled = false; return;
    }
    if (!/\.docx$/i.test(file.name)) { showErr($('err556'), 'الملف يجب أن يكون PDF (أو Word بصيغة .docx).'); return; }
    try { a556 = await EDD556.analyse(JSZip, buf); } catch (e) { $('imgStats').appendChild(chip('تعذّر فتح الملف', true)); return; }
    const ok = a556.images.filter((i) => i.ok).length, bad = a556.images.length - ok;
    $('imgStats').appendChild(chip(ok + ' صورة جدول في الملف'));
    if (bad) $('imgStats').appendChild(chip(bad + ' صورة بصيغة لا تُقرأ (تبقى كما هي)', true));
    $('go556').disabled = ok === 0;
    if (!ok) showErr($('err556'), 'لا توجد صور جداول داخل هذا الملف.');
  }
  $('inp556').addEventListener('change', (e) => load556(e.target.files[0]));
  const drop556 = $('drop556');
  ['dragenter', 'dragover'].forEach((ev) => drop556.addEventListener(ev, (e) => { e.preventDefault(); drop556.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop556.addEventListener(ev, (e) => { e.preventDefault(); drop556.classList.remove('over'); }));
  drop556.addEventListener('drop', (e) => load556(e.dataTransfer.files[0]));
  // one PDF page as a sharp picture (about 300 dpi for an A4 page) + its text layer, if any
  async function pdfPage(n) {
    const page = await pdf556.getPage(n);
    const v1 = page.getViewport({ scale: 1 });
    const scale = Math.min(2600 / v1.width, 4200 / v1.height, 4.5);
    const v = page.getViewport({ scale });
    const c = document.createElement('canvas'); c.width = Math.floor(v.width); c.height = Math.floor(v.height);
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    await page.render({ canvasContext: g, viewport: v }).promise;
    const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.95));
    let text = '';
    try { const tc = await page.getTextContent(); text = tc.items.map((i) => i.str).join(' ').replace(/\s+/g, ' ').trim(); } catch (e) {}
    page.cleanup();
    return { blob, text: text.length > 30 ? text.slice(0, 6000) : '' };
  }
  $('stop556').addEventListener('click', () => ctl556 && ctl556.abort());
  $('go556').addEventListener('click', async () => {
    if (!a556 && !pdf556) return;
    showErr($('err556'), '');
    const sample = window.claude ? await claude.use('sample') : null;
    if (!sample) { showErr($('err556'), 'قراءة الصور تعمل فقط عند فتح الصفحة داخل Claude.'); return; }
    const lim = await sample.limits().catch(() => null);
    if (!lim || !lim.images) { showErr($('err556'), 'إرسال الصور إلى Claude غير متاح في هذه النافذة.'); return; }
    const isPdf = !!pdf556;
    const imgs = isPdf ? Array.from({ length: pdf556.numPages }, (_, p) => ({ page: p })) : a556.images.filter((i) => i.ok);
    const unit = isPdf ? 'الصفحة' : 'الصورة';
    const results = {}; let lastTable = null, failed = 0, stopped = false, readOk = 0;
    $('go556').disabled = true; $('stop556').hidden = false; $('dl556wrap').hidden = true; $('rep556').hidden = true;
    $('prog556').hidden = false;
    for (let k = 0; k < imgs.length; k++) {
      const im = imgs[k];
      $('prog556').textContent = `جارٍ قراءة ${unit} ${k + 1} من ${imgs.length}… (قد تأخذ كل ${unit === 'الصفحة' ? 'صفحة' : 'صورة'} بين دقيقة ودقيقتين لقراءتها بدقة)`;
      ctl556 = new AbortController();
      try {
        let raw, pdfText = '';
        if (isPdf) { const pg = await pdfPage(im.page + 1); raw = pg.blob; pdfText = pg.text; }
        else {
          const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[im.ext];
          raw = new Blob([await a556.zip.file(im.path).async('uint8array')], { type: mime });
        }
        const parts = await EDD556.slices(raw, lim.images.maxCount);
        const ans = await sample.json(EDD556.prompt(lastTable, parts.length, pdfText), { images: parts, signal: ctl556.signal, modelTier: 'complex' });
        const blocks = ans && Array.isArray(ans.blocks) && !ans.unreadable ? ans.blocks.filter((b) => b && (b.type === 'text' || (b.type === 'table' && Array.isArray(b.rows)))) : [];
        if (blocks.some((b) => b.type === 'table')) {
          const key = isPdf ? im.page : im.block; (results[key] = results[key] || []).push(...blocks); readOk++;
          for (const b of blocks) if (b.type === 'table') {
            const g = EDD556.toGrid(b.rows);
            const hr = parseInt(b.header_rows, 10) || 0;
            lastTable = { cols: g.ncols, head: hr ? b.rows.slice(0, hr).map((r) => r.map((c) => c && c.t)) : (lastTable ? lastTable.head : []) };
          }
        } else failed++;
      } catch (e) {
        if (e && e.code === 'cancelled') { stopped = true; break; }
        if (e && (e.code === 'not_granted' || e.code === 'sampling_disabled' || e.code === 'images_unavailable')) { showErr($('err556'), 'لم يُسمح لهذه الصفحة باستعمال Claude، فلا يمكن قراءة الصور.'); stopped = true; break; }
        if (e && e.code === 'rate_limited') { showErr($('err556'), 'تم الوصول إلى حد الاستعمال الآن. انتظري قليلاً ثم أعيدي المحاولة.'); stopped = true; break; }
        failed++;
      }
    }
    $('stop556').hidden = true; $('go556').disabled = false; ctl556 = null;
    if (!Object.keys(results).length) { $('prog556').textContent = stopped ? 'تم الإيقاف.' : 'لم تُقرأ أي صورة.'; return; }
    const built = isPdf ? await EDD556.buildFromPages(JSZip, SHELLS.s556, results, pdf556.numPages) : await EDD556.build(JSZip, a556, results);
    out556 = built.bytes;
    $('prog556').textContent = `${stopped ? 'تم الإيقاف. ' : ''}قُرئت ${readOk} ${isPdf ? 'صفحة' : 'صورة'} من ${imgs.length}` + (failed ? `، و${failed} لم تُقرأ (${isPdf ? 'عليها ملاحظة في الملف' : 'بقيت صوراً في الملف'})` : '') + '.';
    const rep = $('rep556'); rep.innerHTML = '';
    for (const r of built.report) {
      const li = document.createElement('li');
      li.textContent = `جدول ${r.n}: ${r.rows} سطر — ` + (r.issues.length ? '⚠ ' + r.issues.join('؛ ') : 'المجموع مطابق أو لا يوجد سطر مجموع');
      if (r.issues.length) li.style.color = 'var(--warn)';
      rep.appendChild(li);
    }
    rep.hidden = false; $('dl556wrap').hidden = false;
  });
  $('dl556').addEventListener('click', () => out556 && save(out556, name556 + ' - عقد جدول الوصفي للتقسيم.docx'));
})();
