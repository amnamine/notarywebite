(function () {
  const $ = (id) => document.getElementById(id);
  const SHELLS = JSON.parse(document.getElementById('shells').textContent);
  let analysed = null, result = null, baseName = 'EDD', a556 = null;
  const showErr = (el, t) => { el.textContent = t; el.hidden = !t; };

  async function load(file) {
    showErr($('err'), ''); $('done').hidden = true; $('dl').hidden = true; result = null;
    if (!file) return;
    if (!/\.(docx|doc)$/i.test(file.name)) {
      showErr($('err'), 'الملف يجب أن يكون Word بصيغة .docx أو .doc.');
      return;
    }
    let buf = await file.arrayBuffer();
    if (/\.doc$/i.test(file.name)) {
      // old Word format: read it here and turn it into .docx first
      try { buf = (await DOC2DOCX.docToDocx(JSZip, buf)).buffer; }
      catch (e) {
        analysed = null; $('go').disabled = true;
        showErr($('err'), e && e.message === 'ENCRYPTED' ? 'هذا الملف محمي بكلمة سر. أزيلي الحماية في Word ثم أعيدي.' : 'تعذّر قراءة ملف .doc هذا. افتحيه في Word واحفظيه باسم .docx ثم أعيدي.');
        return;
      }
    }
    try {
      analysed = await EDD.analyse(JSZip, buf);
    } catch (e) {
      analysed = null; $('go').disabled = true;
      showErr($('err'), 'تعذّر فتح الملف. تأكدي أنه ملف Word سليم بصيغة .docx.');
      return;
    }
    baseName = file.name.replace(/\.(docx|doc)$/i, '');
    $('fname').textContent = file.name;
    const nP = analysed.blocks.filter((b) => b.t === 'p' && b.x.replace(/<[^>]+>/g, '').trim()).length;
    const nT = analysed.blocks.filter((b) => b.t === 'tbl').length;
    $('stats').innerHTML = '';
    [[nP + ' فقرة'], [nT + ' جدول'], analysed.boxed ? ['الملف داخل الإطار أصلاً'] : null].filter(Boolean).forEach(([t]) => {
      const s = document.createElement('span'); s.className = 'chip'; s.textContent = t; $('stats').appendChild(s);
    });
    $('loaded').hidden = false;
    const b = analysed.box;
    $('fFihris').value = b.fihris || ''; $('fDate').value = '';
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
        date: $('fDate').value, fihris: $('fFihris').value,
        pageLayout: true,
        idaHajm: $('fIdaHajm').value, idaRaqm: $('fIdaRaqm').value, fi: $('fFi').value,
        mujallad: $('fMujallad').value, mujRaqm: $('fMujRaqm').value, rasm: $('fRasm').value,
      });
      const s = result.stats;
      $('done').textContent = `تمّ: ${s.paras} فقرة، ${s.tables} جدول (${s.rows} سطر)، و«-----» في آخر ${s.dashes} فقرة. حمّلي الملف وافتحيه في Word للتأكد.`;
      $('done').hidden = false; $('dl').hidden = false;
    } catch (e) {
      showErr($('dlErr'), 'حدث خطأ أثناء التحويل. أرسلي هذا الملف إلى المطوّر ليصلح المحوّل.');
    }
    $('go').disabled = false; $('go').textContent = 'إنشاء الملف من جديد';
  });

  async function save(bytes, name) {
    showErr($('dlErr'), '');
    try {
      const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
      const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 4000);
    } catch (e) { showErr($('dlErr'), 'لم يتم التحميل. أعيدي المحاولة.'); }
  }
  $('dlTpl').addEventListener('click', () => result && save(result.template, baseName + ' - إجراء إشهار عقاري.docx'));

})();
