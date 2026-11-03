(function () {
  const $ = (id) => document.getElementById(id);
  const csrfToken = document.querySelector('meta[name="csrf-token"]')?.getAttribute('content') || '';
  
  let SHELLS = null;
  let analysed = null, result = null, baseName = 'EDD';
  const showErr = (el, t) => { if (el) { el.textContent = t; el.hidden = !t; } };

  // Load shells template JSON once
  async function ensureShells() {
    if (SHELLS) return SHELLS;
    try {
      const resp = await fetch('/static/shells.json');
      SHELLS = await resp.json();
      return SHELLS;
    } catch (e) {
      console.error('Failed to load shells template:', e);
      throw e;
    }
  }

  async function load(file) {
    showErr($('err'), '');
    if ($('done')) $('done').hidden = true;
    if ($('dl')) $('dl').hidden = true;
    result = null;

    if (!file) return;
    if (!/\.(docx|doc)$/i.test(file.name)) {
      showErr($('err'), 'الملف يجب أن يكون Word بصيغة .docx أو .doc.');
      return;
    }

    let buf;
    try {
      buf = await file.arrayBuffer();
    } catch (e) {
      showErr($('err'), 'تعذّر قراءة الملف.');
      return;
    }

    if (/\.doc$/i.test(file.name)) {
      try {
        buf = (await DOC2DOCX.docToDocx(JSZip, buf)).buffer;
      } catch (e) {
        analysed = null;
        if ($('go')) $('go').disabled = true;
        showErr($('err'), e && e.message === 'ENCRYPTED' ? 'هذا الملف محمي بكلمة سر. أزيلي الحماية في Word ثم أعيدي.' : 'تعذّر قراءة ملف .doc هذا. افتحيه في Word واحفظيه باسم .docx ثم أعيدي.');
        return;
      }
    }

    try {
      analysed = await EDD.analyse(JSZip, buf);
    } catch (e) {
      analysed = null;
      if ($('go')) $('go').disabled = true;
      showErr($('err'), 'تعذّر فتح الملف. تأكدي أنه ملف Word سليم بصيغة .docx.');
      return;
    }

    baseName = file.name.replace(/\.(docx|doc)$/i, '');
    if ($('fname')) $('fname').textContent = file.name;

    const nP = analysed.blocks.filter((b) => b.t === 'p' && b.x.replace(/<[^>]+>/g, '').trim()).length;
    const nT = analysed.blocks.filter((b) => b.t === 'tbl').length;

    if ($('stats')) {
      $('stats').innerHTML = '';
      [[nP + ' فقرة'], [nT + ' جدول'], analysed.boxed ? ['الملف داخل الإطار أصلاً'] : null].filter(Boolean).forEach(([t]) => {
        const s = document.createElement('span');
        s.className = 'chip';
        s.textContent = t;
        $('stats').appendChild(s);
      });
    }
    if ($('loaded')) $('loaded').hidden = false;

    const b = analysed.box || {};
    if ($('fFihris')) $('fFihris').value = b.fihris || '';
    if ($('fDate')) $('fDate').value = '';
    if ($('fQism')) $('fQism').value = b.qism || '';
    if ($('fBal')) $('fBal').value = b.baladiya || '';
    if ($('fMaj')) $('fMaj').value = b.majmoua || '';
    if ($('fHissa')) $('fHissa').value = '';
    if ($('fWil')) $('fWil').value = b.wilaya || '';
    if ($('detHint')) $('detHint').textContent = 'مُلئت الخانات من نص الملف. تأكدي منها وصحّحيها إن لزم، خاصة البلدية.';
    if ($('go')) $('go').disabled = false;
  }

  if ($('inp')) $('inp').addEventListener('change', (e) => load(e.target.files[0]));

  const drop = $('drop');
  if (drop) {
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => load(e.dataTransfer.files[0]));
  }

  if ($('go')) {
    $('go').addEventListener('click', async () => {
      if (!analysed) return;
      $('go').disabled = true;
      $('go').textContent = 'جارٍ الإنشاء…';

      try {
        const shells = await ensureShells();
        result = await EDD.convert(JSZip, analysed, { tpl: shells.tpl }, {
          qism: $('fQism') ? $('fQism').value : '',
          baladiya: $('fBal') ? $('fBal').value : '',
          majmoua: $('fMaj') ? $('fMaj').value : '',
          hissa: $('fHissa') ? $('fHissa').value : '',
          wilaya: $('fWil') ? $('fWil').value : '',
          date: $('fDate') ? $('fDate').value : '',
          fihris: $('fFihris') ? $('fFihris').value : '',
          pageLayout: true,
          idaHajm: $('fIdaHajm') ? $('fIdaHajm').value : '',
          idaRaqm: $('fIdaRaqm') ? $('fIdaRaqm').value : '',
          fi: $('fFi') ? $('fFi').value : '',
          mujallad: $('fMujallad') ? $('fMujallad').value : '',
          mujRaqm: $('fMujRaqm') ? $('fMujRaqm').value : '',
          rasm: $('fRasm') ? $('fRasm').value : ''
        });

        const s = result.stats;
        $('done').textContent = `تمّ: ${s.paras} فقرة، ${s.tables} جدول (${s.rows} سطر)، و«-----» في آخر ${s.dashes} فقرة. حمّلي الملف وافتحيه في Word للتأكد.`;
        $('done').hidden = false;
        $('dl').hidden = false;

        // Log metadata to server
        fetch('/api/log', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'X-CSRFToken': csrfToken
          },
          body: JSON.stringify({
            original_file_name: baseName + (analysed.isDoc ? '.doc' : '.docx'),
            page_count_estimate: s.paras,
            tables: s.tables
          })
        }).catch(err => console.log('Metadata log error:', err));

      } catch (e) {
        console.error('Conversion error:', e);
        showErr($('dlErr'), 'حدث خطأ أثناء التحويل. تأكدي من سلامة الملف.');
      }

      $('go').disabled = false;
      $('go').textContent = 'إنشاء الملف من جديد';
    });
  }

  async function save(bytes, name) {
    showErr($('dlErr'), '');
    try {
      const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 4000);
    } catch (e) {
      showErr($('dlErr'), 'لم يتم التحميل. أعيدي المحاولة.');
    }
  }

  if ($('dlTpl')) {
    $('dlTpl').addEventListener('click', () => {
      if (result) save(result.template, baseName + ' - إجراء إشهار عقاري.docx');
    });
  }

})();
