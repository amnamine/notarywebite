(function () {
  const $ = (id) => document.getElementById(id);
  const csrfToken = document.querySelector('meta[name="csrf-token"]')?.getAttribute('content') || '';

  // ---------------- TOOL A: إجــراء إشهار عقاري ----------------
  let fileA = null;

  async function handleFileA(file) {
    $('errA').hidden = true;
    $('doneA').hidden = true;
    $('dlErrA').hidden = true;
    $('btnGoA').disabled = true;

    if (!file) return;
    if (file.name.toLowerCase().endsWith('.doc')) {
      $('errA').textContent = 'هذا ملف .doc قديم. افتحه في Word ثم «حفظ باسم» واختر .docx، وضعه هنا من جديد.';
      $('errA').hidden = false;
      return;
    }
    if (!file.name.toLowerCase().endsWith('.docx')) {
      $('errA').textContent = 'الملف يجب أن يكون Word بصيغة .docx.';
      $('errA').hidden = false;
      return;
    }

    fileA = file;
    const formData = new FormData();
    formData.append('file', file);

    try {
      const resp = await fetch('/api/ijraa/analyse', {
        method: 'POST',
        headers: { 'X-CSRFToken': csrfToken },
        body: formData
      });
      const data = await resp.json();
      if (!resp.ok) {
        $('errA').textContent = data.error || 'تعذّر فحص الملف';
        $('errA').hidden = false;
        return;
      }

      $('fnameA').textContent = data.filename;
      $('statsA').innerHTML = `
        <span class="chip">${data.paras} فقرة</span>
        <span class="chip">${data.tables} جدول</span>
        ${data.boxed ? '<span class="chip">الملف داخل الإطار أصلاً</span>' : ''}
      `;
      $('loadedA').hidden = false;

      // Fill cadastral detected fields
      if (data.box) {
        $('fQism').value = data.box.qism || '';
        $('fBal').value = data.box.baladiya || '';
        $('fMaj').value = data.box.majmoua || '';
        $('fWil').value = data.box.wilaya || '';
      }

      $('btnGoA').disabled = false;
    } catch (e) {
      $('errA').textContent = 'خطأ في الاتصال بالخادم.';
      $('errA').hidden = false;
    }
  }

  $('inpA').addEventListener('change', (e) => handleFileA(e.target.files[0]));

  const dropA = $('dropA');
  ['dragenter', 'dragover'].forEach(ev => dropA.addEventListener(ev, (e) => { e.preventDefault(); dropA.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => dropA.addEventListener(ev, (e) => { e.preventDefault(); dropA.classList.remove('over'); }));
  dropA.addEventListener('drop', (e) => { handleFileA(e.dataTransfer.files[0]); });

  $('btnGoA').addEventListener('click', async () => {
    if (!fileA) return;
    $('btnGoA').disabled = true;
    $('btnGoA').textContent = 'جارٍ الإنشاء…';

    const formData = new FormData();
    formData.append('file', fileA);
    formData.append('qism', $('fQism').value);
    formData.append('baladiya', $('fBal').value);
    formData.append('majmoua', $('fMaj').value);
    formData.append('hissa', $('fHissa').value);
    formData.append('wilaya', $('fWil').value);
    formData.append('idaHajm', $('fIdaHajm').value);
    formData.append('idaRaqm', $('fIdaRaqm').value);
    formData.append('fi', $('fFi').value);
    formData.append('mujallad', $('fMujallad').value);
    formData.append('mujRaqm', $('fMujRaqm').value);
    formData.append('rasm', $('fRasm').value);

    try {
      const resp = await fetch('/api/ijraa/convert', {
        method: 'POST',
        headers: { 'X-CSRFToken': csrfToken },
        body: formData
      });

      if (!resp.ok) {
        const err = await resp.json();
        $('dlErrA').textContent = err.error || 'حدث خطأ أثناء الإنشاء.';
        $('dlErrA').hidden = false;
      } else {
        const blob = await resp.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const baseName = fileA.name.replace(/\.docx$/i, '');
        a.download = `${baseName} - إجراء إشهار عقاري.docx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        $('doneA').textContent = 'تم إنشاء الملف وتحميله بنجاح.';
        $('doneA').hidden = false;
      }
    } catch (e) {
      $('dlErrA').textContent = 'تعذّر إنشاء الملف.';
      $('dlErrA').hidden = false;
    } finally {
      $('btnGoA').disabled = false;
      $('btnGoA').textContent = 'إنشاء الملف من جديد';
    }
  });
})();
