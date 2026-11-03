(function () {
  // entry code: only its SHA-256 (with a fixed prefix) is in the page, never the code itself
  const HASH = '6403ad4c6893e4b2fb57cc2ebb994b71d8993c3dd01145cf864693bcfa818748';
  const form = document.getElementById('gateForm'), inp = document.getElementById('code'), err = document.getElementById('codeErr');
  let tries = 0, waitUntil = 0;
  const hex = (buf) => Array.from(new Uint8Array(buf), (b) => b.toString(16).padStart(2, '0')).join('');
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (Date.now() < waitUntil) { err.textContent = 'محاولات كثيرة. انتظري قليلاً ثم أعيدي.'; err.hidden = false; return; }
    let ok = false;
    try { ok = hex(await crypto.subtle.digest('SHA-256', new TextEncoder().encode('djebari-edd:' + inp.value.trim()))) === HASH; } catch (x) { ok = false; }
    if (ok) { document.getElementById('gate').hidden = true; document.getElementById('app').hidden = false; inp.value = ''; return; }
    tries++; inp.value = ''; inp.focus();
    if (tries >= 5) { waitUntil = Date.now() + 30000; tries = 0; err.textContent = 'رمز خاطئ 5 مرات. انتظري 30 ثانية.'; }
    else err.textContent = 'الرمز غير صحيح.';
    err.hidden = false;
  });
  inp.focus();
})();
(function () {
  // two tools as tabs; the last one used is remembered on this computer
  const tabs = [['tabTpl', 'panelTpl'], ['tab556', 'panel556']];
  const show = (id) => {
    for (const [t, p] of tabs) {
      const on = t === id;
      document.getElementById(t).setAttribute('aria-selected', on ? 'true' : 'false');
      document.getElementById(p).hidden = !on;
    }
    try { localStorage.setItem('edd_tab', id); } catch (e) {}
  };
  for (const [t] of tabs) document.getElementById(t).addEventListener('click', () => show(t));
  let last = null; try { last = localStorage.getItem('edd_tab'); } catch (e) {}
  if (last && tabs.some(([t]) => t === last)) show(last);
})();
