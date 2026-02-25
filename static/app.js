/* DL sporttools - frontend (geen tracking, geen third-party) */

(function () {
  'use strict';

  // ----------------------------
  // Helpers
  // ----------------------------
  function filenameFromContentDisposition(res) {
    const cd = res.headers.get('Content-Disposition') || res.headers.get('content-disposition');
    if (!cd) return null;

    // RFC 5987 / RFC 6266: filename*=UTF-8''...
    const star = cd.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
    if (star && star[1]) {
      try { return decodeURIComponent(star[1]); } catch (e) { return star[1]; }
    }

    const basic = cd.match(/filename\s*=\s*"([^"]+)"/i);
    if (basic && basic[1]) return basic[1];

    return null;
  }

  function yyyymmddAmsterdam() {
    const now = new Date();
    const parts = new Intl.DateTimeFormat('nl-NL', {
      timeZone: 'Europe/Amsterdam',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    }).formatToParts(now);

    const y = (parts.find(p => p.type === 'year') || {}).value || '1970';
    const m = (parts.find(p => p.type === 'month') || {}).value || '01';
    const d = (parts.find(p => p.type === 'day') || {}).value || '01';
    return `${y}${m}${d}`;
  }

  function fallbackFilenameForAction(action) {
    let path = action;
    try { path = new URL(action).pathname; } catch (e) {}

    const date = yyyymmddAmsterdam();
    const map = {
      '/convert/amateur': `${date}_cue_print_uitslagen_amateurs.txt`,
      '/convert/amateur-online': `${date}_cue_word_uitslagen_amateurs.docx`,
      '/convert/regiosport': `${date}_cue_print_uitslagen_regiosport.txt`,
      '/convert/topscorers': `${date}_cue_word_topscorers_amateurs.docx`,
      '/convert/topscorers-cumulated': `${date}_cue_word_gecumuleerde_topscorers_amateurs.docx`,
    };
    return map[path] || null;
  }

  function isDocxResponse(res) {
    const ct = (res.headers.get('Content-Type') || '').toLowerCase();
    return ct.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename || 'export';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function handleConvertForm(form) {
    const targetId = form.getAttribute('data-target');
    const target = targetId ? document.getElementById(targetId) : null;
    if (target) target.value = '';

    const fd = new FormData(form);
    const res = await fetch(form.action, { method: 'POST', body: fd });

    if (!res.ok) {
      const msg = await res.text();
      alert(msg || 'Converteren is mislukt.');
      return;
    }

    const fname =
      filenameFromContentDisposition(res) ||
      fallbackFilenameForAction(form.action) ||
      form.getAttribute('data-filename') ||
      'export';

    if (isDocxResponse(res)) {
      const blob = await res.blob();
      downloadBlob(blob, fname);
      return;
    }

    const text = await res.text();
    if (target) target.value = text;

    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    downloadBlob(blob, fname);
  }

  // ----------------------------
  // Standaard converters (forms)
  // ----------------------------
  document.querySelectorAll('form.js-convert').forEach(form => {
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      handleConvertForm(form);
    });
  });

  // Bestandsknop label updaten
  document.querySelectorAll('[data-filepicker] input[type=file]').forEach(inp => {
    const wrap = inp.closest('[data-filepicker]');
    const label = wrap ? wrap.querySelector('.file__label') : null;
    const defaultLabel = wrap ? wrap.getAttribute('data-default-label') : null;

    inp.addEventListener('change', () => {
      if (!label) return;
      if (inp.files && inp.files[0]) {
        label.textContent = inp.files[0].name;
      } else {
        label.textContent = defaultLabel || 'Bestand kiezen';
      }
    });
  });

  // Copy-knoppen
  document.querySelectorAll('.copy-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const targetId = btn.getAttribute('data-target');
      const ta = targetId ? document.getElementById(targetId) : null;
      if (!ta) return;

      try {
        await navigator.clipboard.writeText(ta.value || '');
        alert('Gekopieerd.');
      } catch (e) {
        ta.select();
        document.execCommand('copy');
        alert('Gekopieerd.');
      }
    });
  });

  // ----------------------------
  // Gecumuleerde topscorers: 2 uploads + 1 export
  // ----------------------------
  function initCumulatedTopscorers() {
    const block = document.querySelector('.js-cumulated-topscorers');
    if (!block) return;

    const pickers = Array.from(block.querySelectorAll('[data-cum-upload]'));
    const exportBtn = block.querySelector('[data-cum-export]');
    const exportAction = block.getAttribute('data-export-action');

    if (!exportBtn || !exportAction || pickers.length < 2) return;

    function setExportEnabled() {
      const okAll = pickers.every(p => p.classList.contains('selected'));
      exportBtn.disabled = !okAll;
    }

    async function uploadOne(picker) {
      const input = picker.querySelector('input[type=file]');
      if (!input || !input.files || !input.files[0]) return;

      picker.classList.remove('selected');
      setExportEnabled();

      const url = picker.getAttribute('data-cum-upload');
      if (!url) return;

      const fd = new FormData();
      const fieldName = input.getAttribute('name') || 'file';
      fd.append(fieldName, input.files[0]);

      const res = await fetch(url, { method: 'POST', body: fd });

      if (!res.ok) {
        let msg = 'Upload is mislukt.';
        try {
          const data = await res.json();
          msg = (data && (data.message || data.code)) ? `${data.code || ''} ${data.message || ''}`.trim() : msg;
        } catch (e) {
          try { msg = await res.text(); } catch (e2) {}
        }
        // reset input
        input.value = '';
        const label = picker.querySelector('.file__label');
        const defaultLabel = picker.getAttribute('data-default-label');
        if (label) label.textContent = defaultLabel || 'Bestand kiezen';
        picker.classList.remove('selected');
        setExportEnabled();
        alert(msg);
        return;
      }

      // succes: markeer picker
      picker.classList.add('selected');
      setExportEnabled();
    }

    function resetUI() {
      pickers.forEach(p => {
        p.classList.remove('selected');
        const input = p.querySelector('input[type=file]');
        if (input) input.value = '';
        const label = p.querySelector('.file__label');
        const defaultLabel = p.getAttribute('data-default-label');
        if (label) label.textContent = defaultLabel || 'Bestand kiezen';
      });
      exportBtn.disabled = true;
    }

    pickers.forEach(picker => {
      const input = picker.querySelector('input[type=file]');
      if (!input) return;
      input.addEventListener('change', () => {
        if (!input.files || !input.files[0]) return;
        uploadOne(picker);
      });
    });

    exportBtn.addEventListener('click', async () => {
      exportBtn.disabled = true;

      const res = await fetch(exportAction, { method: 'POST' });

      // Na exportpoging resetten (server ruimt altijd op)
      if (!res.ok) {
        const msg = await res.text();
        resetUI();
        alert(msg || 'Converteren is mislukt.');
        return;
      }

      const fname =
        filenameFromContentDisposition(res) ||
        fallbackFilenameForAction(exportAction) ||
        'export.docx';

      const blob = await res.blob();
      downloadBlob(blob, fname);
      resetUI();
    });
  }

  initCumulatedTopscorers();
})();
