// Consent print window. Fills the provider, patient (page 1 and printed name), reason and witness printed name, then prints.
// The fields arrive by message from the ESK Tracker panel (Office dialog messaging, or postMessage from the same origin)
// and live only in this page's memory. Nothing is written to the URL, storage, the console, or the workbook.
(function () {
  'use strict';
  var FIELDS = ['provider', 'patient', 'reason', 'witness'];
  var filled = false;

  function fill(d) {
    document.querySelectorAll('[data-field]').forEach(function (el) {
      var k = el.getAttribute('data-field');
      el.textContent = d && FIELDS.indexOf(k) >= 0 && typeof d[k] === 'string' ? d[k] : '';
    });
    filled = !!(d && d.patient);
    var p = document.getElementById('cPrint'); if (p) p.disabled = !filled;
    var w = document.getElementById('cWait'); if (w) w.hidden = filled;
  }

  var viaPost = false;   // true when the panel framed this page itself (full-screen view / demo) and sent the fields by postMessage
  function closeWin() {
    fill(null);
    if (viaPost && window.parent !== window) { window.parent.postMessage({ ev: 'consentClose' }, location.origin); return; }
    if (window.Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
      try { Office.context.ui.messageParent(JSON.stringify({ ev: 'consentClose' })); return; } catch (e) { /* not a dialog */ }
    }
    window.close();
  }

  function toolbar() {
    var bar = document.createElement('div');
    bar.className = 'noprint';
    bar.setAttribute('style', 'position:sticky;top:0;z-index:5;display:flex;gap:8px;align-items:center;justify-content:center;padding:8px;background:#1f3864;color:#fff;font:14px "Segoe UI",sans-serif');
    bar.innerHTML = '<span id="cWait">Waiting for the form details from the ESK Tracker…</span>'
      + '<button id="cPrint" disabled style="font:inherit;padding:6px 14px;border:0;border-radius:6px;background:#fff;color:#1f3864;font-weight:700;cursor:pointer">Print</button>'
      + '<button id="cClose" style="font:inherit;padding:6px 14px;border:1px solid #fff;border-radius:6px;background:transparent;color:#fff;cursor:pointer">Close</button>';
    document.body.insertBefore(bar, document.body.firstChild);
    document.getElementById('cPrint').onclick = function () { if (filled) window.print(); };
    document.getElementById('cClose').onclick = closeWin;
  }

  function onData(raw) {
    var m; try { m = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch (e) { return; }
    if (m && m.ev === 'consentData') fill(m.data);
  }

  toolbar();
  fill(window.__consentTest || null);   // the render test sets fictional values here

  // Opened as an Office dialog from the task pane
  if (window.Office && Office.onReady) {
    Office.onReady(function () {
      var ui = Office.context && Office.context.ui;
      if (!ui || !ui.addHandlerAsync || !ui.messageParent) return;
      try {
        ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, function (a) { onData(a.message); }, function () {
          ui.messageParent(JSON.stringify({ ev: 'consentReady' }));
        });
      } catch (e) { /* not opened as a dialog */ }
    });
  }
  // Framed by the panel's own overlay (full-screen view or demo): same-origin postMessage both ways
  window.addEventListener('message', function (e) { if (e.origin === location.origin && e.source === window.parent) { viaPost = true; onData(e.data); } });
  if (window.parent !== window) { try { window.parent.postMessage({ ev: 'consentReady' }, location.origin); } catch (e) { /* parent on another origin (Office dialog): ignored */ } }
  window.addEventListener('pagehide', function () { fill(null); });
})();
