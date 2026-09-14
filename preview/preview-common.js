// Shared by the two attachment viewers, office.html and pdf.html. Both run in
// an iframe the Flutter app opens from an attachment tile — the other half of
// this conversation is showDocumentPreview() in lib/shared/attachments.dart.
//
// Everything travels by postMessage:
//   frame -> app  preview-ready     the page is up and listening
//   app -> frame  preview-load      { kind, name, dataUrl }
//   frame -> app  preview-activity  the mouse moved in here (throttled). The app
//                                   never sees pointer events that land in an
//                                   iframe, so without this the idle auto-logout
//                                   fires on someone reading a long document.
//   frame -> app  preview-close     Escape pressed while the frame had focus
//   frame -> app  preview-error     rendering failed (the frame says so itself)
// Every message carries the token from this page's #hash, and both sides drop
// a message that doesn't.
(function () {
  'use strict';

  var token = location.hash.slice(1);

  function post(type, extra) {
    var message = { rap: type, token: token };
    if (extra) for (var key in extra) message[key] = extra[key];
    // '*' because office.html is sandboxed into an opaque origin, which no
    // concrete target origin can name. Nothing secret goes this way.
    parent.postMessage(message, '*');
  }

  var lastActivity = 0;
  function activity() {
    var now = Date.now();
    if (now - lastActivity < 5000) return;
    lastActivity = now;
    post('preview-activity');
  }
  ['pointermove', 'pointerdown', 'wheel', 'keydown', 'touchstart'].forEach(function (type) {
    window.addEventListener(type, activity, { capture: true, passive: true });
  });
  window.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') post('preview-close');
  });

  // A hyperlink inside a document must never navigate this frame away from the
  // viewer. In-document bookmarks (#...) still scroll; everything else is inert.
  document.addEventListener('click', function (e) {
    var link = e.target && e.target.closest ? e.target.closest('a[href]') : null;
    if (!link) return;
    e.preventDefault();
    var href = link.getAttribute('href') || '';
    if (href.charAt(0) !== '#' || href.length < 2) return;
    var id = decodeURIComponent(href.slice(1));
    var target = document.getElementById(id) || document.getElementsByName(id)[0];
    if (target) target.scrollIntoView({ block: 'start' });
  }, true);

  function onLoad(handler) {
    window.addEventListener('message', function (e) {
      if (e.source !== parent) return;
      var m = e.data;
      if (!m || m.rap !== 'preview-load' || m.token !== token) return;
      handler(m);
    });
    post('preview-ready');
  }

  function bytesFrom(dataUrl) {
    var comma = dataUrl.indexOf(',');
    if (dataUrl.slice(0, 5) !== 'data:' || comma < 0) throw new Error('Not a data URL');
    var meta = dataUrl.slice(5, comma);
    var body = dataUrl.slice(comma + 1);
    if (!/;base64$/i.test(meta)) return new TextEncoder().encode(decodeURIComponent(body));
    var binary = atob(body);
    var bytes = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
  }

  function clearStatus(host) {
    var old = host.querySelectorAll(':scope > .status');
    for (var i = 0; i < old.length; i++) old[i].remove();
  }

  function showStatus(host, kind, title, detail) {
    clearStatus(host);
    var box = el('div', 'status ' + kind);
    if (kind === 'loading') box.appendChild(el('div', 'spinner'));
    box.appendChild(el('div', 'status-title', title));
    if (detail) box.appendChild(el('div', 'status-detail', detail));
    host.appendChild(box);
  }

  function fail(host, err, title) {
    console.error(err);
    showStatus(host, 'error', title || "This file can't be previewed",
      'Use the download button above to open it in its own program.');
    post('preview-error', { message: String((err && err.message) || err) });
  }

  // The floating -, 100%, + pill. `o.get()` is the scale in effect right now,
  // `o.set(scale)` applies one, and `o.set(null)` goes back to fit-to-width.
  // Ctrl+wheel zooms too, which is also what a trackpad pinch sends.
  var STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];

  function zoomBar(host, scroller, o) {
    var bar = el('div', 'toolbar');
    function button(text, title) {
      var b = el('button', null, text);
      b.type = 'button';
      b.title = title;
      bar.appendChild(b);
      return b;
    }
    if (o.extra) {
      bar.appendChild(o.extra);
      bar.appendChild(el('span', 'sep'));
    }
    var out = button('−', 'Zoom out');
    var label = button('', 'Fit to width');
    label.classList.add('label');
    var into = button('+', 'Zoom in');

    function step(direction) {
      var current = o.get();
      if (direction > 0) {
        for (var i = 0; i < STEPS.length; i++) if (STEPS[i] > current + 0.001) return STEPS[i];
        return STEPS[STEPS.length - 1];
      }
      for (var j = STEPS.length - 1; j >= 0; j--) if (STEPS[j] < current - 0.001) return STEPS[j];
      return STEPS[0];
    }
    out.onclick = function () { o.set(step(-1)); };
    into.onclick = function () { o.set(step(1)); };
    label.onclick = function () { o.set(null); };

    scroller.addEventListener('wheel', function (e) {
      if (!e.ctrlKey) return;
      e.preventDefault();
      var next = o.get() * Math.exp(-e.deltaY * 0.01);
      o.set(Math.max(STEPS[0], Math.min(STEPS[STEPS.length - 1], next)));
    }, { passive: false });

    host.appendChild(bar);
    return {
      update: function () { label.textContent = Math.round(o.get() * 100) + '%'; }
    };
  }

  window.RapPreview = {
    post: post,
    onLoad: onLoad,
    bytesFrom: bytesFrom,
    el: el,
    showStatus: showStatus,
    clearStatus: clearStatus,
    fail: fail,
    zoomBar: zoomBar
  };
})();
