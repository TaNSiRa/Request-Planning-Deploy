// Draws one Word or Excel file. See office.html for why this runs
// sandboxed, and preview-common.js for how the file arrives.
(function () {
  'use strict';

  var P = window.RapPreview;
  var app = document.getElementById('app');

  P.showStatus(app, 'loading', 'Opening file…');

  P.onLoad(function (message) {
    var render = { sheet: renderSheet, doc: renderDoc }[message.kind];
    Promise.resolve()
      .then(function () {
        if (!render) throw new Error('No viewer for "' + message.kind + '"');
        return render(P.bytesFrom(message.dataUrl));
      })
      .catch(function (err) { P.fail(app, err); });
  });

  // Word renders at its natural page width, then CSS zoom fits that to the
  // frame. Fit never enlarges a page past `maxFit` (100% for Word — text just
  // gets huge).
  function fitToWidth(scroller, content, naturalWidth, maxFit) {
    var manual = null;
    function effective() {
      if (manual != null) return manual;
      var fit = (scroller.clientWidth - 24) / naturalWidth;
      return Math.max(0.25, Math.min(maxFit, fit));
    }
    var bar = P.zoomBar(app, scroller, {
      get: effective,
      set: function (scale) {
        var ratio = scroller.scrollTop / Math.max(1, scroller.scrollHeight);
        manual = scale;
        apply();
        scroller.scrollTop = ratio * scroller.scrollHeight;
      }
    });
    function apply() {
      content.style.zoom = effective();
      bar.update();
    }
    new ResizeObserver(apply).observe(scroller);
    apply();
  }

  function stage() {
    var scroller = P.el('div', 'scroller');
    var host = P.el('div', 'zoom-host');
    scroller.appendChild(host);
    app.appendChild(scroller);
    return { scroller: scroller, host: host };
  }

  // ---- Word ---------------------------------------------------------------
  function renderDoc(bytes) {
    var s = stage();
    return docx.renderAsync(bytes, s.host, null, {
      className: 'docx',
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: true,
      experimental: true,
      // An altChunk is a raw HTML part inside the .docx — an easy way to smuggle
      // markup in, and rare in real documents.
      renderAltChunks: false,
      renderComments: false,
      renderChanges: false
    }).then(function () {
      P.clearStatus(app);
      fitToWidth(s.scroller, s.host, s.host.scrollWidth, 1);
    });
  }

  // ---- Excel --------------------------------------------------------------
  // A plain grid: values as Excel formats them, merged cells, column widths,
  // solid fills, a tab per visible sheet. No formulas are recalculated — the
  // value shown is the one Excel last saved.
  var MAX_COLS = 200;
  var MAX_CELLS = 60000;

  function renderSheet(bytes) {
    var book = XLSX.read(bytes, { type: 'array', cellStyles: true, cellDates: true });
    document.body.classList.add('sheet');

    var notice = P.el('div', 'notice');
    notice.hidden = true;
    var scroller = P.el('div', 'scroller');
    var tabs = P.el('div', 'tabs');
    app.appendChild(notice);
    app.appendChild(scroller);
    app.appendChild(tabs);

    var meta = (book.Workbook && book.Workbook.Sheets) || [];
    var sheets = book.SheetNames
      .map(function (name, i) { return { name: name, hidden: meta[i] && meta[i].Hidden }; })
      .filter(function (s) { return !s.hidden; });
    if (!sheets.length) sheets = book.SheetNames.map(function (name) { return { name: name }; });

    var buttons = sheets.map(function (sheet) {
      var b = P.el('button', 'tab', sheet.name);
      b.type = 'button';
      b.onclick = function () { show(sheet); };
      tabs.appendChild(b);
      return b;
    });

    function show(sheet) {
      buttons.forEach(function (b, i) { b.classList.toggle('active', sheets[i] === sheet); });
      var built = buildGrid(book.Sheets[sheet.name]);
      scroller.textContent = '';
      scroller.scrollTop = 0;
      scroller.scrollLeft = 0;
      scroller.appendChild(built.node);
      notice.hidden = !built.note;
      notice.textContent = built.note || '';
    }

    P.clearStatus(app);
    show(sheets[0]);
  }

  function columnWidth(col) {
    if (!col) return 80;
    if (col.wpx) return Math.round(col.wpx);
    if (col.wch) return Math.round(col.wch * 7 + 10);
    if (col.width) return Math.round(col.width * 7);
    return 80;
  }

  function buildGrid(ws) {
    var ref = ws && ws['!ref'];
    if (!ref) return { node: P.el('div', 'empty-sheet', 'This sheet has no cells to show.') };

    var range = XLSX.utils.decode_range(ref);
    // Start at A1 whatever the used range says, so row numbers and column
    // letters match what the person sees in Excel.
    var lastCol = Math.min(range.e.c, MAX_COLS - 1);
    var cols = lastCol + 1;
    var lastRow = Math.min(range.e.r, Math.max(200, Math.floor(MAX_CELLS / cols)) - 1);
    var note = range.e.r > lastRow || range.e.c > lastCol
      ? 'Showing rows 1–' + (lastRow + 1) + ' and columns A–' + XLSX.utils.encode_col(lastCol) +
        ' of ' + ref + '. Download the file to see the rest.'
      : '';

    var spans = {};
    var covered = {};
    (ws['!merges'] || []).forEach(function (m) {
      if (m.s.r > lastRow || m.s.c > lastCol) return;
      var er = Math.min(m.e.r, lastRow);
      var ec = Math.min(m.e.c, lastCol);
      spans[m.s.r + ':' + m.s.c] = { rows: er - m.s.r + 1, cols: ec - m.s.c + 1 };
      for (var r = m.s.r; r <= er; r++) {
        for (var c = m.s.c; c <= ec; c++) {
          if (r !== m.s.r || c !== m.s.c) covered[r + ':' + c] = true;
        }
      }
    });

    var table = P.el('table', 'grid');
    var colgroup = P.el('colgroup');
    var rowHeaderCol = P.el('col');
    rowHeaderCol.style.width = Math.max(44, String(lastRow + 1).length * 8 + 16) + 'px';
    colgroup.appendChild(rowHeaderCol);
    var colMeta = ws['!cols'] || [];
    for (var c = 0; c <= lastCol; c++) {
      var col = P.el('col');
      col.style.width = columnWidth(colMeta[c]) + 'px';
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);

    var head = P.el('thead');
    var headRow = P.el('tr');
    headRow.appendChild(P.el('th', 'corner'));
    for (c = 0; c <= lastCol; c++) headRow.appendChild(P.el('th', null, XLSX.utils.encode_col(c)));
    head.appendChild(headRow);
    table.appendChild(head);

    var body = P.el('tbody');
    var rowMeta = ws['!rows'] || [];
    for (var r = 0; r <= lastRow; r++) {
      var tr = P.el('tr');
      if (rowMeta[r] && rowMeta[r].hpx) tr.style.height = Math.round(rowMeta[r].hpx) + 'px';
      tr.appendChild(P.el('th', null, String(r + 1)));
      for (c = 0; c <= lastCol; c++) {
        var key = r + ':' + c;
        if (covered[key]) continue;
        var td = P.el('td');
        var span = spans[key];
        if (span) {
          td.rowSpan = span.rows;
          td.colSpan = span.cols;
          td.style.textAlign = 'center';
          td.style.verticalAlign = 'middle';
        }
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        if (cell) {
          var text = XLSX.utils.format_cell(cell);
          td.textContent = text;
          if (text) td.title = text;
          if (cell.t === 'n' || cell.t === 'd') td.classList.add('num');
          else if (cell.t === 'b') td.classList.add('bool');
          else if (cell.t === 'e') td.classList.add('err');
          var style = cell.s;
          if (style && style.patternType === 'solid' && style.fgColor && style.fgColor.rgb) {
            td.style.background = '#' + String(style.fgColor.rgb).slice(-6);
          }
        }
        tr.appendChild(td);
      }
      body.appendChild(tr);
    }
    table.appendChild(body);
    return { node: table, note: note };
  }
})();
