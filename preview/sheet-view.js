// Lays an Excel sheet out the way Excel itself draws it at 100%: column widths
// and row heights in Excel's pixels, fonts, solid fills, cell borders,
// alignment / wrapping / shrink-to-fit, merged cells, text running on into
// empty neighbours, and the pictures and text boxes of the sheet's drawing
// (signatures, form codes) at their anchor cells.
//
// SheetJS (community) reads values and number formats but almost no styling,
// so the styling is read here straight from the zip: styles.xml, the theme,
// and each sheet's XML. Cell values still come from SheetJS (format_cell), so
// what a cell says is exactly what the plain grid used to say.
//
// Every cell is placed absolutely from the same column / row edges the
// drawings use, so a signature anchored in a cell lands in that cell.
(function () {
  'use strict';

  var P = window.RapPreview;
  var EMU_PER_PX = 9525;
  var HEAD_H = 20;
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' };
  var GRIDLINE = '#e1e1e1';

  // Office 2013+ default theme, used when a workbook carries none.
  var DEFAULT_THEME = ['000000', 'FFFFFF', '44546A', 'E7E6E6', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  // Theme colour index -> clrScheme position (Excel swaps dk/lt pairs).
  var THEME_ORDER = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];
  var DEFAULT_INDEXED = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
  ];
  // Excel border style -> [CSS width px, CSS style].
  var BORDER = {
    thin: [1, 'solid'], medium: [2, 'solid'], thick: [3, 'solid'], double: [3, 'double'],
    hair: [1, 'dotted'], dotted: [1, 'dotted'], dashed: [1, 'dashed'], mediumDashed: [2, 'dashed'],
    dashDot: [1, 'dashed'], mediumDashDot: [2, 'dashed'], dashDotDot: [1, 'dashed'],
    mediumDashDotDot: [2, 'dashed'], slantDashDot: [2, 'dashed']
  };
  var DEFAULT_FONT = { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000' };
  var NO_BORDER = { left: null, right: null, top: null, bottom: null };

  // ---- xml helpers ----------------------------------------------------------
  function xml(text) { return new DOMParser().parseFromString(text, 'application/xml'); }
  function byLocal(node, name) {
    var out = [];
    if (!node) return out;
    var all = node.getElementsByTagName('*');
    for (var i = 0; i < all.length; i++) if (all[i].localName === name) out.push(all[i]);
    return out;
  }
  function child(node, name) {
    for (var c = node && node.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c;
    return null;
  }
  function kids(node, name) {
    var out = [];
    for (var c = node && node.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) out.push(c);
    return out;
  }
  function attr(node, name) { return node ? node.getAttribute(name) : null; }
  function flag(node, name) {
    var v = attr(node, name);
    return v === '1' || v === 'true';
  }
  // <b/>, <b val="1"/> are on; <b val="0"/> is off.
  function on(node) {
    if (!node) return false;
    var v = node.getAttribute('val');
    return v == null || (v !== '0' && v !== 'false');
  }
  function num(node, name) {
    var c = child(node, name);
    return c ? Number(c.textContent) || 0 : 0;
  }
  // 'xl/worksheets' + '../drawings/drawing1.xml' -> 'xl/drawings/drawing1.xml'
  function resolve(dir, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = dir ? dir.split('/') : [];
    target.split('/').forEach(function (p) {
      if (p === '..') parts.pop();
      else if (p && p !== '.') parts.push(p);
    });
    return parts.join('/');
  }
  function dirOf(path) { return path.slice(0, path.lastIndexOf('/')); }
  function readText(zip, path) {
    var file = path && zip.file(path);
    return file ? file.async('string') : Promise.resolve(null);
  }
  function relsOf(zip, path) {
    var relsPath = dirOf(path) + '/_rels/' + path.slice(path.lastIndexOf('/') + 1) + '.rels';
    return readText(zip, relsPath).then(function (text) {
      var map = {};
      if (!text) return map;
      byLocal(xml(text), 'Relationship').forEach(function (r) {
        map[r.getAttribute('Id')] = { target: resolve(dirOf(path), r.getAttribute('Target')), type: r.getAttribute('Type') || '' };
      });
      return map;
    });
  }
  function relOfType(rels, suffix) {
    for (var k in rels) if (rels[k].type.slice(-suffix.length) === suffix) return rels[k];
    return null;
  }
  // 'AB12' -> { r: 11, c: 27 }
  function decodeRef(ref) {
    var m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref || '');
    if (!m) return null;
    var c = 0;
    for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
    return { r: Number(m[2]) - 1, c: c - 1 };
  }
  function decodeRange(ref) {
    var parts = String(ref || '').split(':');
    var s = decodeRef(parts[0]);
    var e = decodeRef(parts[1] || parts[0]);
    return s && e ? { r0: s.r, c0: s.c, r1: e.r, c1: e.c } : null;
  }

  // ---- colours ---------------------------------------------------------------
  function hexToRgb(hex) {
    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
  }
  function rgbToHex(rgb) {
    return rgb.map(function (v) {
      var s = Math.max(0, Math.min(255, Math.round(v))).toString(16);
      return s.length < 2 ? '0' + s : s;
    }).join('').toUpperCase();
  }
  // Excel's tint: darken (tint < 0) or lighten (tint > 0) the HSL lightness.
  function applyTint(hex, tint) {
    var rgb = hexToRgb(hex).map(function (v) { return v / 255; });
    var max = Math.max.apply(null, rgb);
    var min = Math.min.apply(null, rgb);
    var l = (max + min) / 2;
    var h = 0;
    var s = 0;
    if (max !== min) {
      var d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === rgb[0]) h = (rgb[1] - rgb[2]) / d + (rgb[1] < rgb[2] ? 6 : 0);
      else if (max === rgb[1]) h = (rgb[2] - rgb[0]) / d + 2;
      else h = (rgb[0] - rgb[1]) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    function hue(p, q, t) {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    }
    var out;
    if (s === 0) out = [l, l, l];
    else {
      var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      var p = 2 * l - q;
      out = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
    }
    return rgbToHex(out.map(function (v) { return v * 255; }));
  }
  function colorReader(theme, indexed) {
    return function (node, fallback) {
      if (!node || flag(node, 'auto')) return fallback;
      var hex = null;
      if (attr(node, 'rgb')) hex = attr(node, 'rgb').slice(-6);
      else if (attr(node, 'theme') != null) hex = theme[THEME_ORDER[Number(attr(node, 'theme'))]];
      else if (attr(node, 'indexed') != null) hex = indexed[Number(attr(node, 'indexed'))];
      if (!hex || !/^[0-9A-Fa-f]{6}$/.test(hex)) return fallback;
      var tint = Number(attr(node, 'tint')) || 0;
      return '#' + (tint ? applyTint(hex, tint) : hex.toUpperCase());
    };
  }

  function readTheme(text) {
    if (!text) return DEFAULT_THEME;
    var scheme = byLocal(xml(text), 'clrScheme')[0];
    if (!scheme) return DEFAULT_THEME;
    var out = [];
    for (var c = scheme.firstElementChild; c; c = c.nextElementSibling) {
      var srgb = child(c, 'srgbClr');
      var sys = child(c, 'sysClr');
      out.push(srgb ? attr(srgb, 'val') : sys ? attr(sys, 'lastClr') : null);
    }
    return DEFAULT_THEME.map(function (d, i) { return out[i] || d; });
  }

  // ---- styles ----------------------------------------------------------------
  var measureCtx = null;
  function cssFont(font, scale) {
    return (font.italic ? 'italic ' : '') + (font.bold ? 'bold ' : '') + (font.size * (scale || 1)) + 'pt "' + font.name + '", sans-serif';
  }
  function textWidth(text, font) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    measureCtx.font = cssFont(font);
    return measureCtx.measureText(text).width;
  }
  // Excel's "maximum digit width" of the workbook's default font: the unit
  // column widths are counted in.
  function digitWidth(font) {
    var w = 0;
    for (var d = 0; d <= 9; d++) w = Math.max(w, textWidth(String(d), font));
    w = Math.round(w);
    return w >= 4 && w <= 30 ? w : 7;
  }

  function readStyles(text, theme) {
    var root = text ? xml(text).documentElement : null;
    var indexed = DEFAULT_INDEXED.slice();
    var custom = byLocal(root, 'indexedColors')[0];
    if (custom) kids(custom, 'rgbColor').forEach(function (c, i) { indexed[i] = String(attr(c, 'rgb') || '').slice(-6); });
    var color = colorReader(theme, indexed);

    var fonts = kids(child(root, 'fonts'), 'font').map(function (f) {
      var u = child(f, 'u');
      return {
        name: attr(child(f, 'name'), 'val') || 'Calibri',
        size: Number(attr(child(f, 'sz'), 'val')) || 11,
        bold: on(child(f, 'b')),
        italic: on(child(f, 'i')),
        underline: !!u && attr(u, 'val') !== 'none',
        strike: on(child(f, 'strike')),
        color: color(child(f, 'color'), '#000000')
      };
    });
    var fills = kids(child(root, 'fills'), 'fill').map(function (f) {
      var p = child(f, 'patternFill');
      if (p) {
        var type = attr(p, 'patternType');
        return type && type !== 'none' ? color(child(p, 'fgColor'), null) : null;
      }
      var g = child(f, 'gradientFill');
      var stop = g && child(g, 'stop');
      return stop ? color(child(stop, 'color'), null) : null;
    });
    var borders = kids(child(root, 'borders'), 'border').map(function (b) {
      function side(name) {
        var n = child(b, name);
        var style = attr(n, 'style');
        if (!style || !BORDER[style]) return null;
        return { w: BORDER[style][0], style: BORDER[style][1], color: color(child(n, 'color'), '#000000') };
      }
      return {
        left: side('left') || side('start'),
        right: side('right') || side('end'),
        top: side('top'),
        bottom: side('bottom')
      };
    });
    var defaultFont = fonts[0] || DEFAULT_FONT;
    var xfs = kids(child(root, 'cellXfs'), 'xf').map(function (x) {
      var a = child(x, 'alignment');
      return {
        font: fonts[Number(attr(x, 'fontId')) || 0] || defaultFont,
        fill: fills[Number(attr(x, 'fillId')) || 0] || null,
        border: borders[Number(attr(x, 'borderId')) || 0] || NO_BORDER,
        h: attr(a, 'horizontal') || 'general',
        v: attr(a, 'vertical') || 'bottom',
        wrap: flag(a, 'wrapText'),
        shrink: flag(a, 'shrinkToFit'),
        indent: Number(attr(a, 'indent')) || 0
      };
    });
    if (!xfs.length) xfs.push({ font: defaultFont, fill: null, border: NO_BORDER, h: 'general', v: 'bottom', wrap: false, shrink: false, indent: 0 });
    return { xfs: xfs, defaultFont: defaultFont, mdw: digitWidth(defaultFont) };
  }

  // ---- drawings ----------------------------------------------------------------
  function anchorPoint(node) {
    return { col: num(node, 'col'), colOff: num(node, 'colOff'), row: num(node, 'row'), rowOff: num(node, 'rowOff') };
  }
  function readDrawing(zip, path, color) {
    var file = zip.file(path);
    if (!file) return Promise.resolve([]);
    return Promise.all([file.async('string'), relsOf(zip, path)]).then(function (got) {
      var doc = xml(got[0]);
      var rels = got[1];
      var anchors = kids(doc.documentElement, 'twoCellAnchor')
        .concat(kids(doc.documentElement, 'oneCellAnchor'), kids(doc.documentElement, 'absoluteAnchor'));
      // One shape: a picture or a text box, at [box] — its share of the
      // anchor's rectangle ({ x, y, w, h } as fractions).
      function readPic(pic, box) {
        var blip = byLocal(pic, 'blip')[0];
        var rel = blip && rels[blip.getAttribute('r:embed') || blip.getAttributeNS(NS_R, 'embed')];
        var type = rel ? rel.target.slice(rel.target.lastIndexOf('.') + 1).toLowerCase() : '';
        var media = rel && IMAGE_TYPES[type] && zip.file(rel.target);
        if (!media) return null; // EMF / WMF: a browser can't draw them
        return media.async('uint8array').then(function (data) {
          return { box: box, src: URL.createObjectURL(new Blob([data], { type: IMAGE_TYPES[type] })) };
        });
      }
      function readTextBox(sp, box) {
        var body = child(sp, 'txBody');
        if (!body) return null;
        var text = byLocal(body, 'p').map(function (p) {
          return byLocal(p, 't').map(function (t) { return t.textContent; }).join('');
        }).join('\n');
        if (!text.trim()) return null;
        var part = { box: box, text: text };
        var rPr = byLocal(body, 'rPr')[0] || byLocal(body, 'defRPr')[0];
        part.fontSize = rPr && attr(rPr, 'sz') ? Number(attr(rPr, 'sz')) / 100 : 11;
        part.bold = flag(rPr, 'b');
        var latin = rPr && child(rPr, 'latin');
        part.fontName = latin ? attr(latin, 'typeface') : '';
        var textFill = rPr && child(rPr, 'solidFill');
        var textRgb = textFill && child(textFill, 'srgbClr');
        part.color = textRgb ? '#' + attr(textRgb, 'val') : '#000000';
        var pPr = byLocal(body, 'pPr')[0];
        var algn = attr(pPr, 'algn');
        part.align = algn === 'ctr' ? 'center' : algn === 'r' ? 'flex-end' : 'flex-start';
        var bodyPr = child(body, 'bodyPr');
        var anchor = attr(bodyPr, 'anchor');
        part.valign = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
        var spPr = child(sp, 'spPr');
        var fill = spPr && child(spPr, 'solidFill');
        var rgb = fill && child(fill, 'srgbClr');
        part.fill = rgb ? '#' + attr(rgb, 'val') : '';
        var ln = spPr && child(spPr, 'ln');
        var lnFill = ln && child(ln, 'solidFill');
        var lnRgb = lnFill && child(lnFill, 'srgbClr');
        part.line = lnRgb ? '#' + attr(lnRgb, 'val') : '';
        return part;
      }
      function xfrmOf(node) {
        var props = child(node, node.localName === 'grpSp' ? 'grpSpPr' : 'spPr');
        var x = child(props, 'xfrm');
        if (!x) return null;
        function pt(name, a, b) {
          var n = child(x, name);
          return n ? { a: Number(attr(n, a)) || 0, b: Number(attr(n, b)) || 0 } : null;
        }
        return { off: pt('off', 'x', 'y'), ext: pt('ext', 'cx', 'cy'), chOff: pt('chOff', 'x', 'y'), chExt: pt('chExt', 'cx', 'cy') };
      }
      // A shape, or a group's shapes laid out by the group's own coordinates
      // (chOff / chExt map onto wherever the group is drawn).
      function readShapes(node, box, out) {
        if (node.localName === 'pic') out.push(readPic(node, box));
        else if (node.localName === 'sp') out.push(readTextBox(node, box));
        else if (node.localName === 'grpSp') {
          var g = xfrmOf(node);
          var origin = g && (g.chOff || g.off);
          var span = g && (g.chExt || g.ext);
          if (!origin || !span || !span.a || !span.b) return;
          for (var c = node.firstElementChild; c; c = c.nextElementSibling) {
            if (c.localName !== 'pic' && c.localName !== 'sp' && c.localName !== 'grpSp') continue;
            var x = xfrmOf(c);
            if (!x || !x.off || !x.ext) continue;
            readShapes(c, {
              x: box.x + ((x.off.a - origin.a) / span.a) * box.w,
              y: box.y + ((x.off.b - origin.b) / span.b) * box.h,
              w: (x.ext.a / span.a) * box.w,
              h: (x.ext.b / span.b) * box.h
            }, out);
          }
        }
      }

      return Promise.all(anchors.map(function (a) {
        var item = {};
        var from = child(a, 'from');
        var to = child(a, 'to');
        var ext = child(a, 'ext');
        var pos = child(a, 'pos');
        if (from) item.from = anchorPoint(from);
        else if (pos) item.pos = { x: Number(attr(pos, 'x')) / EMU_PER_PX, y: Number(attr(pos, 'y')) / EMU_PER_PX };
        else return null;
        if (a.localName === 'twoCellAnchor' && to) item.to = anchorPoint(to);
        else if (ext) item.size = { w: Number(attr(ext, 'cx')) / EMU_PER_PX, h: Number(attr(ext, 'cy')) / EMU_PER_PX };
        else return null;
        var parts = [];
        for (var c = a.firstElementChild; c; c = c.nextElementSibling) readShapes(c, { x: 0, y: 0, w: 1, h: 1 }, parts);
        return Promise.all(parts).then(function (got) {
          item.parts = got.filter(Boolean);
          return item.parts.length ? item : null;
        });
      })).then(function (items) { return items.filter(Boolean); });
    });
  }

  // ---- sheet -------------------------------------------------------------------
  function readSheet(zip, path) {
    return Promise.all([readText(zip, path), relsOf(zip, path)]).then(function (got) {
      if (!got[0]) return null;
      var root = xml(got[0]).documentElement;
      var rels = got[1];
      var fmt = child(root, 'sheetFormatPr');
      var view = byLocal(child(root, 'sheetViews'), 'sheetView')[0];
      var sheet = {
        defaultColWidth: Number(attr(fmt, 'defaultColWidth')) || 0,
        defaultRowHeight: Number(attr(fmt, 'defaultRowHeight')) || 15,
        gridlines: attr(view, 'showGridLines') !== '0' && attr(view, 'showGridLines') !== 'false',
        cols: kids(child(root, 'cols'), 'col').map(function (c) {
          return {
            min: Number(attr(c, 'min')) - 1,
            max: Number(attr(c, 'max')) - 1,
            width: Number(attr(c, 'width')) || 0,
            hidden: flag(c, 'hidden'),
            style: attr(c, 'style') != null ? Number(attr(c, 'style')) : null
          };
        }),
        rows: {},
        styleAt: {},
        lastRow: -1,
        lastCol: -1,
        merges: kids(child(root, 'mergeCells'), 'mergeCell').map(function (m) { return decodeRange(attr(m, 'ref')); }).filter(Boolean),
        drawing: null
      };
      var dim = decodeRange(attr(child(root, 'dimension'), 'ref'));
      if (dim) {
        sheet.lastRow = dim.r1;
        sheet.lastCol = dim.c1;
      }
      var r = -1;
      kids(child(root, 'sheetData'), 'row').forEach(function (row) {
        r = attr(row, 'r') ? Number(attr(row, 'r')) - 1 : r + 1;
        var meta = {};
        if (attr(row, 'ht') != null) meta.ht = Number(attr(row, 'ht'));
        if (flag(row, 'hidden')) meta.hidden = true;
        if (flag(row, 'customFormat') && attr(row, 's') != null) meta.style = Number(attr(row, 's'));
        sheet.rows[r] = meta;
        var c = -1;
        kids(row, 'c').forEach(function (cell) {
          var at = decodeRef(attr(cell, 'r'));
          c = at ? at.c : c + 1;
          if (attr(cell, 's') != null) sheet.styleAt[r + ':' + c] = Number(attr(cell, 's'));
          if (!dim) {
            sheet.lastRow = Math.max(sheet.lastRow, r);
            sheet.lastCol = Math.max(sheet.lastCol, c);
          }
        });
      });
      // Page setup, for printing (see printView): manual page breaks, paper,
      // orientation, margins and any fit-to-page / scale.
      var setup = child(root, 'pageSetup');
      var fit = child(child(root, 'sheetPr'), 'pageSetUpPr');
      var margins = child(root, 'pageMargins');
      var options = child(root, 'printOptions');
      function margin(name, fallback) {
        var v = attr(margins, name);
        return v != null && isFinite(Number(v)) ? Number(v) : fallback;
      }
      function count(name) {
        var v = attr(setup, name);
        return v == null ? 1 : Number(v) || 0;
      }
      sheet.page = {
        breaks: kids(child(root, 'rowBreaks'), 'brk').map(function (b) { return Number(attr(b, 'id')); })
          .filter(function (n) { return n > 0; }),
        paper: Number(attr(setup, 'paperSize')) || 9,
        landscape: attr(setup, 'orientation') === 'landscape',
        scale: Number(attr(setup, 'scale')) || 100,
        fitToPage: flag(fit, 'fitToPage'),
        fitWidth: count('fitToWidth'),
        fitHeight: count('fitToHeight'),
        margins: { left: margin('left', 0.7), right: margin('right', 0.7), top: margin('top', 0.75), bottom: margin('bottom', 0.75) },
        centered: flag(options, 'horizontalCentered'),
        gridlines: flag(options, 'gridLines')
      };
      var drawingNode = child(root, 'drawing');
      var rid = drawingNode && (drawingNode.getAttribute('r:id') || drawingNode.getAttributeNS(NS_R, 'id'));
      var drawingRel = rid ? rels[rid] : relOfType(rels, '/drawing');
      sheet.drawingPath = drawingRel ? drawingRel.target : null;
      return sheet;
    });
  }

  // { sheets: [{ name, hidden }], draw(name, ws) -> Promise<{ node, width, note }> }
  function open(bytes) {
    return JSZip.loadAsync(bytes).then(function (zip) {
      return Promise.all([readText(zip, 'xl/workbook.xml'), relsOf(zip, 'xl/workbook.xml')]).then(function (got) {
        var wbRels = got[1];
        var wb = got[0] ? xml(got[0]) : null;
        var sheets = byLocal(wb, 'sheet').map(function (sh) {
          var rel = wbRels[sh.getAttribute('r:id') || sh.getAttributeNS(NS_R, 'id')];
          return { name: sh.getAttribute('name'), hidden: (attr(sh, 'state') || 'visible') !== 'visible', path: rel && rel.target };
        });
        // The print area's last column: what "fit to width" fits, so a form
        // opens showing the form rather than the empty columns beside it.
        byLocal(wb, 'definedName').forEach(function (dn) {
          var sheet = sheets[Number(attr(dn, 'localSheetId'))];
          if (attr(dn, 'name') !== '_xlnm.Print_Area' || !sheet) return;
          dn.textContent.split(',').forEach(function (part) {
            var range = decodeRange(part.slice(part.lastIndexOf('!') + 1).replace(/\$/g, ''));
            if (range) sheet.printLastCol = Math.max(sheet.printLastCol || 0, range.c1);
            if (range) {
              var area = sheet.printArea;
              sheet.printArea = area
                ? { r0: Math.min(area.r0, range.r0), c0: Math.min(area.c0, range.c0), r1: Math.max(area.r1, range.r1), c1: Math.max(area.c1, range.c1) }
                : range;
            }
          });
        });
        var stylesRel = relOfType(wbRels, '/styles');
        var themeRel = relOfType(wbRels, '/theme');
        return Promise.all([
          readText(zip, stylesRel ? stylesRel.target : 'xl/styles.xml'),
          readText(zip, themeRel ? themeRel.target : 'xl/theme/theme1.xml')
        ]).then(function (st) {
          var theme = readTheme(st[1]);
          var styles = readStyles(st[0], theme);
          return {
            sheets: sheets,
            draw: function (name, ws, limits) {
              var info = sheets.filter(function (s) { return s.name === name; })[0];
              if (!info || !info.path) return Promise.resolve(null);
              return readSheet(zip, info.path).then(function (sheet) {
                if (!sheet) return null;
                var drawing = sheet.drawingPath
                  ? readDrawing(zip, sheet.drawingPath).catch(function (err) {
                    console.warn('drawings skipped', err);
                    return [];
                  })
                  : Promise.resolve([]);
                return drawing.then(function (items) { return layout(sheet, items, ws, styles, limits, info); });
              });
            }
          };
        });
      });
    });
  }

  // ---- layout --------------------------------------------------------------------
  function layout(sheet, items, ws, styles, limits, info) {
    var printLastCol = info.printLastCol;
    var mdw = styles.mdw;
    // Stored widths already include Excel's 5px of cell padding.
    function widthPx(w) { return Math.floor(((256 * w + Math.floor(128 / mdw)) / 256) * mdw); }
    function heightPx(pt) { return Math.round(pt * 96 / 72); }

    // How far to draw: the used range, every merge, every drawing.
    var lastRow = sheet.lastRow;
    var lastCol = sheet.lastCol;
    sheet.merges.forEach(function (m) {
      lastRow = Math.max(lastRow, m.r1);
      lastCol = Math.max(lastCol, m.c1);
    });
    items.forEach(function (it) {
      [it.from, it.to].forEach(function (p) {
        if (!p) return;
        lastRow = Math.max(lastRow, p.row);
        lastCol = Math.max(lastCol, p.col);
      });
    });
    if (ws && ws['!ref']) {
      var used = decodeRange(ws['!ref']);
      if (used) {
        lastRow = Math.max(lastRow, used.r1);
        lastCol = Math.max(lastCol, used.c1);
      }
    }
    if (lastRow < 0 || lastCol < 0) return null;
    var fullRows = lastRow;
    var fullCols = lastCol;
    lastCol = Math.min(lastCol, limits.maxCols - 1);
    lastRow = Math.min(lastRow, Math.max(200, Math.floor(limits.maxCells / (lastCol + 1))) - 1);
    var note = fullRows > lastRow || fullCols > lastCol
      ? 'Showing rows 1–' + (lastRow + 1) + ' and columns A–' + XLSX.utils.encode_col(lastCol) + '. Download the file to see the rest.'
      : '';

    var colStyle = [];
    var colW = [];
    var defaultW = sheet.defaultColWidth ? widthPx(sheet.defaultColWidth) : Math.round(64 * mdw / 7);
    for (var c = 0; c <= lastCol; c++) {
      colW[c] = defaultW;
      colStyle[c] = null;
    }
    sheet.cols.forEach(function (spec) {
      for (var k = Math.max(0, spec.min); k <= Math.min(spec.max, lastCol); k++) {
        colW[k] = spec.hidden ? 0 : spec.width ? widthPx(spec.width) : defaultW;
        colStyle[k] = spec.style;
      }
    });
    var rowH = [];
    var defaultH = heightPx(sheet.defaultRowHeight);
    for (var r = 0; r <= lastRow; r++) {
      var meta = sheet.rows[r];
      rowH[r] = meta && meta.hidden ? 0 : meta && meta.ht != null ? heightPx(meta.ht) : defaultH;
    }
    var colLeft = [0];
    for (c = 0; c <= lastCol; c++) colLeft.push(colLeft[c] + colW[c]);
    var rowTop = [0];
    for (r = 0; r <= lastRow; r++) rowTop.push(rowTop[r] + rowH[r]);
    var W = colLeft[lastCol + 1];
    var H = rowTop[lastRow + 1];

    var mergeAt = {};
    sheet.merges.forEach(function (m) {
      for (var mr = m.r0; mr <= Math.min(m.r1, lastRow); mr++) {
        for (var mc = m.c0; mc <= Math.min(m.c1, lastCol); mc++) mergeAt[mr + ':' + mc] = m;
      }
    });
    function xfAt(r, c) {
      var s = sheet.styleAt[r + ':' + c];
      if (s == null) s = sheet.rows[r] && sheet.rows[r].style;
      if (s == null) s = colStyle[c];
      return styles.xfs[s || 0] || styles.xfs[0];
    }
    function cellAt(r, c) { return ws ? ws[XLSX.utils.encode_cell({ r: r, c: c })] : null; }
    function textOf(cell) {
      if (!cell) return '';
      var t = XLSX.utils.format_cell(cell);
      return t == null ? '' : String(t);
    }
    function isEmpty(r, c) {
      if (c < 0 || c > lastCol || mergeAt[r + ':' + c]) return false;
      return textOf(cellAt(r, c)) === '';
    }

    var root = P.el('div', 'xl');
    var rowHeadW = Math.max(40, String(lastRow + 1).length * 8 + 14);
    root.style.width = rowHeadW + W + 'px';

    // Column letters and row numbers, pinned while scrolling.
    var colHead = P.el('div', 'xl-colhead');
    var corner = P.el('div', 'xl-corner');
    corner.style.width = rowHeadW + 'px';
    colHead.appendChild(corner);
    var letters = P.el('div', 'xl-letters');
    letters.style.width = W + 'px';
    for (c = 0; c <= lastCol; c++) {
      if (!colW[c]) continue;
      var lh = P.el('div', 'xl-head', XLSX.utils.encode_col(c));
      lh.style.left = colLeft[c] + 'px';
      lh.style.width = colW[c] + 'px';
      letters.appendChild(lh);
    }
    colHead.appendChild(letters);
    root.appendChild(colHead);

    var bodyRow = P.el('div', 'xl-body');
    var numbers = P.el('div', 'xl-numbers');
    numbers.style.width = rowHeadW + 'px';
    numbers.style.height = H + 'px';
    for (r = 0; r <= lastRow; r++) {
      if (!rowH[r]) continue;
      var nh = P.el('div', 'xl-head', String(r + 1));
      nh.style.top = rowTop[r] + 'px';
      nh.style.height = rowH[r] + 'px';
      numbers.appendChild(nh);
    }
    bodyRow.appendChild(numbers);

    var grid = P.el('div', 'xl-sheet');
    grid.style.width = W + 'px';
    grid.style.height = H + 'px';
    bodyRow.appendChild(grid);
    root.appendChild(bodyRow);

    function box(cls, x, y, w, h) {
      var d = P.el('div', cls);
      d.style.left = x + 'px';
      d.style.top = y + 'px';
      d.style.width = w + 'px';
      d.style.height = h + 'px';
      return d;
    }

    // Gridlines: the last pixel column / row of every cell, as in Excel.
    var gridLayer = document.createDocumentFragment();
    if (sheet.gridlines) {
      for (c = 0; c <= lastCol; c++) {
        if (colW[c]) gridLayer.appendChild(box('xl-gl', colLeft[c + 1] - 1, 0, 1, H));
      }
      for (r = 0; r <= lastRow; r++) {
        if (rowH[r]) gridLayer.appendChild(box('xl-gl', 0, rowTop[r + 1] - 1, W, 1));
      }
    }

    var fillLayer = document.createDocumentFragment();
    var textLayer = document.createDocumentFragment();
    var segments = {};
    function segment(dir, at, from, to, side) {
      var key = dir + '|' + at + '|' + side.w + '|' + side.style + '|' + side.color;
      (segments[key] = segments[key] || []).push([from, to]);
    }
    // A border sits centred on the shared edge between two cells, which in
    // Excel is the earlier cell's last pixel.
    function lineAt(edge, w) { return Math.max(0, edge - 1 - Math.floor((w - 1) / 2)); }

    for (r = 0; r <= lastRow; r++) {
      if (!rowH[r]) continue;
      for (c = 0; c <= lastCol; c++) {
        if (!colW[c]) continue;
        var m = mergeAt[r + ':' + c];
        var xf = xfAt(r, c);
        var region = m || { r0: r, c0: c, r1: r, c1: c };
        var bd = xf.border;
        var x0 = colLeft[c];
        var x1 = colLeft[c + 1];
        var y0 = rowTop[r];
        var y1 = rowTop[r + 1];
        if (bd.top && r === region.r0) segment('h', lineAt(y0, bd.top.w), x0 - 1, x1 + Math.floor(bd.top.w / 2), bd.top);
        if (bd.bottom && r === region.r1) segment('h', lineAt(y1, bd.bottom.w), x0 - 1, x1 + Math.floor(bd.bottom.w / 2), bd.bottom);
        if (bd.left && c === region.c0) segment('v', lineAt(x0, bd.left.w), y0 - 1, y1 + Math.floor(bd.left.w / 2), bd.left);
        if (bd.right && c === region.c1) segment('v', lineAt(x1, bd.right.w), y0 - 1, y1 + Math.floor(bd.right.w / 2), bd.right);
        if (m && (r !== m.r0 || c !== m.c0)) continue;

        var bx0 = x0;
        var bx1 = m ? colLeft[Math.min(m.c1, lastCol) + 1] : x1;
        var by0 = y0;
        var by1 = m ? rowTop[Math.min(m.r1, lastRow) + 1] : y1;
        if (xf.fill) {
          var f = box('xl-fill', bx0, by0, bx1 - bx0, by1 - by0);
          f.style.background = xf.fill;
          fillLayer.appendChild(f);
        } else if (m) {
          // A merged cell shows no gridlines inside it.
          fillLayer.appendChild(box('xl-fill xl-blank', bx0, by0, bx1 - bx0 - 1, by1 - by0 - 1));
        }

        var cell = cellAt(r, c);
        var text = textOf(cell);
        if (!text) continue;
        var align = xf.h;
        if (align === 'general') align = cell.t === 'n' || cell.t === 'd' ? 'right' : cell.t === 'b' || cell.t === 'e' ? 'center' : 'left';
        else if (align === 'centerContinuous' || align === 'distributed') align = 'center';
        else if (align === 'fill' || align === 'justify') align = 'left';
        var font = xf.font;
        var scale = 1;
        var oneLine = !xf.wrap;
        var shown = oneLine ? text.replace(/\r?\n/g, ' ') : text;
        if (oneLine) {
          var need = textWidth(shown, font) + 4 + xf.indent * 3 * mdw;
          if (need > bx1 - bx0) {
            if (xf.shrink) scale = Math.max(0.3, (bx1 - bx0 - 4) / (need - 4));
            else if (!m && cell.t !== 'n') {
              // Text runs on over empty neighbours, as Excel draws it.
              var lc = c;
              var rc = c;
              while (colLeft[rc + 1] - colLeft[lc] < need) {
                var grew = false;
                if ((align === 'left' || align === 'center') && isEmpty(r, rc + 1)) { rc++; grew = true; }
                if ((align === 'right' || align === 'center') && isEmpty(r, lc - 1)) { lc--; grew = true; }
                if (!grew) break;
              }
              bx0 = colLeft[lc];
              bx1 = colLeft[rc + 1];
            }
          }
        }
        var t = box('xl-text', bx0, by0, Math.max(0, bx1 - bx0 - 1), Math.max(0, by1 - by0 - 1));
        t.style.justifyContent = xf.v === 'top' ? 'flex-start' : xf.v === 'bottom' ? 'flex-end' : 'center';
        var inner = P.el('div', oneLine ? 'xl-line' : 'xl-wrap', shown);
        inner.style.textAlign = align;
        inner.style.font = cssFont(font, scale);
        inner.style.lineHeight = '1.2';
        inner.style.color = font.color;
        if (font.underline || font.strike) inner.style.textDecoration = (font.underline ? 'underline ' : '') + (font.strike ? 'line-through' : '');
        if (xf.indent) inner.style[align === 'right' ? 'paddingRight' : 'paddingLeft'] = 2 + xf.indent * 3 * mdw + 'px';
        t.appendChild(inner);
        textLayer.appendChild(t);
      }
    }

    var borderLayer = document.createDocumentFragment();
    Object.keys(segments).forEach(function (key) {
      var parts = key.split('|');
      var dir = parts[0];
      var at = Number(parts[1]);
      var w = Number(parts[2]);
      var css = w + 'px ' + parts[3] + ' ' + parts[4];
      var list = segments[key].sort(function (a, b) { return a[0] - b[0]; });
      var cur = null;
      function flush() {
        if (!cur) return;
        var from = Math.max(0, cur[0]);
        var d = dir === 'h' ? box('xl-line-h', from, at, cur[1] - from, 0) : box('xl-line-v', at, from, 0, cur[1] - from);
        d.style[dir === 'h' ? 'borderTop' : 'borderLeft'] = css;
        borderLayer.appendChild(d);
      }
      list.forEach(function (s) {
        if (cur && s[0] <= cur[1]) cur[1] = Math.max(cur[1], s[1]);
        else {
          flush();
          cur = [s[0], s[1]];
        }
      });
      flush();
    });

    // Drawings: an offset past its cell runs on into the next ones.
    function point(p) {
      var col = Math.min(p.col, lastCol + 1);
      var row = Math.min(p.row, lastRow + 1);
      var x = p.colOff / EMU_PER_PX;
      var y = p.rowOff / EMU_PER_PX;
      while (col <= lastCol && x > colW[col]) { x -= colW[col]; col++; }
      while (row <= lastRow && y > rowH[row]) { y -= rowH[row]; row++; }
      return { x: colLeft[col] + x, y: rowTop[row] + y };
    }
    var drawLayer = document.createDocumentFragment();
    items.forEach(function (item) {
      var at = item.from ? point(item.from) : item.pos;
      var w;
      var h;
      if (item.to) {
        var end = point(item.to);
        w = end.x - at.x;
        h = end.y - at.y;
      } else {
        w = item.size.w;
        h = item.size.h;
      }
      if (!(w > 0 && h > 0)) return;
      // A group's shapes each take their share of the anchor's rectangle.
      item.parts.forEach(function (part) {
        var node;
        if (part.src) {
          node = document.createElement('img');
          node.src = part.src;
          node.alt = '';
          node.className = 'xl-pic';
        } else {
          node = P.el('div', 'xl-shape');
          var span = P.el('div', null, part.text);
          span.style.fontSize = part.fontSize + 'pt';
          if (part.bold) span.style.fontWeight = '700';
          if (part.fontName) span.style.fontFamily = '"' + part.fontName + '", sans-serif';
          span.style.color = part.color;
          node.appendChild(span);
          node.style.alignItems = part.align;
          node.style.justifyContent = part.valign;
          if (part.fill) node.style.background = part.fill;
          if (part.line) node.style.outline = '1px solid ' + part.line;
        }
        node.style.left = at.x + part.box.x * w + 'px';
        node.style.top = at.y + part.box.y * h + 'px';
        node.style.width = part.box.w * w + 'px';
        node.style.height = part.box.h * h + 'px';
        drawLayer.appendChild(node);
      });
    });

    grid.appendChild(gridLayer);
    grid.appendChild(fillLayer);
    grid.appendChild(textLayer);
    grid.appendChild(borderLayer);
    grid.appendChild(drawLayer);
    var fitCol = printLastCol != null ? Math.min(printLastCol, lastCol) : lastCol;
    var area = info.printArea || { r0: 0, c0: 0, r1: lastRow, c1: lastCol };
    return {
      node: root,
      width: rowHeadW + colLeft[fitCol + 1],
      note: note,
      print: {
        grid: grid,
        colLeft: colLeft,
        rowTop: rowTop,
        page: sheet.page,
        area: {
          r0: Math.min(area.r0, lastRow),
          c0: Math.min(area.c0, lastCol),
          r1: Math.min(area.r1, lastRow),
          c1: Math.min(area.c1, lastCol)
        }
      }
    };
  }

  // ---- printing ------------------------------------------------------------------
  // Paper sizes in mm by Excel's paperSize code; anything unknown prints on A4.
  var PAPER = { 1: [215.9, 279.4], 5: [215.9, 355.6], 8: [297, 420], 9: [210, 297], 11: [148, 210] };
  var PX_PER_MM = 96 / 25.4;

  // The drawn sheet cut into printed pages the way Excel would: only the print
  // area, a new page at every manual page break (an OT form per page), more
  // breaks wherever a page fills up, scaled to fit the paper width, no row or
  // column headers, and gridlines only when the file asks for them.
  // Returns { node, pageCss } — the pages, and the @page rule they are sized for.
  function printView(print) {
    var page = print.page || { breaks: [], paper: 9, landscape: false, scale: 100, fitToPage: false, fitWidth: 1, fitHeight: 1, margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 } };
    var area = print.area;
    var colLeft = print.colLeft;
    var rowTop = print.rowTop;
    var paper = PAPER[page.paper] || PAPER[9];
    var paperW = page.landscape ? paper[1] : paper[0];
    var paperH = page.landscape ? paper[0] : paper[1];
    var m = page.margins;
    // Room on the paper, in CSS px; a couple of px spare so rounding never
    // pushes a page's last row onto a page of its own.
    var roomW = paperW * PX_PER_MM - (m.left + m.right) * 96 - 2;
    var roomH = paperH * PX_PER_MM - (m.top + m.bottom) * 96 - 2;

    var x0 = colLeft[area.c0];
    var areaW = colLeft[area.c1 + 1] - x0;
    var areaH = rowTop[area.r1 + 1] - rowTop[area.r0];
    var zoom = page.fitToPage ? 1 : page.scale / 100;
    if (page.fitToPage && page.fitWidth > 0) zoom = Math.min(zoom, roomW * page.fitWidth / areaW);
    if (page.fitToPage && page.fitHeight > 0) zoom = Math.min(zoom, roomH * page.fitHeight / areaH);
    // Columns are never split across pages here: a sheet wider than the paper
    // is shrunk to fit it instead.
    zoom = Math.max(0.1, Math.min(zoom, roomW / areaW));

    // Row ranges, one per page.
    var manual = {};
    page.breaks.forEach(function (row) { manual[row] = true; });
    var pages = [];
    var start = area.r0;
    for (var r = area.r0; r <= area.r1; r++) {
      var tall = (rowTop[r + 1] - rowTop[start]) * zoom > roomH;
      if (r > start && (manual[r] || tall)) {
        pages.push([start, r - 1]);
        start = r;
      }
    }
    pages.push([start, area.r1]);

    var parts = Array.prototype.slice.call(print.grid.children);
    var root = P.el('div', 'print-root');
    pages.forEach(function (range) {
      var y0 = rowTop[range[0]];
      var y1 = rowTop[range[1] + 1];
      var sheet = P.el('div', 'print-page');
      sheet.style.zoom = zoom;
      if (page.centered) sheet.style.margin = '0 auto';
      sheet.style.width = areaW + 'px';
      sheet.style.height = y1 - y0 + 'px';
      var grid = print.grid.cloneNode(false);
      grid.style.left = -x0 + 'px';
      grid.style.top = -y0 + 'px';
      // Only what reaches into this page's rows: every page holding a copy of
      // the whole sheet would make a 31-page month very heavy to print.
      parts.forEach(function (part) {
        if (!page.gridlines && part.classList.contains('xl-gl')) return;
        var top = parseFloat(part.style.top) || 0;
        var bottom = top + (parseFloat(part.style.height) || part.offsetHeight || 0);
        if (bottom > y0 && top < y1) grid.appendChild(part.cloneNode(true));
      });
      sheet.appendChild(grid);
      root.appendChild(sheet);
    });

    var css = '@page { size: ' + paperW + 'mm ' + paperH + 'mm; margin: ' +
      m.top + 'in ' + m.right + 'in ' + m.bottom + 'in ' + m.left + 'in; }';
    return { node: root, pageCss: css };
  }

  window.RapSheet = { open: open, printView: printView };
})();
