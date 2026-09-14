// Draws one PDF. See pdf.html for why it is shaped this way.
import * as pdfjsLib from './vendor/pdfjs/pdf.min.js';

const P = window.RapPreview;
const app = document.getElementById('app');
const base = new URL('./vendor/pdfjs/', import.meta.url).href;

// The files are .js rather than pdf.js's own .mjs on purpose: a module only
// needs a JavaScript MIME type, and not every server maps .mjs to one.
pdfjsLib.GlobalWorkerOptions.workerSrc = base + 'pdf.worker.min.js';

P.showStatus(app, 'loading', 'Opening PDF…');

P.onLoad(async (message) => {
  let locked = false;
  try {
    const task = pdfjsLib.getDocument({
      data: P.bytesFrom(message.dataUrl),
      cMapUrl: base + 'cmaps/',
      cMapPacked: true,
      standardFontDataUrl: base + 'standard_fonts/',
      wasmUrl: base + 'wasm/',
      isEvalSupported: false,
      enableXfa: false
    });
    // No password prompt: an encrypted attachment is opened in a real reader.
    task.onPassword = () => {
      locked = true;
      task.destroy();
    };
    await show(await task.promise);
  } catch (err) {
    P.fail(app, err, locked ? 'This PDF is password-protected' : "This PDF can't be previewed");
  }
});

async function show(pdf) {
  const pages = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const viewport = page.getViewport({ scale: 1 });
    pages.push({ n, page, width: viewport.width, height: viewport.height, div: P.el('div', 'pdf-page') });
  }

  const scroller = P.el('div', 'scroller');
  const list = P.el('div', 'pages');
  pages.forEach((p, i) => {
    p.div.dataset.index = String(i);
    list.appendChild(p.div);
  });
  scroller.appendChild(list);
  app.appendChild(scroller);
  P.clearStatus(app);

  const widest = Math.max(...pages.map((p) => p.width));
  let manual = null;
  const effective = () => manual ?? Math.max(0.25, Math.min(2, (scroller.clientWidth - 40) / widest));

  const pageLabel = P.el('span', 'pages-label');
  const bar = P.zoomBar(app, scroller, {
    extra: pageLabel,
    get: effective,
    set: (scale) => {
      const ratio = scroller.scrollTop / Math.max(1, scroller.scrollHeight);
      manual = scale;
      layout();
      scroller.scrollTop = ratio * scroller.scrollHeight;
    }
  });

  // Only pages near the viewport hold a canvas; the rest are sized boxes, so
  // a 300-page manual costs what the few visible pages cost.
  const near = new Set();
  const observer = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      const p = pages[Number(entry.target.dataset.index)];
      if (entry.isIntersecting) {
        near.add(p);
        draw(p);
      } else {
        near.delete(p);
        release(p);
      }
    }
  }, { root: scroller, rootMargin: '800px 0px' });
  pages.forEach((p) => observer.observe(p.div));

  let redraw = 0;
  function layout() {
    const scale = effective();
    for (const p of pages) {
      p.div.style.width = Math.round(p.width * scale) + 'px';
      p.div.style.height = Math.round(p.height * scale) + 'px';
    }
    bar.update();
    // Boxes resize at once (the old canvas stretches with them); the sharp
    // redraw waits until a zoom gesture settles.
    clearTimeout(redraw);
    redraw = setTimeout(() => near.forEach(draw), 140);
  }

  function draw(p) {
    const scale = effective();
    if (p.canvas && p.scale === scale) return;
    // Already drawing at this scale: let it finish. Cancelling here restarted
    // the first page every time the observer and the layout both asked for it,
    // which held the first paint back by seconds.
    if (p.task && p.taskScale === scale) return;
    if (p.task) p.task.cancel();
    const ratio = Math.min(window.devicePixelRatio || 1, 2);
    const viewport = p.page.getViewport({ scale: scale * ratio });
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const task = p.page.render({ canvas, viewport });
    p.task = task;
    p.taskScale = scale;
    task.promise.then(() => {
      if (p.task !== task) return;
      p.task = null;
      p.div.textContent = '';
      p.div.appendChild(canvas);
      p.canvas = canvas;
      p.scale = scale;
    }, (err) => {
      if (err && err.name !== 'RenderingCancelledException') console.error(err);
    });
  }

  function release(p) {
    if (p.task) p.task.cancel();
    p.task = null;
    p.canvas = null;
    p.scale = 0;
    p.div.textContent = '';
  }

  let pending = false;
  function updatePageLabel() {
    pending = false;
    const middle = scroller.getBoundingClientRect().top + scroller.clientHeight / 2;
    let current = 1;
    for (const p of pages) {
      if (p.div.getBoundingClientRect().top <= middle) current = p.n;
      else break;
    }
    pageLabel.textContent = current + ' / ' + pages.length;
  }
  scroller.addEventListener('scroll', () => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(updatePageLabel);
  }, { passive: true });

  new ResizeObserver(() => { if (manual == null) layout(); }).observe(scroller);
  layout();
  updatePageLabel();
}
