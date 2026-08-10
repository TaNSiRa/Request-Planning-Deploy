// Puts a PNG on the system clipboard for the "copy as image" buttons (Weekly
// plan in Meeting mode). Lives in JS rather than Dart because every browser
// path here is a try/catch over an API that may simply not exist.
//
// Three paths, tried in order:
//   1. navigator.clipboard.write + ClipboardItem — the real thing, pastes as a
//      bitmap everywhere. Needs a SECURE CONTEXT (https or localhost), so it is
//      unavailable on the production server (plain http://172.23.10.51).
//   2. document.execCommand("copy") over a hidden contenteditable holding an
//      <img>. Still works over plain http; the clipboard gets an HTML flavour,
//      which Word / Outlook / Teams / Excel paste as a picture — but chat apps
//      that only accept a real bitmap (LINE) paste NOTHING from it.
//   3. Neither worked → download the .png so the user can attach it manually.
//
// Because of that LINE gap, whenever path 1 is unavailable we also put the
// picture on screen in a plain DOM overlay: right-clicking a real <img> and
// choosing "Copy image" is the browser's own path to a true bitmap on the
// clipboard, and it works on plain http. The overlay is DOM (not Flutter) on
// purpose — the Flutter canvas has no right-click "Copy image".
//
// Resolves to "clipboard" | "html" | "manual" | "download" | "failed" so the
// Dart side can tell the user what actually happened.
(function () {
  function download(blob, fileName) {
    try {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName || "image.png";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      // Revoke late — Firefox cancels an in-flight download otherwise.
      setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
      return true;
    } catch (err) {
      console.error("copy_image: download failed", err);
      return false;
    }
  }

  function blobToDataUrl(blob) {
    return new Promise(function (resolve, reject) {
      const reader = new FileReader();
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsDataURL(blob);
    });
  }

  // Selects an off-screen <img> and lets the browser's own copy command put it
  // on the clipboard. The node must be visible to layout (not display:none and
  // not visibility:hidden) or the selection copies nothing — hence the
  // off-screen position instead.
  async function copyViaExecCommand(blob) {
    const dataUrl = await blobToDataUrl(blob);
    const holder = document.createElement("div");
    holder.setAttribute("contenteditable", "true");
    holder.style.cssText =
      "position:fixed;left:-100000px;top:0;width:1px;height:1px;overflow:hidden;user-select:text;";
    const img = document.createElement("img");
    img.src = dataUrl;
    holder.appendChild(img);
    document.body.appendChild(holder);
    try {
      // Wait for the bitmap so the browser can offer an image flavour too, not
      // just the HTML markup. Not fatal if it fails.
      if (img.decode) {
        try { await img.decode(); } catch (err) { /* ignore */ }
      }
      const range = document.createRange();
      range.selectNodeContents(holder);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      const ok = document.execCommand("copy");
      selection.removeAllRanges();
      return ok;
    } finally {
      document.body.removeChild(holder);
    }
  }

  // The overlay described at the top: the picture itself, what already made it
  // to the clipboard, and how to get a real bitmap for LINE. Only ever one at a
  // time. Returns true when it was put on screen.
  function showPicturePanel(blob, fileName, copiedAsHtml) {
    try {
      const existing = document.getElementById("rap-copy-image-overlay");
      if (existing) existing.remove();

      const url = URL.createObjectURL(blob);
      const overlay = document.createElement("div");
      overlay.id = "rap-copy-image-overlay";
      overlay.style.cssText =
        "position:fixed;inset:0;z-index:2147483600;background:rgba(15,23,42,.55);" +
        "display:flex;align-items:center;justify-content:center;padding:24px;" +
        "font-family:'IBM Plex Sans','Segoe UI',system-ui,sans-serif;";

      const panel = document.createElement("div");
      panel.style.cssText =
        "background:#fff;border-radius:18px;box-shadow:0 24px 60px rgba(15,23,42,.35);" +
        "max-width:min(860px,92vw);max-height:88vh;overflow:auto;padding:20px 22px 18px;";

      const title = document.createElement("div");
      title.textContent = "Copy this picture into LINE";
      title.style.cssText =
        "font-size:17px;font-weight:800;color:#0f172a;margin-bottom:6px;";

      const hint = document.createElement("div");
      hint.innerHTML =
        (copiedAsHtml
          ? "It is already on the clipboard for <b>Outlook, Word and Teams</b> — just paste there. "
          : "") +
        "LINE only accepts a real image on the clipboard, which this browser can hand over " +
        "only on a secure (https) address. To paste into LINE: <b>right-click the picture " +
        "below → “Copy image”</b>, then paste in the chat.";
      hint.style.cssText =
        "font-size:13px;line-height:1.55;color:#475569;margin-bottom:14px;max-width:760px;";

      const frame = document.createElement("div");
      frame.style.cssText =
        "border:1px solid #e2e8f0;border-radius:12px;padding:8px;background:#f8fafc;" +
        "display:flex;justify-content:center;";
      const img = document.createElement("img");
      img.src = url;
      img.alt = fileName || "image.png";
      img.style.cssText = "max-width:100%;height:auto;display:block;border-radius:6px;";
      frame.appendChild(img);

      const actions = document.createElement("div");
      actions.style.cssText =
        "display:flex;gap:10px;justify-content:flex-end;margin-top:14px;";

      const buttonStyle =
        "border-radius:10px;padding:9px 16px;font-size:13px;font-weight:700;cursor:pointer;";
      const downloadBtn = document.createElement("button");
      downloadBtn.type = "button";
      downloadBtn.textContent = "Download .png";
      downloadBtn.style.cssText =
        buttonStyle + "background:#fff;color:#334155;border:1.4px solid #cbd5e1;";
      downloadBtn.addEventListener("click", function () { download(blob, fileName); });

      const closeBtn = document.createElement("button");
      closeBtn.type = "button";
      closeBtn.textContent = "Close";
      closeBtn.style.cssText = buttonStyle + "background:#1e2b6f;color:#fff;border:none;";

      function close() {
        window.removeEventListener("keydown", onKey, true);
        URL.revokeObjectURL(url);
        overlay.remove();
      }
      function onKey(event) {
        if (event.key === "Escape") {
          event.stopPropagation();
          close();
        }
      }
      closeBtn.addEventListener("click", close);
      overlay.addEventListener("mousedown", function (event) {
        if (event.target === overlay) close();
      });
      // Capture phase: the Flutter view must not act on the Esc that closes this.
      window.addEventListener("keydown", onKey, true);

      actions.appendChild(downloadBtn);
      actions.appendChild(closeBtn);
      panel.appendChild(title);
      panel.appendChild(hint);
      panel.appendChild(frame);
      panel.appendChild(actions);
      overlay.appendChild(panel);
      document.body.appendChild(overlay);
      return true;
    } catch (err) {
      console.error("copy_image: could not show the picture panel", err);
      return false;
    }
  }

  // bytes: Uint8Array of a PNG. Never throws — always resolves to a status.
  window.rapCopyImage = async function (bytes, fileName) {
    let blob;
    try {
      blob = new Blob([bytes], { type: "image/png" });
    } catch (err) {
      console.error("copy_image: could not build the blob", err);
      return "failed";
    }

    if (window.isSecureContext && navigator.clipboard && window.ClipboardItem) {
      try {
        await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
        return "clipboard";
      } catch (err) {
        console.warn("copy_image: clipboard.write refused, falling back", err);
      }
    }

    let copiedAsHtml = false;
    try {
      copiedAsHtml = !!(await copyViaExecCommand(blob));
    } catch (err) {
      console.warn("copy_image: execCommand copy failed, falling back", err);
    }

    // No real bitmap went to the clipboard, so offer the picture itself.
    if (showPicturePanel(blob, fileName, copiedAsHtml)) {
      return copiedAsHtml ? "html" : "manual";
    }
    if (copiedAsHtml) return "html";

    return download(blob, fileName) ? "download" : "failed";
  };
})();
