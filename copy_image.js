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
//      which Word / Outlook / Teams / Excel paste as a picture — but anything
//      that wants a real bitmap or a file (LINE, Claude, Discord, Facebook)
//      pastes NOTHING from it.
//   3. Neither worked → download the .png so the user can attach it manually.
//
// Because of that gap, the picture ALSO goes on screen every time, in a plain
// DOM overlay — including after a successful path 1, so the button does not
// behave one way on localhost (a secure context) and another on the server.
// The overlay offers the two routes a browser allows even from an insecure
// origin:
//   * right-click the <img> -> "Copy image" — the browser's own copy, which
//     puts a TRUE bitmap on the clipboard; that one pastes into LINE and Claude.
//   * drag the <img> straight into the chat window — the overlay tags the drag
//     with a DownloadURL entry, so what lands in LINE / Claude / Explorer is a
//     real .png FILE, not a link to a blob: URL they cannot resolve.
// The overlay is DOM (not Flutter) on purpose: the Flutter canvas has no
// right-click "Copy image" and cannot start a file drag.
//
// The permanent fix for all of this is serving the app over https, which makes
// path 1 work and turns the button back into one click everywhere.
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
  // to the clipboard, and the two ways to get it into LINE / Claude. Only ever
  // one at a time. Returns true when it was put on screen.
  //
  // `copied` says what the clipboard actually holds — "bitmap" (the real thing),
  // "html" (the plain-http fallback) or "" (nothing). It changes only the line
  // of text at the top: the picture, the right-click route and the drag route
  // are offered every time, because a successful clipboard copy still leaves
  // people wanting to drag the file somewhere or keep it.
  function showPicturePanel(blob, fileName, copied) {
    try {
      const existing = document.getElementById("rap-copy-image-overlay");
      if (existing) existing.remove();

      const name = fileName || "image.png";
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
        "max-width:min(880px,92vw);max-height:88vh;overflow:auto;padding:20px 22px 18px;";

      const title = document.createElement("div");
      title.textContent = "Paste this picture into LINE, Claude or anywhere else";
      title.style.cssText =
        "font-size:17px;font-weight:800;color:#0f172a;margin-bottom:6px;";

      const hint = document.createElement("div");
      hint.innerHTML =
        copied === "bitmap"
          ? "It is on the clipboard as a real image — press <b>Ctrl+V</b> in LINE, Claude, Teams, " +
            "Word, anywhere. The copy below is here if you would rather drag it or keep it:"
          : (copied === "html"
              ? "It is already on the clipboard for <b>Outlook, Word and Teams</b> — just press Ctrl+V there. "
              : "") +
            "Apps that want a real image (<b>LINE, Claude</b>) need one of these, because this browser " +
            "hands a picture to the clipboard only on a secure (https) address:";
      hint.style.cssText =
        "font-size:13px;line-height:1.55;color:#475569;margin-bottom:10px;max-width:790px;";

      const steps = document.createElement("div");
      steps.innerHTML =
        "<div style=\"margin-bottom:4px\"><b>1&nbsp;&nbsp;Right-click the picture &rarr; &ldquo;Copy image&rdquo;</b>, then paste in the chat.</div>" +
        "<div><b>2&nbsp;&nbsp;Or drag the picture</b> straight into the chat window — it arrives as a .png file.</div>";
      steps.style.cssText =
        "font-size:13px;line-height:1.6;color:#0f172a;background:#f1f5f9;border-radius:10px;" +
        "padding:10px 12px;margin-bottom:14px;max-width:790px;";

      const frame = document.createElement("div");
      frame.style.cssText =
        "border:1px solid #e2e8f0;border-radius:12px;padding:8px;background:#f8fafc;" +
        "display:flex;justify-content:center;";
      const img = document.createElement("img");
      img.src = url;
      img.alt = name;
      img.title = "Right-click for “Copy image”, or drag me into the chat";
      img.draggable = true;
      img.style.cssText =
        "max-width:100%;height:auto;display:block;border-radius:6px;cursor:grab;";
      // Chrome/Edge read this entry when a drag leaves the window and write the
      // real file at the other end — dropping into LINE or Claude then attaches
      // a .png instead of a link back to a blob: URL they cannot resolve.
      img.addEventListener("dragstart", function (event) {
        try {
          event.dataTransfer.effectAllowed = "copy";
          event.dataTransfer.setData("DownloadURL", "image/png:" + name + ":" + url);
          event.dataTransfer.setData("text/uri-list", url);
        } catch (err) {
          console.warn("copy_image: could not tag the drag", err);
        }
      });
      // The Flutter view disables the browser context menu on its own host
      // element; keep this one out of anything listening further up the tree so
      // "Copy image" is definitely offered here.
      img.addEventListener("contextmenu", function (event) { event.stopPropagation(); });
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
      downloadBtn.addEventListener("click", function () { download(blob, name); });

      const closeBtn = document.createElement("button");
      closeBtn.type = "button";
      closeBtn.textContent = "Close";
      closeBtn.style.cssText = buttonStyle + "background:#1e2b6f;color:#fff;border:none;";

      function close() {
        window.removeEventListener("keydown", onKey, true);
        // Revoked late, not now: a drag that has just left the window still
        // fetches the file from this URL after the overlay is gone.
        setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
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
      panel.appendChild(steps);
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
        // The picture goes up even though the clipboard already has the real
        // thing: the panel is also how people drag the file into an app or save
        // it, and a button that behaves one way on localhost (secure) and
        // another on the server (not) is a button nobody can learn.
        showPicturePanel(blob, fileName, "bitmap");
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
    if (showPicturePanel(blob, fileName, copiedAsHtml ? "html" : "")) {
      return copiedAsHtml ? "html" : "manual";
    }
    if (copiedAsHtml) return "html";

    return download(blob, fileName) ? "download" : "failed";
  };
})();
