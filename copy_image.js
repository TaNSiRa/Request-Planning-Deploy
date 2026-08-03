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
//      which Word / Outlook / Teams / Excel paste as a picture.
//   3. Neither worked → download the .png so the user can attach it manually.
//
// Resolves to "clipboard" | "html" | "download" | "failed" so the Dart side can
// tell the user what actually happened.
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

    try {
      if (await copyViaExecCommand(blob)) return "html";
    } catch (err) {
      console.warn("copy_image: execCommand copy failed, falling back", err);
    }

    return download(blob, fileName) ? "download" : "failed";
  };
})();
