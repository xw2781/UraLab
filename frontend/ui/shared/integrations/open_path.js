function normalizeOpenPathResult(result) {
  if (result === true) return { ok: true, error: "" };
  if (!result || typeof result !== "object") {
    return { ok: false, error: "Open path failed." };
  }
  return {
    ok: result.ok === true,
    error: String(result.error || ""),
  };
}

/** Opens a path through the desktop host, using the shell message bridge from iframe pages. */
export function openPathThroughDesktopHost(
  targetPath,
  { readOnly = false, preferredApp = "" } = {},
  windowRef = globalThis.window,
) {
  const path = String(targetPath || "").trim();
  if (!path) return Promise.resolve({ ok: false, error: "Empty path." });

  const hostApi = windowRef?.ADAHost || null;
  if (hostApi && typeof hostApi.openPath === "function") {
    return Promise.resolve(hostApi.openPath({ path, readOnly: !!readOnly, preferredApp }))
      .then(normalizeOpenPathResult)
      .catch((error) => ({ ok: false, error: String(error?.message || error) }));
  }

  return new Promise((resolve) => {
    const parentWindow = windowRef?.parent;
    if (!parentWindow || parentWindow === windowRef) {
      resolve({ ok: false, error: "Open path requires desktop app." });
      return;
    }

    const requestId = `open-path-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    let settled = false;
    let timeoutId = null;
    const finish = (result) => {
      if (settled) return;
      settled = true;
      if (timeoutId !== null) windowRef.clearTimeout(timeoutId);
      windowRef.removeEventListener("message", handleMessage);
      resolve(normalizeOpenPathResult(result));
    };
    const handleMessage = (event) => {
      const message = event?.data;
      if (!message || message.type !== "arcrho:open-path-result") return;
      if (String(message.requestId || "") !== requestId) return;
      finish(message);
    };

    windowRef.addEventListener("message", handleMessage);
    timeoutId = windowRef.setTimeout(() => {
      finish({ ok: false, error: "Open path timed out." });
    }, 5000);
    try {
      parentWindow.postMessage({
        type: "arcrho:open-path",
        requestId,
        path,
        readOnly: !!readOnly,
        ...(preferredApp ? { preferredApp } : {}),
      }, "*");
    } catch {
      finish({ ok: false, error: "Open path requires desktop app." });
    }
  });
}

/** Returns the folder a path sits in, or "" when the path names no folder. */
function parentFolderOf(targetPath) {
  const path = String(targetPath || "").trim().replace(/[\\/]+$/u, "");
  const cut = Math.max(path.lastIndexOf("\\"), path.lastIndexOf("/"));
  if (cut <= 0) return "";
  const folder = path.slice(0, cut);
  // A bare drive letter is not a folder path; Windows needs the trailing slash.
  return /^[A-Za-z]:$/u.test(folder) ? `${folder}\\` : folder;
}

/**
 * Shows a file in the desktop file manager, selecting it where the host can.
 * Pages inside an iframe reach the host only through the open-path bridge, so
 * they fall back to opening the containing folder itself.
 */
export function revealPathThroughDesktopHost(targetPath, windowRef = globalThis.window) {
  const path = String(targetPath || "").trim();
  if (!path) return Promise.resolve({ ok: false, error: "Empty path." });

  const hostApi = windowRef?.ADAHost || null;
  if (hostApi && typeof hostApi.showItemInFolder === "function") {
    return Promise.resolve(hostApi.showItemInFolder({ path }))
      .then(normalizeOpenPathResult)
      .catch((error) => ({ ok: false, error: String(error?.message || error) }));
  }

  const folder = parentFolderOf(path);
  if (!folder) return Promise.resolve({ ok: false, error: `No folder for path: ${path}` });
  return openPathThroughDesktopHost(folder, {}, windowRef);
}

/**
 * Copies text to the clipboard, falling back to a hidden field where the
 * async clipboard is unavailable or refused.
 */
export function copyTextToClipboard(text, windowRef = globalThis.window) {
  const value = String(text ?? "");
  if (!value) return Promise.resolve({ ok: false, error: "Nothing to copy." });

  const clipboard = windowRef?.navigator?.clipboard;
  if (clipboard && typeof clipboard.writeText === "function") {
    return Promise.resolve(clipboard.writeText(value))
      .then(() => ({ ok: true, error: "" }))
      .catch(() => copyThroughHiddenField(value, windowRef));
  }
  return Promise.resolve(copyThroughHiddenField(value, windowRef));
}

function copyThroughHiddenField(value, windowRef) {
  const documentRef = windowRef?.document;
  const host = documentRef?.body || documentRef?.documentElement;
  if (!documentRef || !host || typeof documentRef.execCommand !== "function") {
    return { ok: false, error: "Copying is unavailable here." };
  }
  const field = documentRef.createElement("textarea");
  field.value = value;
  field.setAttribute("readonly", "");
  field.style.position = "fixed";
  field.style.left = "-9999px";
  field.style.top = "0";
  host.appendChild(field);
  try {
    field.select?.();
    const copied = documentRef.execCommand("copy");
    if (copied === false) return { ok: false, error: "The clipboard refused the copy." };
    return { ok: true, error: "" };
  } catch (error) {
    return { ok: false, error: String(error?.message || error) };
  } finally {
    field.remove?.();
  }
}
