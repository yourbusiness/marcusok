/** Trigger a browser download from a Blob. No-op in Node (document undefined). */
export function triggerDownload(blob: Blob, filename: string): void {
  if (typeof document === "undefined") return;
  const url = URL.createObjectURL(blob);
  // Release the objectURL on every path: a sandboxed DOM can throw between
  // createObjectURL and the delayed revoke (createElement/appendChild/click),
  // and each leaked URL keeps its whole Blob alive for the page's lifetime.
  // The revoke stays delayed: the browser reads the URL asynchronously after
  // the synchronous click, so revoking immediately can cancel the download.
  try {
    const a = document.createElement("a");
    a.href = url;
    a.download = filename.toLowerCase().endsWith(".xlsx")
      ? filename
      : `${filename}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}

/**
 * Coerce a Uint8Array into a BlobPart. TS 5.7+ widened Uint8Array to a generic
 * over ArrayBufferLike (which includes SharedArrayBuffer), making it incompatible
 * with BlobPart's ArrayBufferView<ArrayBuffer>. At runtime any Uint8Array is a
 * valid BlobPart; this cast is the documented workaround.
 */
export function toBlobPart(bytes: Uint8Array): BlobPart {
  return bytes as unknown as BlobPart;
}
