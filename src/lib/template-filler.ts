import JSZip from "jszip";

export function downloadBuffer(buf: ArrayBuffer, baseName: string) {
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = baseName.endsWith(".xlsx") ? baseName : `${baseName}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2500);
}

/**
 * Downloads multiple files sequentially with a stagger delay to prevent
 * browser download drop or popup blocking when downloading 20-30+ files.
 */
export async function downloadMultipleBuffers(
  files: { buf: ArrayBuffer; fileName: string }[],
  delayMs = 300
): Promise<void> {
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    downloadBuffer(f.buf, f.fileName);
    if (i < files.length - 1) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }
}

/**
 * Packs all buffer files into a single ZIP archive and triggers a single 1-click download.
 * Ensures 100% data integrity even when downloading 25+ files without any missing data.
 */
export async function downloadZipBundle(
  files: { buf: ArrayBuffer; fileName: string }[],
  zipName: string
): Promise<void> {
  const zip = new JSZip();
  for (const f of files) {
    const name = f.fileName.endsWith(".xlsx") ? f.fileName : `${f.fileName}.xlsx`;
    zip.file(name, f.buf);
  }
  const content = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  const url = URL.createObjectURL(content);
  const a = document.createElement("a");
  a.href = url;
  a.download = zipName.endsWith(".zip") ? zipName : `${zipName}.zip`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

