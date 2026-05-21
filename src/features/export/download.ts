import JSZip from "jszip";
import type { ExportProgress, PngExportResult, RenderedPngPage, RenderedPngSelection } from "../../shared/types";

function padPageNumber(pageNumber: number): string {
  return String(pageNumber).padStart(3, "0");
}

export function getPngFileName(pageNumber: number, dpi: number): string {
  return `mac-slideSCI-slide-${padPageNumber(pageNumber)}-${dpi}dpi.png`;
}

export function getZipFileName(dpi: number): string {
  return `mac-slideSCI-png-export-${dpi}dpi.zip`;
}

export function getSelectedObjectsPngFileName(dpi: number): string {
  return `mac-slideSCI-selection-${dpi}dpi.png`;
}

export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.rel = "noopener";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function downloadRenderedPngSelection(selection: RenderedPngSelection, onProgress?: (progress: ExportProgress) => void): PngExportResult {
  onProgress?.({ phase: "download", current: 1, total: 1, message: "正在下载选中对象 PNG..." });
  const fileName = getSelectedObjectsPngFileName(selection.dpi);
  downloadBlob(selection.blob, fileName);
  return {
    count: 1,
    fileName,
    message: `已导出 ${selection.selectedShapeCount} 个选中对象的 ${selection.dpi} DPI PNG（${selection.width} × ${selection.height}px）。`,
  };
}

export async function downloadRenderedPngPages(
  pages: RenderedPngPage[],
  dpi: number,
  onProgress?: (progress: ExportProgress) => void,
): Promise<PngExportResult> {
  if (pages.length === 0) {
    throw new Error("没有可下载的 PNG 页面。");
  }

  onProgress?.({ phase: "download", current: 0, total: pages.length, message: "正在准备下载..." });

  if (pages.length === 1) {
    const fileName = getPngFileName(pages[0].pageNumber, dpi);
    downloadBlob(pages[0].blob, fileName);
    return { count: 1, fileName, message: `已导出 1 张 ${dpi} DPI PNG。` };
  }

  const zip = new JSZip();
  pages.forEach((page) => {
    zip.file(getPngFileName(page.pageNumber, dpi), page.blob);
  });

  const fileName = getZipFileName(dpi);
  const blob = await zip.generateAsync({ type: "blob", streamFiles: true }, (metadata) => {
    onProgress?.({
      phase: "zip",
      current: Math.round(metadata.percent),
      total: 100,
      message: `正在打包 ZIP：${Math.round(metadata.percent)}%...`,
    });
  });

  downloadBlob(blob, fileName);
  return { count: pages.length, fileName, message: `已导出 ${pages.length} 张 ${dpi} DPI PNG。` };
}
