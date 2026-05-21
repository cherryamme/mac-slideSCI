import type { ExportProgress, RenderedPngPage } from "../../shared/types";
import { assertCanvasSize, dpiToScale, normalizeDpi, normalizePageIndices } from "./pngExportMath";

type PdfJsModule = typeof import("pdfjs-dist");

let pdfJsModule: PdfJsModule | null = null;

async function loadPdfJs(): Promise<PdfJsModule> {
  if (!pdfJsModule) {
    pdfJsModule = await import("pdfjs-dist");
    pdfJsModule.GlobalWorkerOptions.workerSrc = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).toString();
  }

  return pdfJsModule;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw new Error("已取消导出。");
  }
}

function canvasToPngBlob(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("无法从渲染结果生成 PNG。"));
        return;
      }

      resolve(blob);
    }, "image/png");
  });
}

export async function renderPdfPagesToPngs(options: {
  pdfBytes: Uint8Array;
  dpi: number;
  pageIndices?: number[];
  onProgress?: (progress: ExportProgress) => void;
  signal?: AbortSignal;
}): Promise<RenderedPngPage[]> {
  throwIfAborted(options.signal);

  const pdfjsLib = await loadPdfJs();
  const dpi = normalizeDpi(options.dpi);
  const pdf = await pdfjsLib.getDocument({ data: options.pdfBytes.slice().buffer }).promise;

  try {
    const targetPageIndices = normalizePageIndices(options.pageIndices, pdf.numPages);
    if (targetPageIndices.length === 0) {
      throw new Error("没有可导出的页面。");
    }

    const renderedPages: RenderedPngPage[] = [];
    const scale = dpiToScale(dpi);

    for (let outputIndex = 0; outputIndex < targetPageIndices.length; outputIndex += 1) {
      throwIfAborted(options.signal);
      const pageIndex = targetPageIndices[outputIndex];
      const pageNumber = pageIndex + 1;
      options.onProgress?.({
        phase: "render",
        current: outputIndex + 1,
        total: targetPageIndices.length,
        message: `正在渲染第 ${pageNumber} 页 (${outputIndex + 1} / ${targetPageIndices.length})...`,
      });

      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale });
      assertCanvasSize(viewport.width, viewport.height);

      const canvas = document.createElement("canvas");
      const context = canvas.getContext("2d");
      if (!context) {
        throw new Error("当前环境无法创建 canvas 渲染上下文。");
      }

      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);

      const renderTask = page.render({ canvas, canvasContext: context, viewport });
      const abortHandler = () => renderTask.cancel();
      options.signal?.addEventListener("abort", abortHandler, { once: true });

      try {
        await renderTask.promise;
        throwIfAborted(options.signal);
        const blob = await canvasToPngBlob(canvas);
        renderedPages.push({ pageNumber, blob, width: canvas.width, height: canvas.height });
      } finally {
        options.signal?.removeEventListener("abort", abortHandler);
        page.cleanup();
        canvas.width = 0;
        canvas.height = 0;
      }
    }

    return renderedPages;
  } finally {
    await pdf.cleanup();
    await pdf.destroy();
  }
}
