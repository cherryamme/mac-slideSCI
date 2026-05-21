import type { ExportProgress } from "../shared/types";

const PDF_SLICE_SIZE = 4 * 1024 * 1024;

function ensureOfficeDocument(): Office.Document {
  if (typeof Office === "undefined" || !Office.context?.document) {
    throw new Error("当前页面不在 PowerPoint Office.js 运行环境中。请在 PowerPoint for Mac 中打开插件。");
  }

  return Office.context.document;
}

function getOfficeErrorMessage(error: Office.Error | undefined, fallback: string): string {
  if (!error) {
    return fallback;
  }

  return `${fallback}${error.message ? `：${error.message}` : ""}`;
}

function sliceDataToUint8Array(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) {
    return data;
  }

  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }

  if (Array.isArray(data)) {
    return Uint8Array.from(data as number[]);
  }

  if (typeof data === "string") {
    return new TextEncoder().encode(data);
  }

  throw new Error("PowerPoint 返回了无法识别的 PDF slice 数据。");
}

function getFileAsync(document: Office.Document): Promise<Office.File> {
  return new Promise((resolve, reject) => {
    document.getFileAsync(Office.FileType.Pdf, { sliceSize: PDF_SLICE_SIZE }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
        return;
      }

      reject(new Error(getOfficeErrorMessage(result.error, "无法从 PowerPoint 获取 PDF")));
    });
  });
}

function getSliceAsync(file: Office.File, index: number): Promise<Office.Slice> {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
        return;
      }

      reject(new Error(getOfficeErrorMessage(result.error, `无法读取 PDF 数据片段 ${index + 1}`)));
    });
  });
}

function closeFileAsync(file: Office.File): Promise<void> {
  return new Promise((resolve) => {
    file.closeAsync(() => resolve());
  });
}

export async function getPresentationPdfBytes(onProgress?: (progress: ExportProgress) => void): Promise<Uint8Array> {
  const document = ensureOfficeDocument();
  onProgress?.({ phase: "pdf", current: 0, total: 1, message: "正在从 PowerPoint 获取 PDF..." });

  const file = await getFileAsync(document);
  const chunks: Uint8Array[] = [];

  try {
    if (file.size <= 0 || file.sliceCount <= 0) {
      throw new Error("PowerPoint 返回的 PDF 为空。");
    }

    for (let index = 0; index < file.sliceCount; index += 1) {
      onProgress?.({
        phase: "pdf",
        current: index + 1,
        total: file.sliceCount,
        message: `正在读取 PDF 数据 ${index + 1} / ${file.sliceCount}...`,
      });
      const slice = await getSliceAsync(file, index);
      chunks.push(sliceDataToUint8Array(slice.data));
    }
  } finally {
    await closeFileAsync(file);
  }

  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  if (totalLength === 0) {
    throw new Error("PowerPoint 返回的 PDF 没有可用数据。");
  }

  const merged = new Uint8Array(totalLength);
  let offset = 0;
  chunks.forEach((chunk) => {
    merged.set(chunk, offset);
    offset += chunk.byteLength;
  });

  return merged;
}

export function normalizeSelectedSlidePageIndices(slides: Office.Slide[] | undefined): number[] {
  if (!slides?.length) {
    return [];
  }

  return [...new Set(slides.map((slide) => slide.index - 1).filter((index) => Number.isInteger(index) && index >= 0))].sort((a, b) => a - b);
}

export function getSelectedSlidePageIndices(): Promise<number[]> {
  const document = ensureOfficeDocument();

  return new Promise((resolve, reject) => {
    document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(getOfficeErrorMessage(result.error, "无法读取选中的幻灯片")));
        return;
      }

      const slideRange = result.value as Office.SlideRange | undefined;
      const indices = normalizeSelectedSlidePageIndices(slideRange?.slides);
      if (indices.length === 0) {
        reject(new Error("没有读取到可导出的选中幻灯片页码，请改用导出全部幻灯片。"));
        return;
      }

      resolve(indices);
    });
  });
}
