import { isPowerPointApiSupported } from "./requirements";
import { ensurePowerPointApi, runPowerPoint } from "./powerpoint";
import { assertCanvasSize, getUnionBounds, normalizeDpi, normalizePaddingPoints, pointsToPixels } from "../features/export/pngExportMath";
import type { ExportProgress, RenderedPngSelection, ShapeBox } from "../shared/types";

const SHAPE_LOAD = "items/id,items/name,items/type,items/left,items/top,items/width,items/height";
const DEFAULT_SLIDE_WIDTH_POINTS = 960;
const DEFAULT_SLIDE_HEIGHT_POINTS = 540;

interface Base64SelectionImage {
  base64: string;
  width: number;
  height: number;
  dpi: number;
  selectedShapeCount: number;
  crop?: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
}

function toShapeBox(shape: PowerPoint.Shape): ShapeBox {
  return {
    id: shape.id,
    name: shape.name ?? "",
    type: String(shape.type ?? "unknown"),
    left: Number(shape.left ?? 0),
    top: Number(shape.top ?? 0),
    width: Number(shape.width ?? 0),
    height: Number(shape.height ?? 0),
  };
}

function stripDataUrlPrefix(base64: string): string {
  const marker = "base64,";
  const markerIndex = base64.indexOf(marker);
  return markerIndex >= 0 ? base64.slice(markerIndex + marker.length) : base64;
}

function base64ToBlob(base64: string, type = "image/png"): Blob {
  const binary = window.atob(stripDataUrlPrefix(base64));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return new Blob([bytes], { type });
}

function loadImage(base64: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("无法加载 PowerPoint 返回的 PNG 图像。"));
    image.src = base64.startsWith("data:") ? base64 : `data:image/png;base64,${base64}`;
  });
}

function canvasToBlob(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("无法生成选中对象 PNG。"));
        return;
      }

      resolve(blob);
    }, "image/png");
  });
}

async function cropBase64Image(imageData: Base64SelectionImage): Promise<Blob> {
  if (!imageData.crop) {
    return base64ToBlob(imageData.base64);
  }

  const image = await loadImage(imageData.base64);
  const canvas = document.createElement("canvas");
  canvas.width = imageData.crop.width;
  canvas.height = imageData.crop.height;
  assertCanvasSize(canvas.width, canvas.height);

  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("当前环境无法创建 canvas 裁剪上下文。");
  }

  context.drawImage(
    image,
    imageData.crop.x,
    imageData.crop.y,
    imageData.crop.width,
    imageData.crop.height,
    0,
    0,
    imageData.crop.width,
    imageData.crop.height,
  );

  const blob = await canvasToBlob(canvas);
  canvas.width = 0;
  canvas.height = 0;
  return blob;
}

async function getSelectedObjectsBase64(options: {
  dpi: number;
  paddingPoints: number;
  onProgress?: (progress: ExportProgress) => void;
}): Promise<Base64SelectionImage> {
  ensurePowerPointApi("1.6", "读取选中对象");
  const canExportShapeImage = isPowerPointApiSupported("1.10");
  const canExportSlideImage = isPowerPointApiSupported("1.8");
  if (!canExportShapeImage && !canExportSlideImage) {
    throw new Error("当前 PowerPoint 版本无法导出选中对象的高清 PNG。请更新到 Microsoft 365 PowerPoint 16.62 (2208) 或更新版本。");
  }
  const canReadPageSetup = isPowerPointApiSupported("1.10");

  return runPowerPoint(async (context) => {
    options.onProgress?.({ phase: "selection", current: 0, total: 1, message: "正在读取选中对象..." });

    const selectedSlides = context.presentation.getSelectedSlides();
    const slide = selectedSlides.getItemAt(0);
    const shapes = context.presentation.getSelectedShapes();
    const pageSetup = canReadPageSetup ? context.presentation.pageSetup : null;
    shapes.load(SHAPE_LOAD);
    pageSetup?.load("slideWidth,slideHeight");
    await context.sync();

    const boxes = shapes.items.map(toShapeBox);
    if (boxes.length === 0) {
      throw new Error("请先选择要导出的图片或形状。");
    }

    const dpi = normalizeDpi(options.dpi);
    if (boxes.length === 1 && canExportShapeImage) {
      const box = boxes[0];
      const width = pointsToPixels(box.width, dpi);
      const height = pointsToPixels(box.height, dpi);
      assertCanvasSize(width, height);
      options.onProgress?.({ phase: "render", current: 1, total: 1, message: "正在渲染选中对象..." });
      const image = shapes.items[0].getImageAsBase64({ format: PowerPoint.ShapeGetImageFormatType.png, width, height });
      await context.sync();
      if (!image.value) {
        throw new Error("PowerPoint 没有返回选中对象的图像数据。");
      }
      return { base64: image.value, width, height, dpi, selectedShapeCount: 1 };
    }

    if (!canExportSlideImage) {
      throw new Error("当前 PowerPoint 版本不支持选区截图式 PNG 导出。请更新 Microsoft 365 PowerPoint 后重试。");
    }

    const slideWidth = canReadPageSetup && pageSetup ? Number(pageSetup.slideWidth) : DEFAULT_SLIDE_WIDTH_POINTS;
    const slideHeight = canReadPageSetup && pageSetup ? Number(pageSetup.slideHeight) : DEFAULT_SLIDE_HEIGHT_POINTS;
    const union = getUnionBounds(boxes, options.paddingPoints, slideWidth, slideHeight);
    const slideWidthPx = pointsToPixels(slideWidth, dpi);
    const slideHeightPx = pointsToPixels(slideHeight, dpi);
    const crop = {
      x: pointsToPixels(union.left, dpi),
      y: pointsToPixels(union.top, dpi),
      width: pointsToPixels(union.width, dpi),
      height: pointsToPixels(union.height, dpi),
    };
    assertCanvasSize(slideWidthPx, slideHeightPx);
    assertCanvasSize(crop.width, crop.height);

    options.onProgress?.({ phase: "render", current: 1, total: 2, message: "正在渲染当前幻灯片并准备裁剪选区..." });
    const image = slide.getImageAsBase64({ width: slideWidthPx, height: slideHeightPx });
    await context.sync();
    if (!image.value) {
      throw new Error("PowerPoint 没有返回幻灯片图像数据。");
    }

    return {
      base64: image.value,
      width: crop.width,
      height: crop.height,
      dpi,
      selectedShapeCount: boxes.length,
      crop,
    };
  });
}

export async function renderSelectedObjectsPng(options: {
  dpi: number;
  paddingPoints: number;
  signal?: AbortSignal;
  onProgress?: (progress: ExportProgress) => void;
}): Promise<RenderedPngSelection> {
  if (options.signal?.aborted) {
    throw new Error("已取消导出。");
  }

  const imageData = await getSelectedObjectsBase64({
    dpi: normalizeDpi(options.dpi),
    paddingPoints: normalizePaddingPoints(options.paddingPoints),
    onProgress: options.onProgress,
  });

  if (options.signal?.aborted) {
    throw new Error("已取消导出。");
  }

  options.onProgress?.({ phase: imageData.crop ? "crop" : "download", current: 2, total: 2, message: imageData.crop ? "正在裁剪选中对象区域..." : "正在准备下载..." });
  const blob = await cropBase64Image(imageData);

  return {
    blob,
    width: imageData.width,
    height: imageData.height,
    dpi: imageData.dpi,
    selectedShapeCount: imageData.selectedShapeCount,
  };
}
