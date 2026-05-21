import type { ShapeBox } from "../../shared/types";

const DEFAULT_DPI = 300;
const MIN_DPI = 72;
const MAX_DPI = 1200;
const MAX_PAGE_MEGA_PIXELS = 100;
const DEFAULT_PADDING_POINTS = 2;
const MIN_PADDING_POINTS = 0;
const MAX_PADDING_POINTS = 36;

export interface Bounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

export function normalizeDpi(value: number): number {
  if (!Number.isFinite(value)) {
    return DEFAULT_DPI;
  }

  return Math.min(MAX_DPI, Math.max(MIN_DPI, Math.round(value)));
}

export function normalizePaddingPoints(value: number): number {
  if (!Number.isFinite(value)) {
    return DEFAULT_PADDING_POINTS;
  }

  return Math.min(MAX_PADDING_POINTS, Math.max(MIN_PADDING_POINTS, value));
}

export function dpiToScale(dpi: number): number {
  return normalizeDpi(dpi) / 72;
}

export function pointsToPixels(points: number, dpi: number): number {
  return Math.ceil(points * normalizeDpi(dpi) / 72);
}

export function pixelsToPoints(pixels: number, dpi: number): number {
  return pixels * 72 / normalizeDpi(dpi);
}

export function getMegaPixels(width: number, height: number): number {
  return (width * height) / 1_000_000;
}

export function assertCanvasSize(width: number, height: number): void {
  const megaPixels = getMegaPixels(width, height);
  if (megaPixels > MAX_PAGE_MEGA_PIXELS) {
    throw new Error(`PNG 导出尺寸约 ${megaPixels.toFixed(1)} MP，超过安全上限。请降低 DPI 或缩小选区后重试。`);
  }
}

export function normalizePageIndices(pageIndices: number[] | undefined, pageCount: number): number[] {
  if (!pageIndices?.length) {
    return Array.from({ length: pageCount }, (_, index) => index);
  }

  return [...new Set(pageIndices.filter((index) => Number.isInteger(index) && index >= 0 && index < pageCount))].sort((a, b) => a - b);
}

export function getUnionBounds(boxes: ShapeBox[], paddingPoints: number, slideWidthPoints: number, slideHeightPoints: number): Bounds {
  if (boxes.length === 0) {
    throw new Error("请先选择要导出的图片或形状。");
  }

  const padding = normalizePaddingPoints(paddingPoints);
  const left = Math.max(0, Math.min(...boxes.map((box) => box.left)) - padding);
  const top = Math.max(0, Math.min(...boxes.map((box) => box.top)) - padding);
  const right = Math.min(slideWidthPoints, Math.max(...boxes.map((box) => box.left + box.width)) + padding);
  const bottom = Math.min(slideHeightPoints, Math.max(...boxes.map((box) => box.top + box.height)) + padding);

  return {
    left,
    top,
    width: Math.max(1, right - left),
    height: Math.max(1, bottom - top),
  };
}
