import { describe, expect, it } from "vitest";
import { assertCanvasSize, dpiToScale, getMegaPixels, getUnionBounds, normalizeDpi, normalizePaddingPoints, normalizePageIndices, pixelsToPoints, pointsToPixels } from "../features/export/pngExportMath";
import type { ShapeBox } from "../shared/types";

const boxes: ShapeBox[] = [
  { id: "a", name: "A", type: "image", left: 10, top: 20, width: 100, height: 60 },
  { id: "b", name: "B", type: "shape", left: 90, top: 70, width: 80, height: 40 },
];

describe("pngExport", () => {
  it("converts DPI to PDF.js scale", () => {
    expect(dpiToScale(300)).toBeCloseTo(4.1666667);
    expect(dpiToScale(600)).toBeCloseTo(8.3333333);
  });

  it("normalizes invalid or unsafe DPI values", () => {
    expect(normalizeDpi(Number.NaN)).toBe(300);
    expect(normalizeDpi(20)).toBe(72);
    expect(normalizeDpi(300.4)).toBe(300);
    expect(normalizeDpi(1400)).toBe(1200);
  });

  it("normalizes padding values", () => {
    expect(normalizePaddingPoints(Number.NaN)).toBe(2);
    expect(normalizePaddingPoints(-2)).toBe(0);
    expect(normalizePaddingPoints(8.5)).toBe(8.5);
    expect(normalizePaddingPoints(80)).toBe(36);
  });

  it("converts points and pixels using DPI", () => {
    expect(pointsToPixels(72, 300)).toBe(300);
    expect(pixelsToPoints(300, 300)).toBe(72);
  });

  it("computes megapixels", () => {
    expect(getMegaPixels(10_000, 10_000)).toBe(100);
  });

  it("rejects excessive canvas sizes", () => {
    expect(() => assertCanvasSize(10_001, 10_001)).toThrow("超过安全上限");
  });

  it("uses all pages when no page indices are supplied", () => {
    expect(normalizePageIndices(undefined, 3)).toEqual([0, 1, 2]);
  });

  it("deduplicates, filters, and sorts requested page indices", () => {
    expect(normalizePageIndices([2, 0, 2, -1, 8, 1], 4)).toEqual([0, 1, 2]);
  });

  it("computes padded union bounds for selected shapes", () => {
    expect(getUnionBounds(boxes, 5, 200, 150)).toEqual({ left: 5, top: 15, width: 170, height: 100 });
  });

  it("clamps union bounds to slide edges", () => {
    expect(getUnionBounds(boxes, 50, 160, 100)).toEqual({ left: 0, top: 0, width: 160, height: 100 });
  });
});
