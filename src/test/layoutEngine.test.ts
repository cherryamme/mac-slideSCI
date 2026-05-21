import { describe, expect, it } from "vitest";
import type { ShapeBox } from "../shared/types";
import { alignShapes, arrangeGrid, arrangeHorizontal, arrangeVertical, distributeShapes, equalizeShapes, sortByPosition } from "../features/layout/layoutEngine";

const boxes: ShapeBox[] = [
  { id: "b", name: "B", type: "image", left: 180, top: 20, width: 50, height: 40 },
  { id: "a", name: "A", type: "image", left: 20, top: 20, width: 60, height: 30 },
  { id: "c", name: "C", type: "image", left: 90, top: 100, width: 40, height: 50 },
];

describe("layoutEngine", () => {
  it("sorts shapes by visual row then left position", () => {
    expect(sortByPosition(boxes).map((box) => box.id)).toEqual(["a", "b", "c"]);
  });

  it("arranges shapes horizontally", () => {
    const result = arrangeHorizontal(boxes, 10);
    expect(result.map((box) => [box.id, box.left, box.top])).toEqual([
      ["a", 20, 20],
      ["b", 90, 20],
      ["c", 150, 20],
    ]);
  });

  it("arranges shapes vertically", () => {
    const result = arrangeVertical(boxes, 8);
    expect(result.map((box) => [box.id, box.left, box.top])).toEqual([
      ["a", 20, 20],
      ["b", 20, 58],
      ["c", 20, 106],
    ]);
  });

  it("arranges shapes into an equal-size grid", () => {
    const result = arrangeGrid(boxes, { columns: 2, spacing: { horizontal: 12, vertical: 16 }, equalizeSize: true });
    expect(result.map((box) => [box.id, box.left, box.top, box.width, box.height])).toEqual([
      ["a", 20, 20, 60, 50],
      ["b", 92, 20, 60, 50],
      ["c", 20, 86, 60, 50],
    ]);
  });

  it("aligns shapes to the right edge", () => {
    const result = alignShapes(boxes, "right");
    expect(result.map((box) => [box.id, box.left])).toEqual([
      ["b", 180],
      ["a", 170],
      ["c", 190],
    ]);
  });

  it("distributes shapes horizontally", () => {
    const result = distributeShapes(boxes, "horizontal");
    expect(result.map((box) => [box.id, Math.round(box.left)])).toEqual([
      ["a", 20],
      ["c", 110],
      ["b", 180],
    ]);
  });

  it("equalizes width and height to the largest selected shape", () => {
    const result = equalizeShapes(boxes, "both");
    expect(result.every((box) => box.width === 60 && box.height === 50)).toBe(true);
  });
});
