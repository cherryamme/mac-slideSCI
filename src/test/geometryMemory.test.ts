import { describe, expect, it } from "vitest";
import type { ShapeBox } from "../shared/types";
import { applyGeometryMemory, toGeometryMemory } from "../features/formatMemory/geometryMemory";

const box: ShapeBox = {
  id: "shape-1",
  name: "Shape 1",
  type: "image",
  left: 10,
  top: 20,
  width: 100,
  height: 80,
};

const memory = {
  left: 50,
  top: 60,
  width: 200,
  height: 160,
};

describe("geometryMemory", () => {
  it("copies geometry from a shape box", () => {
    expect(toGeometryMemory(box)).toEqual({ left: 10, top: 20, width: 100, height: 80 });
  });

  it("applies position only", () => {
    expect(applyGeometryMemory(box, memory, "position")).toEqual({ ...box, left: 50, top: 60 });
  });

  it("applies size only", () => {
    expect(applyGeometryMemory(box, memory, "size")).toEqual({ ...box, width: 200, height: 160 });
  });

  it("applies position and size", () => {
    expect(applyGeometryMemory(box, memory, "all")).toEqual({ ...box, left: 50, top: 60, width: 200, height: 160 });
  });
});
