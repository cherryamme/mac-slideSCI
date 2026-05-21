import type { GeometryApplyMode, GeometryMemory, ShapeBox } from "../../shared/types";

export function toGeometryMemory(box: ShapeBox): GeometryMemory {
  return {
    left: box.left,
    top: box.top,
    width: box.width,
    height: box.height,
  };
}

export function applyGeometryMemory(box: ShapeBox, memory: GeometryMemory, mode: GeometryApplyMode): ShapeBox {
  return {
    ...box,
    left: mode === "size" ? box.left : memory.left,
    top: mode === "size" ? box.top : memory.top,
    width: mode === "position" ? box.width : memory.width,
    height: mode === "position" ? box.height : memory.height,
  };
}
