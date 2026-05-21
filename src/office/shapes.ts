import { getPanelLabel } from "../features/labels/labelGenerator";
import { sortByPosition } from "../features/layout/layoutEngine";
import type { CaptionSettings, GeometryApplyMode, GeometryMemory, LabelSettings, OperationResult, ShapeBox } from "../shared/types";
import { applyGeometryMemory, toGeometryMemory } from "../features/formatMemory/geometryMemory";
import { ensurePowerPointApi, runPowerPoint } from "./powerpoint";

const SHAPE_LOAD = "items/id,items/name,items/type,items/left,items/top,items/width,items/height";

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

function formatCaptionText(settings: CaptionSettings, index: number): string {
  const letter = String.fromCharCode("A".charCodeAt(0) + (index % 26));
  return settings.text.replaceAll("{index}", String(index + 1)).replaceAll("{letter}", letter);
}

function applyTextStyle(shape: PowerPoint.Shape, fontSize: number, color: string, bold = false): void {
  shape.textFrame.textRange.font.size = fontSize;
  shape.textFrame.textRange.font.color = color;
  shape.textFrame.textRange.font.bold = bold;
}

export async function getSelectedShapeBoxes(): Promise<ShapeBox[]> {
  ensurePowerPointApi("1.6", "读取选中对象");

  return runPowerPoint(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load(SHAPE_LOAD);
    await context.sync();
    return shapes.items.map(toShapeBox);
  });
}

export async function applyShapeBoxes(targetBoxes: ShapeBox[]): Promise<OperationResult> {
  ensurePowerPointApi("1.6", "更新选中对象位置和尺寸");

  return runPowerPoint(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();

    const byId = new Map(targetBoxes.map((box) => [box.id, box]));
    let count = 0;

    shapes.items.forEach((shape, index) => {
      const target = byId.get(shape.id) ?? targetBoxes[index];
      if (!target) {
        return;
      }

      shape.left = target.left;
      shape.top = target.top;
      shape.width = target.width;
      shape.height = target.height;
      count += 1;
    });

    await context.sync();
    return { count, message: `已更新 ${count} 个对象。` };
  });
}

export async function copyGeometryFromSelection(): Promise<GeometryMemory> {
  const boxes = await getSelectedShapeBoxes();
  if (boxes.length === 0) {
    throw new Error("请先选择一个对象。");
  }

  return toGeometryMemory(boxes[0]);
}

export async function applyGeometryToSelection(memory: GeometryMemory, mode: GeometryApplyMode): Promise<OperationResult> {
  const boxes = await getSelectedShapeBoxes();
  if (boxes.length === 0) {
    throw new Error("请先选择要应用的位置或尺寸的对象。");
  }

  const targets = boxes.map((box) => applyGeometryMemory(box, memory, mode));
  return applyShapeBoxes(targets);
}

export async function addPanelLabelsToSelection(settings: LabelSettings): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "添加标签");
  ensurePowerPointApi("1.6", "读取选中对象");

  return runPowerPoint(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = context.presentation.getSelectedShapes();
    shapes.load(SHAPE_LOAD);
    await context.sync();

    const boxes = sortByPosition(shapes.items.map(toShapeBox));
    boxes.forEach((box, index) => {
      const label = slide.shapes.addTextBox(getPanelLabel(index, settings), {
        left: box.left + settings.offsetX,
        top: box.top + settings.offsetY,
        width: settings.width,
        height: settings.height,
      });
      label.name = `mac-slideSCI:label:${box.id}`;
      applyTextStyle(label, settings.fontSize, settings.color, settings.bold);
      if (settings.backgroundColor !== "transparent") {
        label.fill.setSolidColor(settings.backgroundColor);
      }
    });

    await context.sync();
    return { count: boxes.length, message: `已添加 ${boxes.length} 个 panel label。` };
  });
}

export async function addCaptionsToSelection(settings: CaptionSettings): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "添加标题");
  ensurePowerPointApi("1.6", "读取选中对象");

  return runPowerPoint(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = context.presentation.getSelectedShapes();
    shapes.load(SHAPE_LOAD);
    await context.sync();

    const boxes = sortByPosition(shapes.items.map(toShapeBox));
    boxes.forEach((box, index) => {
      const caption = slide.shapes.addTextBox(formatCaptionText(settings, index), {
        left: box.left,
        top: box.top + box.height + settings.gap,
        width: box.width,
        height: settings.height,
      });
      caption.name = `mac-slideSCI:caption:${box.id}`;
      applyTextStyle(caption, settings.fontSize, settings.color);
    });

    await context.sync();
    return { count: boxes.length, message: `已添加 ${boxes.length} 个 caption。` };
  });
}

export async function insertLineNearSelection(): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "插入标注线");

  return runPowerPoint(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: 160,
      top: 160,
      width: 120,
      height: 0,
    });
    line.name = "mac-slideSCI:line";
    await context.sync();
    return { count: 1, message: "已插入标注线。" };
  });
}

export async function insertCalloutNearSelection(): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "插入 callout");

  return runPowerPoint(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const callout = slide.shapes.addTextBox("Callout", {
      left: 160,
      top: 190,
      width: 140,
      height: 36,
    });
    callout.name = "mac-slideSCI:callout";
    callout.fill.setSolidColor("#FFF7CC");
    applyTextStyle(callout, 14, "#111111", true);
    await context.sync();
    return { count: 1, message: "已插入 callout 文本框。" };
  });
}

export async function insertArrowNearSelection(): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "插入箭头");

  return runPowerPoint(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const arrow = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rightArrow, {
      left: 160,
      top: 245,
      width: 120,
      height: 28,
    });
    arrow.name = "mac-slideSCI:arrow";
    arrow.fill.setSolidColor("#C62828");
    await context.sync();
    return { count: 1, message: "已插入箭头。" };
  });
}
