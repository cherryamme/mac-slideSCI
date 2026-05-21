import { applyGeometryToSelection, applyShapeBoxes, copyGeometryFromSelection, addCaptionsToSelection, addPanelLabelsToSelection, getSelectedShapeBoxes, insertArrowNearSelection, insertCalloutNearSelection, insertLineNearSelection } from "../office/shapes";
import { applyFontToAllSlides, applyFontToCurrentSlide, applyFontToSelection } from "../office/font";
import { captureSelectionAsEditableCombo, insertEditableCombo } from "../office/comboCapture";
import { alignShapes, arrangeGrid, arrangeHorizontal, arrangeVertical, distributeShapes, equalizeShapes } from "../features/layout/layoutEngine";
import { loadCaptionSettings, loadFontUnifySettings, loadGeometryMemory, loadLabelSettings, loadLayoutSettings, saveGeometryMemory } from "./settings";
import { deleteCombo, getCombo, listCombos, renameCombo, saveCombo } from "./comboStore";
import type { AlignmentMode, DistributionMode, EqualizeMode, GeometryApplyMode, OperationResult, SavedCombo, ShapeBox } from "./types";

async function transformSelection(transform: (boxes: ShapeBox[]) => ShapeBox[]): Promise<OperationResult> {
  const boxes = await getSelectedShapeBoxes();
  if (boxes.length === 0) {
    throw new Error("请先选择需要处理的图片或形状。");
  }

  return applyShapeBoxes(transform(boxes));
}

export function alignSelection(mode: AlignmentMode): Promise<OperationResult> {
  return transformSelection((boxes) => alignShapes(boxes, mode));
}

export function distributeSelection(mode: DistributionMode): Promise<OperationResult> {
  return transformSelection((boxes) => distributeShapes(boxes, mode));
}

export function equalizeSelection(mode: EqualizeMode): Promise<OperationResult> {
  return transformSelection((boxes) => equalizeShapes(boxes, mode));
}

export function arrangeSelectionHorizontal(): Promise<OperationResult> {
  const settings = loadLayoutSettings();
  return transformSelection((boxes) => arrangeHorizontal(boxes, settings.horizontalSpacing));
}

export function arrangeSelectionVertical(): Promise<OperationResult> {
  const settings = loadLayoutSettings();
  return transformSelection((boxes) => arrangeVertical(boxes, settings.verticalSpacing));
}

export function arrangeSelectionGrid(): Promise<OperationResult> {
  const settings = loadLayoutSettings();
  return transformSelection((boxes) => arrangeGrid(boxes, {
    columns: settings.columns,
    spacing: {
      horizontal: settings.horizontalSpacing,
      vertical: settings.verticalSpacing,
    },
    equalizeSize: settings.equalizeGrid,
  }));
}

export function addDefaultPanelLabels(): Promise<OperationResult> {
  return addPanelLabelsToSelection(loadLabelSettings());
}

export function addDefaultCaptions(): Promise<OperationResult> {
  return addCaptionsToSelection(loadCaptionSettings());
}

export async function copySelectionGeometry(): Promise<OperationResult> {
  const memory = await copyGeometryFromSelection();
  saveGeometryMemory(memory);
  return { count: 1, message: "已复制当前对象的位置和尺寸。" };
}

export function pasteSelectionGeometry(mode: GeometryApplyMode): Promise<OperationResult> {
  const memory = loadGeometryMemory();
  if (!memory) {
    throw new Error("请先复制一个对象的位置或尺寸。");
  }

  return applyGeometryToSelection(memory, mode);
}

export function insertLine(): Promise<OperationResult> {
  return insertLineNearSelection();
}

export function insertArrow(): Promise<OperationResult> {
  return insertArrowNearSelection();
}

export function insertCallout(): Promise<OperationResult> {
  return insertCalloutNearSelection();
}

export function unifyFontSelection(): Promise<OperationResult> {
  return applyFontToSelection(loadFontUnifySettings());
}

export function unifyFontCurrentSlide(): Promise<OperationResult> {
  return applyFontToCurrentSlide(loadFontUnifySettings());
}

export function unifyFontAllSlides(): Promise<OperationResult> {
  return applyFontToAllSlides(loadFontUnifySettings());
}

export async function saveSelectionAsCombo(name?: string): Promise<OperationResult> {
  const combo = await captureSelectionAsEditableCombo(name ?? "");
  await saveCombo(combo);
  const note = combo.unsupportedCount > 0 ? `，其中 ${combo.unsupportedCount} 个复杂对象暂未保存为可编辑元素` : "";
  return { count: combo.elements.length, message: `已保存组合 "${combo.name}"${note}。` };
}

export async function insertSavedCombo(id: string): Promise<OperationResult> {
  const combo = await getCombo(id);
  if (!combo) {
    throw new Error("组合不存在或已被删除。");
  }
  return insertEditableCombo(combo);
}

export async function deleteSavedCombo(id: string): Promise<OperationResult> {
  await deleteCombo(id);
  return { count: 1, message: "已删除组合。" };
}

export async function renameSavedCombo(id: string, name: string): Promise<OperationResult> {
  const combo = await renameCombo(id, name);
  if (!combo) {
    throw new Error("组合不存在。");
  }
  return { count: 1, message: `已重命名为 "${combo.name}"。` };
}

export async function listSavedCombos(): Promise<SavedCombo[]> {
  return listCombos();
}
