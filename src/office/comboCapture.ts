import { isPowerPointApiSupported } from "./requirements";
import { ensurePowerPointApi, runPowerPoint } from "./powerpoint";
import { renderSelectedObjectsPng } from "./selectedObjectImage";
import type { ComboElement, ComboElementType, OperationResult, SavedCombo } from "../shared/types";

const DEFAULT_SLIDE_WIDTH = 960;
const DEFAULT_SLIDE_HEIGHT = 540;
const PADDING_POINTS = 4;
const THUMBNAIL_DPI = 144;

const TEXT_BEARING_TYPES = new Set(["TextBox", "GeometricShape", "Callout", "Placeholder"]);

function getShapeKind(type: string): ComboElementType {
  if (type === "TextBox") {
    return "textBox";
  }
  if (type === "GeometricShape" || type === "Callout" || type === "Placeholder") {
    return "geometricShape";
  }
  if (type === "Line") {
    return "line";
  }
  if (type === "Group") {
    return "group";
  }
  return "unsupported";
}

function shapeTypeToGeometric(shapeName: string): string {
  const name = (shapeName ?? "").toLowerCase();
  if (name.includes("oval") || name.includes("circle") || name.includes("ellipse")) {
    return "Ellipse";
  }
  if (name.includes("right arrow")) {
    return "RightArrow";
  }
  if (name.includes("left arrow")) {
    return "LeftArrow";
  }
  if (name.includes("up arrow")) {
    return "UpArrow";
  }
  if (name.includes("down arrow")) {
    return "DownArrow";
  }
  if (name.includes("round")) {
    return "RoundRectangle";
  }
  return "Rectangle";
}

function safeNumber(value: unknown, fallback = 0): number {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function readBoolean(value: boolean | null | undefined): boolean | undefined {
  if (value === null || value === undefined) {
    return undefined;
  }
  return Boolean(value);
}

interface ShapeSnapshot {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  children?: ShapeSnapshot[];
}

async function snapshotShapeArray(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[], depth = 0): Promise<ShapeSnapshot[]> {
  if (shapes.length === 0) {
    return [];
  }

  shapes.forEach((shape) => {
    shape.load("id,name,type,left,top,width,height");
  });
  await context.sync();

  const supportsGroupApi = isPowerPointApiSupported("1.8");
  type ShapeLikeCollection = { items: PowerPoint.Shape[]; load: (option: string) => void };
  const groupCollections: Array<{ collection: ShapeLikeCollection } | null> = shapes.map((shape) => {
    if (!supportsGroupApi) {
      return null;
    }
    const type = String(shape.type ?? "");
    if (type !== "Group") {
      return null;
    }
    try {
      const collection = shape.group.shapes as unknown as ShapeLikeCollection;
      collection.load("items");
      return { collection };
    } catch {
      return null;
    }
  });

  if (groupCollections.some((entry) => entry !== null)) {
    await context.sync();
  }

  const textFrames = shapes.map((shape) => {
    const type = String(shape.type ?? "");
    if (!TEXT_BEARING_TYPES.has(type)) {
      return null;
    }
    try {
      return shape.textFrame;
    } catch {
      return null;
    }
  });

  textFrames.forEach((frame) => {
    if (!frame) {
      return;
    }
    try {
      frame.textRange.load("text");
      frame.textRange.font.load("name,size,color,bold,italic");
    } catch {
      // shape 不支持 textFrame，跳过
    }
  });

  try {
    await context.sync();
  } catch {
    // 即使部分 shape 不支持 textFrame，也允许继续
  }

  const snapshots: ShapeSnapshot[] = [];
  for (let index = 0; index < shapes.length; index += 1) {
    const shape = shapes[index];
    const snapshot: ShapeSnapshot = {
      id: String(shape.id ?? `shape-${depth}-${index}`),
      name: String(shape.name ?? ""),
      type: String(shape.type ?? "Unsupported"),
      left: safeNumber(shape.left),
      top: safeNumber(shape.top),
      width: safeNumber(shape.width),
      height: safeNumber(shape.height),
    };

    const frame = textFrames[index];
    if (frame) {
      try {
        const range = frame.textRange;
        const text = typeof range.text === "string" ? range.text : "";
        if (text.length > 0) {
          snapshot.text = text;
        }
        if (range.font.name) {
          snapshot.fontName = String(range.font.name);
        }
        const size = safeNumber(range.font.size, 0);
        if (size > 0) {
          snapshot.fontSize = size;
        }
        if (typeof range.font.color === "string" && range.font.color.length > 0) {
          snapshot.fontColor = range.font.color;
        }
        snapshot.bold = readBoolean(range.font.bold);
        snapshot.italic = readBoolean(range.font.italic);
      } catch {
        // 不影响整体捕获
      }
    }

    const groupEntry = groupCollections[index];
    if (groupEntry) {
      try {
        snapshot.children = await snapshotShapeArray(context, groupEntry.collection.items, depth + 1);
      } catch {
        snapshot.children = [];
      }
    }

    snapshots.push(snapshot);
  }

  return snapshots;
}

async function loadSelectionSnapshots(context: PowerPoint.RequestContext): Promise<ShapeSnapshot[]> {
  const shapes = context.presentation.getSelectedShapes();
  shapes.load("items");
  await context.sync();

  if (shapes.items.length === 0) {
    throw new Error("请先选择要保存的对象。");
  }

  return snapshotShapeArray(context, shapes.items);
}

function snapshotToElement(snapshot: ShapeSnapshot, originLeft: number, originTop: number, depth = 0, position = 0): ComboElement {
  const kind = getShapeKind(snapshot.type);
  const element: ComboElement = {
    id: `${snapshot.id}-${depth}-${position}`,
    type: kind,
    relativeLeft: snapshot.left - originLeft,
    relativeTop: snapshot.top - originTop,
    width: snapshot.width,
    height: snapshot.height,
  };

  if (kind === "textBox" || kind === "geometricShape") {
    if (snapshot.text) {
      element.text = snapshot.text;
    }
    if (snapshot.fontName) {
      element.fontName = snapshot.fontName;
    }
    if (snapshot.fontSize) {
      element.fontSize = snapshot.fontSize;
    }
    if (snapshot.fontColor) {
      element.fontColor = snapshot.fontColor;
    }
    if (snapshot.bold !== undefined) {
      element.bold = snapshot.bold;
    }
    if (snapshot.italic !== undefined) {
      element.italic = snapshot.italic;
    }
  }

  if (kind === "geometricShape") {
    element.shapeType = shapeTypeToGeometric(snapshot.name);
  }

  if (kind === "line") {
    element.connectorType = "Straight";
  }

  if (kind === "group" && snapshot.children?.length) {
    element.children = snapshot.children.map((child, childIndex) => snapshotToElement(child, snapshot.left, snapshot.top, depth + 1, childIndex));
  }

  return element;
}

function snapshotToElements(snapshots: ShapeSnapshot[]): { elements: ComboElement[]; width: number; height: number } {
  const lefts = snapshots.map((shape) => shape.left);
  const tops = snapshots.map((shape) => shape.top);
  const rights = snapshots.map((shape) => shape.left + shape.width);
  const bottoms = snapshots.map((shape) => shape.top + shape.height);
  const unionLeft = Math.min(...lefts);
  const unionTop = Math.min(...tops);
  const unionRight = Math.max(...rights);
  const unionBottom = Math.max(...bottoms);

  const elements = snapshots.map((snapshot, index) => snapshotToElement(snapshot, unionLeft, unionTop, 0, index));
  return {
    elements,
    width: Math.max(1, unionRight - unionLeft),
    height: Math.max(1, unionBottom - unionTop),
  };
}

function countUnsupported(elements: ComboElement[]): number {
  let total = 0;
  elements.forEach((element) => {
    if (element.type === "unsupported") {
      total += 1;
      return;
    }
    if (element.children?.length) {
      total += countUnsupported(element.children);
    }
  });
  return total;
}

function hasSupportedLeaf(elements: ComboElement[]): boolean {
  return elements.some((element) => {
    if (element.type === "textBox" || element.type === "geometricShape" || element.type === "line") {
      return true;
    }
    if (element.type === "group" && element.children?.length) {
      return hasSupportedLeaf(element.children);
    }
    return false;
  });
}

async function tryRenderThumbnail(): Promise<string> {
  if (!isPowerPointApiSupported("1.8") && !isPowerPointApiSupported("1.10")) {
    return "";
  }
  try {
    const png = await renderSelectedObjectsPng({ dpi: THUMBNAIL_DPI, paddingPoints: PADDING_POINTS });
    return await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : "");
      reader.onerror = () => reject(reader.error ?? new Error("无法读取缩略图。"));
      reader.readAsDataURL(png.blob);
    });
  } catch {
    return "";
  }
}

export async function captureSelectionAsEditableCombo(name: string): Promise<SavedCombo> {
  ensurePowerPointApi("1.4", "保存可编辑组合");

  let snapshots: ShapeSnapshot[] = [];
  await runPowerPoint(async (context) => {
    snapshots = await loadSelectionSnapshots(context);
  });

  if (snapshots.length === 0) {
    throw new Error("请先选择要保存的对象。");
  }

  const { elements, width, height } = snapshotToElements(snapshots);
  if (!hasSupportedLeaf(elements)) {
    throw new Error("选中的对象暂不能保存为可编辑组合（仅支持文本框、几何形状、线条和它们的分组）。");
  }

  const thumbnailBase64 = await tryRenderThumbnail();
  const unsupportedCount = countUnsupported(elements);
  const now = Date.now();
  const trimmedName = name.trim() || `组合 ${new Date().toLocaleString()}`;

  return {
    id: `combo-${now}-${Math.random().toString(36).slice(2, 8)}`,
    name: trimmedName,
    createdAt: now,
    updatedAt: now,
    width,
    height,
    thumbnailBase64,
    elements,
    unsupportedCount,
  };
}

function configureTextStyle(shape: PowerPoint.Shape, element: ComboElement): void {
  try {
    const range = shape.textFrame.textRange;
    if (element.text !== undefined) {
      range.text = element.text;
    }
    if (element.fontName) {
      range.font.name = element.fontName;
    }
    if (element.fontSize) {
      range.font.size = element.fontSize;
    }
    if (element.fontColor) {
      range.font.color = element.fontColor;
    }
    if (element.bold !== undefined) {
      range.font.bold = element.bold;
    }
    if (element.italic !== undefined) {
      range.font.italic = element.italic;
    }
  } catch {
    // 当前 shape 不支持 textFrame，忽略
  }
}

function clampInsertOrigin(slideWidth: number, slideHeight: number, comboWidth: number, comboHeight: number): { left: number; top: number } {
  const left = Math.max(0, (slideWidth - comboWidth) / 2);
  const top = Math.max(0, (slideHeight - comboHeight) / 2);
  return { left, top };
}

function insertLeafElement(slide: PowerPoint.Slide, element: ComboElement, left: number, top: number, combo: SavedCombo): { shape: PowerPoint.Shape | null; skipped: number } {
  const options: PowerPoint.ShapeAddOptions = {
    left,
    top,
    width: Math.max(1, element.width),
    height: Math.max(1, element.height),
  };

  let shape: PowerPoint.Shape | null = null;
  try {
    if (element.type === "textBox") {
      shape = slide.shapes.addTextBox(element.text ?? "", options);
    } else if (element.type === "geometricShape") {
      const geom = (element.shapeType ?? "Rectangle") as PowerPoint.GeometricShapeType;
      shape = slide.shapes.addGeometricShape(geom, options);
    } else if (element.type === "line") {
      const connector = (element.connectorType ?? "Straight") as PowerPoint.ConnectorType;
      shape = slide.shapes.addLine(connector, options);
    }
  } catch {
    shape = null;
  }

  if (!shape) {
    return { shape: null, skipped: 1 };
  }

  try {
    shape.name = `mac-slideSCI:combo:${combo.id}:${element.id}`;
  } catch {
    // 命名失败忽略
  }

  if (element.type !== "line") {
    configureTextStyle(shape, element);
  }

  return { shape, skipped: 0 };
}

interface InsertOutcome {
  inserted: number;
  skipped: number;
  shapes: PowerPoint.Shape[];
}

function insertElements(slide: PowerPoint.Slide, elements: ComboElement[], originLeft: number, originTop: number, combo: SavedCombo): InsertOutcome {
  let inserted = 0;
  let skipped = 0;
  const shapes: PowerPoint.Shape[] = [];

  elements.forEach((element) => {
    const absoluteLeft = originLeft + element.relativeLeft;
    const absoluteTop = originTop + element.relativeTop;

    if (element.type === "group" && element.children?.length) {
      const outcome = insertElements(slide, element.children, absoluteLeft, absoluteTop, combo);
      inserted += outcome.inserted;
      skipped += outcome.skipped;
      shapes.push(...outcome.shapes);
      return;
    }

    if (element.type === "unsupported") {
      skipped += 1;
      return;
    }

    const result = insertLeafElement(slide, element, absoluteLeft, absoluteTop, combo);
    if (result.shape) {
      inserted += 1;
      shapes.push(result.shape);
    } else {
      skipped += 1;
    }
  });

  return { inserted, skipped, shapes };
}

function regroupTopLevel(slide: PowerPoint.Slide, elements: ComboElement[], outcomes: PowerPoint.Shape[][]): number {
  if (!isPowerPointApiSupported("1.8")) {
    return 0;
  }

  let grouped = 0;
  elements.forEach((element, index) => {
    if (element.type !== "group") {
      return;
    }
    const members = outcomes[index];
    if (!members || members.length < 2) {
      return;
    }
    try {
      slide.shapes.addGroup(members);
      grouped += 1;
    } catch {
      // group API 在当前版本不可用或失败，忽略
    }
  });
  return grouped;
}

export async function insertEditableCombo(combo: SavedCombo): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "插入组合");

  return runPowerPoint(async (context) => {
    const presentation = context.presentation;
    const slide = presentation.getSelectedSlides().getItemAt(0);

    let slideWidth = DEFAULT_SLIDE_WIDTH;
    let slideHeight = DEFAULT_SLIDE_HEIGHT;
    if (isPowerPointApiSupported("1.10")) {
      try {
        const pageSetup = presentation.pageSetup;
        pageSetup.load("slideWidth,slideHeight");
        await context.sync();
        slideWidth = safeNumber(pageSetup.slideWidth, DEFAULT_SLIDE_WIDTH);
        slideHeight = safeNumber(pageSetup.slideHeight, DEFAULT_SLIDE_HEIGHT);
      } catch {
        // 旧版本无 pageSetup，沿用默认
      }
    }

    const origin = clampInsertOrigin(slideWidth, slideHeight, combo.width, combo.height);
    const perElementShapes: PowerPoint.Shape[][] = [];

    let totalInserted = 0;
    let totalSkipped = 0;

    combo.elements.forEach((element) => {
      const absoluteLeft = origin.left + element.relativeLeft;
      const absoluteTop = origin.top + element.relativeTop;

      if (element.type === "group" && element.children?.length) {
        const outcome = insertElements(slide, element.children, absoluteLeft, absoluteTop, combo);
        totalInserted += outcome.inserted;
        totalSkipped += outcome.skipped;
        perElementShapes.push(outcome.shapes);
        return;
      }

      if (element.type === "unsupported") {
        totalSkipped += 1;
        perElementShapes.push([]);
        return;
      }

      const result = insertLeafElement(slide, element, absoluteLeft, absoluteTop, combo);
      if (result.shape) {
        totalInserted += 1;
        perElementShapes.push([result.shape]);
      } else {
        totalSkipped += 1;
        perElementShapes.push([]);
      }
    });

    await context.sync();

    let groupedCount = 0;
    try {
      groupedCount = regroupTopLevel(slide, combo.elements, perElementShapes);
      if (groupedCount > 0) {
        await context.sync();
      }
    } catch {
      groupedCount = 0;
    }

    if (totalInserted === 0) {
      throw new Error("当前 PowerPoint 版本无法插入组合中的元素，请确认 Office.js API 已更新。");
    }

    const notes: string[] = [];
    if (totalSkipped > 0) {
      notes.push(`跳过 ${totalSkipped} 个不支持的元素`);
    }
    if (groupedCount > 0) {
      notes.push(`已重新组合 ${groupedCount} 组对象`);
    }
    const suffix = notes.length > 0 ? `，${notes.join("，")}` : "";
    return { count: totalInserted, message: `已插入组合 "${combo.name}"：${totalInserted} 个元素${suffix}。` };
  });
}
