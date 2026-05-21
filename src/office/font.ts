import { isPowerPointApiSupported } from "./requirements";
import { ensurePowerPointApi, runPowerPoint } from "./powerpoint";
import { resolveFontName } from "../shared/settings";
import type { FontHorizontalAlignment, FontUnifySettings, OperationResult } from "../shared/types";

const TEXT_BEARING_TYPES = new Set(["TextBox", "GeometricShape", "Callout", "Placeholder"]);
const GROUP_TYPE = "Group";
const TABLE_TYPE = "Table";

interface AppliedCount {
  shapes: number;
  groups: number;
  tables: number;
  skipped: number;
}

function alignmentValue(alignment: FontHorizontalAlignment): "Left" | "Center" | "Right" | "Justify" {
  switch (alignment) {
    case "center":
      return "Center";
    case "right":
      return "Right";
    case "justify":
      return "Justify";
    case "left":
    default:
      return "Left";
  }
}

function applyFontToShape(shape: PowerPoint.Shape, settings: FontUnifySettings, counts: AppliedCount): void {
  const type = String(shape.type ?? "");
  if (type === GROUP_TYPE) {
    counts.groups += 1;
    return;
  }
  if (type === TABLE_TYPE) {
    counts.tables += 1;
    return;
  }
  if (!TEXT_BEARING_TYPES.has(type)) {
    counts.skipped += 1;
    return;
  }

  try {
    const range = shape.textFrame.textRange;
    range.font.name = resolveFontName(settings);
    if (settings.unifyAllFormatting) {
      range.font.size = settings.fontSize;
      range.font.bold = settings.bold;
      range.font.italic = settings.italic;
      range.paragraphFormat.horizontalAlignment = alignmentValue(settings.horizontalAlignment) as PowerPoint.ParagraphHorizontalAlignment;
    }
    counts.shapes += 1;
  } catch {
    counts.skipped += 1;
  }
}

async function processShapes(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[], settings: FontUnifySettings, counts: AppliedCount): Promise<void> {
  shapes.forEach((shape) => applyFontToShape(shape, settings, counts));

  if (!isPowerPointApiSupported("1.8")) {
    return;
  }

  const groups = shapes.filter((shape) => String(shape.type ?? "") === GROUP_TYPE);
  if (groups.length === 0) {
    return;
  }

  groups.forEach((shape) => shape.group.shapes.load("items/type"));
  await context.sync();

  for (const shape of groups) {
    await processShapes(context, shape.group.shapes.items, settings, counts);
  }
}

function summarize(counts: AppliedCount, scope: string, fontName: string): OperationResult {
  if (counts.shapes === 0 && counts.skipped === 0 && counts.tables === 0 && counts.groups === 0) {
    return { count: 0, message: `${scope}没有可处理的文本对象。` };
  }

  const parts: string[] = [`已为 ${counts.shapes} 个${scope}文本设置字体为 ${fontName}`];
  if (counts.groups > 0) {
    parts.push(`递归处理 ${counts.groups} 个分组`);
  }
  if (counts.tables > 0) {
    parts.push(`${counts.tables} 个表格暂未自动应用，请直接在 PowerPoint 中调整`);
  }
  if (counts.skipped > 0) {
    parts.push(`跳过 ${counts.skipped} 个无文本对象`);
  }

  return { count: counts.shapes, message: `${parts.join("，")}。` };
}

export async function applyFontToSelection(settings: FontUnifySettings): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "字体统一");

  return runPowerPoint(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();

    if (shapes.items.length === 0) {
      throw new Error("请先选择至少一个对象。");
    }

    const counts: AppliedCount = { shapes: 0, groups: 0, tables: 0, skipped: 0 };
    await processShapes(context, shapes.items, settings, counts);
    await context.sync();
    return summarize(counts, "选中对象", resolveFontName(settings));
  });
}

export async function applyFontToCurrentSlide(settings: FontUnifySettings): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "字体统一");

  return runPowerPoint(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load("items/type");
    await context.sync();

    const counts: AppliedCount = { shapes: 0, groups: 0, tables: 0, skipped: 0 };
    await processShapes(context, slide.shapes.items, settings, counts);
    await context.sync();
    return summarize(counts, "本页", resolveFontName(settings));
  });
}

export async function applyFontToAllSlides(settings: FontUnifySettings): Promise<OperationResult> {
  ensurePowerPointApi("1.4", "字体统一");

  return runPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    slides.items.forEach((slide) => slide.shapes.load("items/type"));
    await context.sync();

    const counts: AppliedCount = { shapes: 0, groups: 0, tables: 0, skipped: 0 };
    for (const slide of slides.items) {
      await processShapes(context, slide.shapes.items, settings, counts);
    }
    await context.sync();
    return summarize(counts, "全部幻灯片", resolveFontName(settings));
  });
}
