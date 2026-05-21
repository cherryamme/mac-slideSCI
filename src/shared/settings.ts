import { defaultCaptionSettings, defaultLabelSettings } from "../features/labels/labelGenerator";
import { normalizeDpi, normalizePaddingPoints } from "../features/export/pngExportMath";
import { loadJson, saveJson } from "./storage";
import type { CaptionSettings, FontUnifySettings, GeometryMemory, LabelSettings, LayoutSettings, PngExportSettings } from "./types";

export const PNG_EXPORT_STORAGE_KEY = "mac-slideSCI:pngExportSettings";
export const LABEL_STORAGE_KEY = "mac-slideSCI:labelSettings";
export const CAPTION_STORAGE_KEY = "mac-slideSCI:captionSettings";
export const GEOMETRY_STORAGE_KEY = "mac-slideSCI:geometryMemory";
export const LAYOUT_SETTINGS_STORAGE_KEY = "mac-slideSCI:layoutSettings";
export const FONT_UNIFY_STORAGE_KEY = "mac-slideSCI:fontUnifySettings";

export const presetFontNames = [
  "Arial",
  "Aptos",
  "Calibri",
  "Times New Roman",
  "Helvetica Neue",
  "PingFang SC",
  "Songti SC",
  "Heiti SC",
  "Microsoft YaHei",
  "SimSun",
  "Source Han Sans SC",
] as const;

export const defaultPngExportSettings: PngExportSettings = {
  dpi: 300,
  paddingPoints: 2,
};

export const defaultLayoutSettings: LayoutSettings = {
  columns: 2,
  horizontalSpacing: 12,
  verticalSpacing: 12,
  equalizeGrid: true,
};

export const defaultFontUnifySettings: FontUnifySettings = {
  fontName: "Arial",
  customFontName: "",
  unifyAllFormatting: false,
  fontSize: 18,
  bold: false,
  italic: false,
  horizontalAlignment: "left",
};

function clampFontSize(size: number): number {
  if (!Number.isFinite(size)) {
    return defaultFontUnifySettings.fontSize;
  }

  return Math.min(96, Math.max(6, Math.round(size)));
}

export function resolveFontName(settings: FontUnifySettings): string {
  const custom = settings.customFontName?.trim();
  if (custom) {
    return custom;
  }

  return settings.fontName?.trim() || defaultFontUnifySettings.fontName;
}

export function loadPngExportSettings(): PngExportSettings {
  const stored = loadJson<Partial<PngExportSettings>>(PNG_EXPORT_STORAGE_KEY, defaultPngExportSettings);
  return {
    dpi: normalizeDpi(stored.dpi ?? defaultPngExportSettings.dpi),
    paddingPoints: normalizePaddingPoints(stored.paddingPoints ?? defaultPngExportSettings.paddingPoints),
  };
}

export function loadLayoutSettings(): LayoutSettings {
  return loadJson(LAYOUT_SETTINGS_STORAGE_KEY, defaultLayoutSettings);
}

export function saveLayoutSettings(settings: LayoutSettings): void {
  saveJson(LAYOUT_SETTINGS_STORAGE_KEY, settings);
}

export function loadLabelSettings(): LabelSettings {
  return loadJson(LABEL_STORAGE_KEY, defaultLabelSettings);
}

export function loadCaptionSettings(): CaptionSettings {
  return loadJson(CAPTION_STORAGE_KEY, defaultCaptionSettings);
}

export function loadGeometryMemory(): GeometryMemory | null {
  return loadJson<GeometryMemory | null>(GEOMETRY_STORAGE_KEY, null);
}

export function saveGeometryMemory(memory: GeometryMemory | null): void {
  saveJson(GEOMETRY_STORAGE_KEY, memory);
}

export function loadFontUnifySettings(): FontUnifySettings {
  const stored = loadJson<Partial<FontUnifySettings>>(FONT_UNIFY_STORAGE_KEY, defaultFontUnifySettings);
  return {
    fontName: stored.fontName ?? defaultFontUnifySettings.fontName,
    customFontName: stored.customFontName ?? "",
    unifyAllFormatting: Boolean(stored.unifyAllFormatting),
    fontSize: clampFontSize(stored.fontSize ?? defaultFontUnifySettings.fontSize),
    bold: Boolean(stored.bold),
    italic: Boolean(stored.italic),
    horizontalAlignment: stored.horizontalAlignment ?? defaultFontUnifySettings.horizontalAlignment,
  };
}

export function saveFontUnifySettings(settings: FontUnifySettings): void {
  saveJson(FONT_UNIFY_STORAGE_KEY, settings);
}
