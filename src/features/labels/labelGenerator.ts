import type { CaptionSettings, LabelSettings } from "../../shared/types";

export const defaultLabelSettings: LabelSettings = {
  prefix: "",
  startIndex: 0,
  fontSize: 18,
  bold: true,
  color: "#111111",
  backgroundColor: "transparent",
  offsetX: 6,
  offsetY: 6,
  width: 28,
  height: 24,
};

export const defaultCaptionSettings: CaptionSettings = {
  text: "Figure caption",
  fontSize: 12,
  color: "#222222",
  gap: 6,
  height: 28,
};

export function getPanelLabel(index: number, settings: LabelSettings): string {
  const value = settings.startIndex + index;
  const letter = String.fromCharCode("A".charCodeAt(0) + (value % 26));
  const suffix = value >= 26 ? Math.floor(value / 26) + 1 : "";
  return `${settings.prefix}${letter}${suffix}`;
}
