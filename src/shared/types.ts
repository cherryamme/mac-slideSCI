export interface OfficeReadyInfo {
  host?: Office.HostType;
  platform?: Office.PlatformType;
}

export interface RequirementStatus {
  name: string;
  version: string;
  supported: boolean;
}

export interface ShapeBox {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface LayoutSpacing {
  horizontal: number;
  vertical: number;
}

export interface GridLayoutOptions {
  columns: number;
  spacing: LayoutSpacing;
  equalizeSize: boolean;
}

export interface LayoutSettings {
  columns: number;
  horizontalSpacing: number;
  verticalSpacing: number;
  equalizeGrid: boolean;
}

export type AlignmentMode = "left" | "center" | "right" | "top" | "middle" | "bottom";
export type DistributionMode = "horizontal" | "vertical";
export type EqualizeMode = "width" | "height" | "both";
export type GeometryApplyMode = "position" | "size" | "all";

export interface GeometryMemory {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface LabelSettings {
  prefix: string;
  startIndex: number;
  fontSize: number;
  bold: boolean;
  color: string;
  backgroundColor: string;
  offsetX: number;
  offsetY: number;
  width: number;
  height: number;
}

export interface CaptionSettings {
  text: string;
  fontSize: number;
  color: string;
  gap: number;
  height: number;
}

export type ExportProgressPhase = "selection" | "pdf" | "render" | "crop" | "zip" | "download";

export interface PngExportSettings {
  dpi: number;
  paddingPoints: number;
}

export interface ExportProgress {
  phase: ExportProgressPhase;
  current: number;
  total: number;
  message: string;
}

export interface RenderedPngPage {
  pageNumber: number;
  blob: Blob;
  width: number;
  height: number;
}

export interface RenderedPngSelection {
  blob: Blob;
  width: number;
  height: number;
  dpi: number;
  selectedShapeCount: number;
}

export interface OperationResult {
  count: number;
  message: string;
}

export interface PngExportResult extends OperationResult {
  fileName: string;
}

export type FontHorizontalAlignment = "left" | "center" | "right" | "justify";

export interface FontUnifySettings {
  fontName: string;
  customFontName: string;
  unifyAllFormatting: boolean;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  horizontalAlignment: FontHorizontalAlignment;
}

export type FontUnifyScope = "selection" | "slide" | "presentation";

export type ComboElementType = "textBox" | "geometricShape" | "line" | "group" | "unsupported";

export interface ComboElement {
  id: string;
  type: ComboElementType;
  relativeLeft: number;
  relativeTop: number;
  width: number;
  height: number;
  text?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  fillColor?: string;
  lineColor?: string;
  shapeType?: string;
  connectorType?: string;
  children?: ComboElement[];
}

export interface SavedCombo {
  id: string;
  name: string;
  createdAt: number;
  updatedAt: number;
  width: number;
  height: number;
  thumbnailBase64: string;
  elements: ComboElement[];
  unsupportedCount: number;
}
