// Types for the Excalidraw MCP Server

// ── Excalidraw Element Types ────────────────────────────────

export type ExcalidrawElementType =
  | "rectangle"
  | "ellipse"
  | "diamond"
  | "arrow"
  | "line"
  | "text"
  | "freedraw"
  | "image"
  | "frame";

export type FillStyle = "solid" | "hachure" | "cross-hatch";
export type StrokeStyle = "solid" | "dashed" | "dotted";
export type Arrowhead = "arrow" | "bar" | "dot" | "triangle" | null;
export type TextAlign = "left" | "center" | "right";
export type VerticalAlign = "top" | "middle" | "bottom";
export type RoundnessType = { type: 1 | 2 | 3 };

export interface BoundElement {
  id: string;
  type: "arrow" | "text";
}

export interface PointBinding {
  elementId: string;
  focus: number;
  gap: number;
  fixedPoint: [number, number] | null;
}

/** Base properties shared by all Excalidraw elements. */
export interface ExcalidrawElementBase {
  id: string;
  type: ExcalidrawElementType;
  x: number;
  y: number;
  width: number;
  height: number;
  angle: number;
  strokeColor: string;
  backgroundColor: string;
  fillStyle: FillStyle;
  strokeWidth: number;
  strokeStyle: StrokeStyle;
  roughness: number;
  opacity: number;
  seed: number;
  version: number;
  versionNonce: number;
  updated?: number;
  index?: string;
  isDeleted: boolean;
  locked: boolean;
  groupIds: string[];
  frameId: string | null;
  boundElements: BoundElement[] | null;
  link: string | null;
  roundness: RoundnessType | null;
}

export interface ExcalidrawRectangle extends ExcalidrawElementBase {
  type: "rectangle";
}

export interface ExcalidrawEllipse extends ExcalidrawElementBase {
  type: "ellipse";
}

export interface ExcalidrawDiamond extends ExcalidrawElementBase {
  type: "diamond";
}

export interface ExcalidrawText extends ExcalidrawElementBase {
  type: "text";
  text: string;
  fontSize: number;
  fontFamily: number;
  textAlign: TextAlign;
  verticalAlign: VerticalAlign;
  containerId: string | null;
  originalText: string;
  autoResize: boolean;
  lineHeight: number;
}

export interface ExcalidrawArrow extends ExcalidrawElementBase {
  type: "arrow";
  points: [number, number][];
  startBinding: PointBinding | null;
  endBinding: PointBinding | null;
  startArrowhead: Arrowhead;
  endArrowhead: Arrowhead;
  elbowed: boolean;
}

export interface ExcalidrawLine extends ExcalidrawElementBase {
  type: "line";
  points: [number, number][];
  startBinding: PointBinding | null;
  endBinding: PointBinding | null;
  startArrowhead: null;
  endArrowhead: null;
}

export interface ExcalidrawImage extends ExcalidrawElementBase {
  type: "image";
  fileId: string;
  status: "pending" | "saved" | "error";
  scale: [number, number];
}

export interface ExcalidrawFrame extends ExcalidrawElementBase {
  type: "frame";
  name: string | null;
}

export type ExcalidrawElement =
  | ExcalidrawRectangle
  | ExcalidrawEllipse
  | ExcalidrawDiamond
  | ExcalidrawText
  | ExcalidrawArrow
  | ExcalidrawLine
  | ExcalidrawImage
  | ExcalidrawFrame;

// ── Excalidraw File (for embedded images) ───────────────────

export interface ExcalidrawFileEntry {
  mimeType: string;
  id: string;
  dataURL: string;
  created: number;
  lastRetrieved: number;
}

// ── Excalidraw Scene ────────────────────────────────────────

export interface ExcalidrawAppState {
  gridSize: number | null;
  gridStep: number;
  gridModeEnabled: boolean;
  viewBackgroundColor: string;
}

export interface ExcalidrawScene {
  type: "excalidraw";
  version: number;
  source: string;
  elements: ExcalidrawElement[];
  appState: ExcalidrawAppState;
  files: Record<string, ExcalidrawFileEntry>;
}

// ── Tool Result Types ───────────────────────────────────────

export interface ElementResult {
  id: string;
  type: ExcalidrawElementType;
  x: number;
  y: number;
  width: number;
  height: number;
  text?: string;
  azure_service?: string;
}

export interface ArrowResult {
  id: string;
  from_id: string;
  to_id: string;
  label?: string;
}

export interface DiagramInfo {
  element_count: number;
  arrow_count: number;
  bounding_box: { x: number; y: number; width: number; height: number } | null;
}

// ── Azure Service Icon Map ──────────────────────────────────

export interface AzureIconEntry {
  category: string;
  svgDataUrl: string;
}

export type AzureIconMap = Record<string, AzureIconEntry>;

// ── Style Constants ─────────────────────────────────────────

export interface AzureColor {
  hex: string;
  r: number;
  g: number;
  b: number;
}

export const AZURE_COLORS: Record<string, AzureColor> = {
  azure_blue: { hex: "#0078D7", r: 0, g: 120, b: 215 },
  dark_blue: { hex: "#004E98", r: 0, g: 78, b: 152 },
  teal: { hex: "#00B294", r: 0, g: 178, b: 148 },
  orange: { hex: "#FF8C00", r: 255, g: 140, b: 0 },
  purple: { hex: "#8764B8", r: 135, g: 100, b: 184 },
  green: { hex: "#7AB800", r: 122, g: 184, b: 0 },
  red: { hex: "#E81123", r: 232, g: 17, b: 35 },
};
