/**
 * Excalidraw Engine
 *
 * Manages in-memory Excalidraw scenes — element CRUD, arrow binding,
 * frame grouping, and file export (.excalidraw, SVG, PNG).
 *
 * Zero native dependencies for scene manipulation.
 * Uses `sharp` only for SVG→PNG rasterization.
 */

import crypto from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import {
  AZURE_COLORS,
  type ArrowResult,
  type Arrowhead,
  type BoundElement,
  type DiagramInfo,
  type ElementResult,
  type ExcalidrawArrow,
  type ExcalidrawDiamond,
  type ExcalidrawElement,
  type ExcalidrawElementBase,
  type ExcalidrawElementType,
  type ExcalidrawEllipse,
  type ExcalidrawFileEntry,
  type ExcalidrawFrame,
  type ExcalidrawImage,
  type ExcalidrawRectangle,
  type ExcalidrawScene,
  type ExcalidrawText,
  type FillStyle,
  type PointBinding,
  type StrokeStyle,
} from "./types.js";

// ── Style Defaults (per STYLE_GUIDE.md) ─────────────────────

const DEFAULT_STROKE_COLOR = "#1e1e1e";
const DEFAULT_BG_COLOR = "transparent";
const DEFAULT_FILL_STYLE: FillStyle = "solid";
const DEFAULT_STROKE_WIDTH = 2;
const DEFAULT_STROKE_STYLE: StrokeStyle = "solid";
const DEFAULT_ROUGHNESS = 0; // architect mode — clean lines
const DEFAULT_OPACITY = 100;
const DEFAULT_FONT_SIZE = 16;
const DEFAULT_FONT_FAMILY = 1; // Virgil (hand-drawn) = 1, Helvetica = 2, Cascadia = 3
const DEFAULT_LINE_HEIGHT = 1.25;

const CONTAINER_STROKE_STYLE: StrokeStyle = "dashed";
const CONTAINER_OPACITY = 40;
const CONTAINER_STROKE_WIDTH = 2;

const ARROW_STROKE_WIDTH = 2;
const ARROW_END_HEAD: Arrowhead = "arrow";

const CANVAS_WIDTH = 1056; // ~11in at 96dpi
const CANVAS_HEIGHT = 816; // ~8.5in at 96dpi

// ── Helpers ─────────────────────────────────────────────────

function generateId(): string {
  return crypto.randomBytes(10).toString("hex").slice(0, 20);
}

function generateSeed(): number {
  return crypto.randomInt(1, 2_000_000_000);
}

function expandTilde(filePath: string): string {
  if (filePath.startsWith("~/") || filePath === "~") {
    return path.join(os.homedir(), filePath.slice(2));
  }
  return filePath;
}

function resolveColor(color?: string): string {
  if (!color) return DEFAULT_BG_COLOR;
  const named = AZURE_COLORS[color];
  if (named) return named.hex;
  if (/^#[0-9a-fA-F]{6}$/.test(color)) return color;
  if (/^[0-9a-fA-F]{6}$/.test(color)) return `#${color}`;
  return color;
}

function makeBase(
  type: ExcalidrawElementType,
  x: number,
  y: number,
  width: number,
  height: number,
  overrides?: Partial<ExcalidrawElementBase>,
): ExcalidrawElementBase {
  return {
    id: generateId(),
    type,
    x,
    y,
    width,
    height,
    angle: 0,
    strokeColor: DEFAULT_STROKE_COLOR,
    backgroundColor: DEFAULT_BG_COLOR,
    fillStyle: DEFAULT_FILL_STYLE,
    strokeWidth: DEFAULT_STROKE_WIDTH,
    strokeStyle: DEFAULT_STROKE_STYLE,
    roughness: DEFAULT_ROUGHNESS,
    opacity: DEFAULT_OPACITY,
    seed: generateSeed(),
    version: 1,
    versionNonce: generateSeed(),
    isDeleted: false,
    locked: false,
    groupIds: [],
    frameId: null,
    boundElements: null,
    link: null,
    roundness: { type: 3 },
    ...overrides,
  };
}

// ── Engine ──────────────────────────────────────────────────

export class ExcalidrawEngine {
  private scene: ExcalidrawScene;

  constructor() {
    this.scene = this.createEmptyScene();
  }

  // ── Scene Management ────────────────────────────────────

  private createEmptyScene(): ExcalidrawScene {
    return {
      type: "excalidraw",
      version: 2,
      source: "https://github.com/ragmha/mcp-server-excalidraw",
      elements: [],
      appState: {
        gridSize: null,
        gridStep: 5,
        gridModeEnabled: false,
        viewBackgroundColor: "#ffffff",
      },
      files: {},
    };
  }

  createDiagram(): string {
    this.scene = this.createEmptyScene();
    return "New Excalidraw diagram created";
  }

  saveDiagram(filePath: string): string {
    const absPath = path.resolve(expandTilde(filePath));
    const ext = path.extname(absPath).toLowerCase();
    const finalPath = ext === ".excalidraw" ? absPath : `${absPath}.excalidraw`;
    fs.writeFileSync(finalPath, JSON.stringify(this.scene, null, 2), "utf-8");
    return finalPath;
  }

  async exportDiagram(filePath: string): Promise<string> {
    const absPath = path.resolve(expandTilde(filePath));
    const ext = path.extname(absPath).toLowerCase();

    if (ext === ".excalidraw") {
      return this.saveDiagram(absPath);
    }

    const svg = this.generateSvg();

    if (ext === ".svg") {
      fs.writeFileSync(absPath, svg, "utf-8");
      return absPath;
    }

    if (ext === ".png" || ext === ".jpg" || ext === ".jpeg") {
      const sharp = (await import("sharp")).default;
      const buf = Buffer.from(svg, "utf-8");
      const pipeline = sharp(buf);
      if (ext === ".png") {
        await pipeline.png().toFile(absPath);
      } else {
        await pipeline.jpeg().toFile(absPath);
      }
      return absPath;
    }

    throw new Error(`Unsupported export format: ${ext}. Use .excalidraw, .svg, .png, or .jpg`);
  }

  getDiagramInfo(): DiagramInfo {
    const elements = this.scene.elements.filter((e) => !e.isDeleted);
    const arrows = elements.filter((e) => e.type === "arrow");
    const nonArrows = elements.filter((e) => e.type !== "arrow");

    let bb: DiagramInfo["bounding_box"] = null;
    if (nonArrows.length > 0) {
      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      for (const el of nonArrows) {
        minX = Math.min(minX, el.x);
        minY = Math.min(minY, el.y);
        maxX = Math.max(maxX, el.x + el.width);
        maxY = Math.max(maxY, el.y + el.height);
      }
      bb = { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
    }

    return {
      element_count: elements.length,
      arrow_count: arrows.length,
      bounding_box: bb,
    };
  }

  // ── Element Operations ──────────────────────────────────

  addElement(
    elementType: "rectangle" | "ellipse" | "diamond",
    x: number,
    y: number,
    width: number,
    height: number,
    text?: string,
    strokeColor?: string,
    backgroundColor?: string,
    fillStyle?: FillStyle,
    strokeStyle?: StrokeStyle,
    strokeWidth?: number,
    opacity?: number,
  ): ElementResult {
    const base = makeBase(elementType, x, y, width, height, {
      strokeColor: resolveColor(strokeColor) || DEFAULT_STROKE_COLOR,
      backgroundColor: resolveColor(backgroundColor),
      fillStyle: fillStyle ?? DEFAULT_FILL_STYLE,
      strokeStyle: strokeStyle ?? DEFAULT_STROKE_STYLE,
      strokeWidth: strokeWidth ?? DEFAULT_STROKE_WIDTH,
      opacity: opacity ?? DEFAULT_OPACITY,
    });

    const element = base as ExcalidrawRectangle | ExcalidrawEllipse | ExcalidrawDiamond;
    this.scene.elements.push(element);

    if (text) {
      this.addBoundText(element, text);
    }

    return this.toElementResult(element, text);
  }

  addAzureIcon(
    serviceKey: string,
    svgDataUrl: string,
    x: number,
    y: number,
    width: number = 48,
    height: number = 48,
    label?: string,
  ): ElementResult {
    const fileId = generateId();

    // Register the SVG in the files map
    this.scene.files[fileId] = {
      mimeType: "image/svg+xml",
      id: fileId,
      dataURL: svgDataUrl,
      created: Date.now(),
      lastRetrieved: Date.now(),
    };

    const base = makeBase("image", x, y, width, height, {
      strokeColor: "transparent",
      backgroundColor: "transparent",
      roundness: null,
    });

    const imageElement: ExcalidrawImage = {
      ...base,
      type: "image",
      fileId,
      status: "saved",
      scale: [1, 1],
    };

    this.scene.elements.push(imageElement);

    // Add label below the icon
    if (label) {
      const labelY = y + height + 8;
      this.addText(x + width / 2, labelY, label, 14, "center");
    }

    return {
      id: imageElement.id,
      type: "image",
      x,
      y,
      width,
      height,
      azure_service: serviceKey,
      text: label,
    };
  }

  addText(
    x: number,
    y: number,
    text: string,
    fontSize?: number,
    textAlign?: "left" | "center" | "right",
  ): ElementResult {
    const size = fontSize ?? DEFAULT_FONT_SIZE;
    const lines = text.split("\n");
    const estimatedWidth = Math.max(...lines.map((l) => l.length)) * size * 0.6;
    const estimatedHeight = lines.length * size * DEFAULT_LINE_HEIGHT;

    const base = makeBase("text", x, y, estimatedWidth, estimatedHeight, {
      strokeColor: DEFAULT_STROKE_COLOR,
      backgroundColor: "transparent",
      roundness: null,
    });

    const textElement: ExcalidrawText = {
      ...base,
      type: "text",
      text,
      fontSize: size,
      fontFamily: DEFAULT_FONT_FAMILY,
      textAlign: textAlign ?? "left",
      verticalAlign: "top",
      containerId: null,
      originalText: text,
      autoResize: true,
      lineHeight: DEFAULT_LINE_HEIGHT,
    };

    this.scene.elements.push(textElement);
    return this.toElementResult(textElement, text);
  }

  modifyElement(
    elementId: string,
    updates: {
      x?: number;
      y?: number;
      width?: number;
      height?: number;
      text?: string;
      strokeColor?: string;
      backgroundColor?: string;
      fillStyle?: FillStyle;
      strokeStyle?: StrokeStyle;
      strokeWidth?: number;
      opacity?: number;
    },
  ): ElementResult {
    const element = this.findElement(elementId);

    if (updates.x !== undefined) element.x = updates.x;
    if (updates.y !== undefined) element.y = updates.y;
    if (updates.width !== undefined) element.width = updates.width;
    if (updates.height !== undefined) element.height = updates.height;
    if (updates.strokeColor) element.strokeColor = resolveColor(updates.strokeColor);
    if (updates.backgroundColor) element.backgroundColor = resolveColor(updates.backgroundColor);
    if (updates.fillStyle) element.fillStyle = updates.fillStyle;
    if (updates.strokeStyle) element.strokeStyle = updates.strokeStyle;
    if (updates.strokeWidth !== undefined) element.strokeWidth = updates.strokeWidth;
    if (updates.opacity !== undefined) element.opacity = updates.opacity;

    element.version += 1;
    element.versionNonce = generateSeed();

    // Update bound text if applicable
    if (updates.text !== undefined && element.type !== "text") {
      const boundTextId = element.boundElements?.find((b) => b.type === "text")?.id;
      if (boundTextId) {
        const textEl = this.scene.elements.find((e) => e.id === boundTextId) as ExcalidrawText | undefined;
        if (textEl) {
          textEl.text = updates.text;
          textEl.originalText = updates.text;
          textEl.version += 1;
          textEl.versionNonce = generateSeed();
        }
      }
    }

    if (updates.text !== undefined && element.type === "text") {
      (element as ExcalidrawText).text = updates.text;
      (element as ExcalidrawText).originalText = updates.text;
    }

    return this.toElementResult(element, updates.text);
  }

  removeElement(elementId: string): string {
    const element = this.findElement(elementId);
    element.isDeleted = true;
    element.version += 1;

    // Also delete bound text
    if (element.boundElements) {
      for (const bound of element.boundElements) {
        const boundEl = this.scene.elements.find((e) => e.id === bound.id);
        if (boundEl) {
          boundEl.isDeleted = true;
          boundEl.version += 1;
        }
      }
    }

    // Remove bindings from arrows that reference this element
    for (const el of this.scene.elements) {
      if (el.type === "arrow") {
        const arrow = el as ExcalidrawArrow;
        if (arrow.startBinding?.elementId === elementId) arrow.startBinding = null;
        if (arrow.endBinding?.elementId === elementId) arrow.endBinding = null;
      }
    }

    return `Removed element ${elementId}`;
  }

  listElements(): ElementResult[] {
    return this.scene.elements
      .filter((e) => !e.isDeleted && e.type !== "text")
      .map((e) => {
        const boundText = this.getBoundText(e);
        return this.toElementResult(e, boundText);
      });
  }

  // ── Arrow Operations ────────────────────────────────────

  addArrow(
    fromId: string,
    toId: string,
    label?: string,
    strokeColor?: string,
    strokeStyle?: StrokeStyle,
    startArrowhead?: Arrowhead,
    endArrowhead?: Arrowhead,
  ): ArrowResult {
    const fromEl = this.findElement(fromId);
    const toEl = this.findElement(toId);

    // Calculate arrow start/end points (center of elements)
    const fromCenterX = fromEl.x + fromEl.width / 2;
    const fromCenterY = fromEl.y + fromEl.height / 2;
    const toCenterX = toEl.x + toEl.width / 2;
    const toCenterY = toEl.y + toEl.height / 2;

    const dx = toCenterX - fromCenterX;
    const dy = toCenterY - fromCenterY;

    // Arrow origin is always the starting point; points are relative to it
    const base = makeBase("arrow", fromCenterX, fromCenterY, Math.abs(dx), Math.abs(dy), {
      strokeColor: resolveColor(strokeColor) || DEFAULT_STROKE_COLOR,
      strokeWidth: ARROW_STROKE_WIDTH,
      strokeStyle: strokeStyle ?? DEFAULT_STROKE_STYLE,
      backgroundColor: "transparent",
      roundness: { type: 2 },
    });

    const startBinding: PointBinding = {
      elementId: fromId,
      focus: 0,
      gap: 4,
      fixedPoint: null,
    };

    const endBinding: PointBinding = {
      elementId: toId,
      focus: 0,
      gap: 4,
      fixedPoint: null,
    };

    const arrow: ExcalidrawArrow = {
      ...base,
      type: "arrow",
      points: [
        [0, 0],
        [dx, dy],
      ],
      startBinding,
      endBinding,
      startArrowhead: startArrowhead ?? null,
      endArrowhead: endArrowhead ?? ARROW_END_HEAD,
      elbowed: false,
    };

    this.scene.elements.push(arrow);

    // Update bound elements on source and target
    this.addBoundElementRef(fromEl, arrow.id, "arrow");
    this.addBoundElementRef(toEl, arrow.id, "arrow");

    // Add label if provided
    if (label) {
      const midX = fromCenterX + dx / 2;
      const midY = fromCenterY + dy / 2 - 12;
      this.addText(midX, midY, label, 14, "center");
    }

    return {
      id: arrow.id,
      from_id: fromId,
      to_id: toId,
      label,
    };
  }

  removeArrow(arrowId: string): string {
    const arrow = this.findElement(arrowId);
    if (arrow.type !== "arrow") {
      throw new Error(`Element ${arrowId} is not an arrow (type: ${arrow.type})`);
    }

    const arrowEl = arrow as ExcalidrawArrow;

    // Remove binding references from connected elements
    if (arrowEl.startBinding) {
      this.removeBoundElementRef(arrowEl.startBinding.elementId, arrowId);
    }
    if (arrowEl.endBinding) {
      this.removeBoundElementRef(arrowEl.endBinding.elementId, arrowId);
    }

    arrow.isDeleted = true;
    arrow.version += 1;

    return `Removed arrow ${arrowId}`;
  }

  // ── Frame Operations ────────────────────────────────────

  addFrame(
    x: number,
    y: number,
    width: number,
    height: number,
    name?: string,
  ): ElementResult {
    const base = makeBase("frame", x, y, width, height, {
      strokeColor: "#bbb",
      backgroundColor: "transparent",
      roundness: null,
    });

    const frame: ExcalidrawFrame = {
      ...base,
      type: "frame",
      name: name ?? null,
    };

    this.scene.elements.push(frame);
    return this.toElementResult(frame, name);
  }

  addContainer(
    x: number,
    y: number,
    width: number,
    height: number,
    label?: string,
    backgroundColor?: string,
    opacity?: number,
  ): ElementResult {
    const bgColor = resolveColor(backgroundColor) || "#e6f3ff";

    const base = makeBase("rectangle", x, y, width, height, {
      strokeColor: resolveColor(backgroundColor) || "#0078D7",
      backgroundColor: bgColor,
      fillStyle: "solid",
      strokeStyle: CONTAINER_STROKE_STYLE,
      strokeWidth: CONTAINER_STROKE_WIDTH,
      opacity: opacity ?? CONTAINER_OPACITY,
      roundness: { type: 3 },
    });

    const container = base as ExcalidrawRectangle;
    this.scene.elements.push(container);

    if (label) {
      this.addText(x + 10, y + 8, label, 14, "left");
    }

    return this.toElementResult(container, label);
  }

  // ── Azure Service Discovery ─────────────────────────────

  listAzureServices(iconMap: Record<string, unknown>): string[] {
    return Object.keys(iconMap).sort();
  }

  // ── SVG Generation ──────────────────────────────────────

  private generateSvg(): string {
    const elements = this.scene.elements.filter((e) => !e.isDeleted);

    // Calculate bounding box
    let minX = 0, minY = 0, maxX = CANVAS_WIDTH, maxY = CANVAS_HEIGHT;
    if (elements.length > 0) {
      minX = Math.min(...elements.map((e) => e.x)) - 20;
      minY = Math.min(...elements.map((e) => e.y)) - 20;
      maxX = Math.max(...elements.map((e) => e.x + e.width)) + 20;
      maxY = Math.max(...elements.map((e) => e.y + e.height)) + 20;
    }

    const viewWidth = maxX - minX;
    const viewHeight = maxY - minY;

    // Collect unique arrow IDs for marker defs
    const arrowIds = elements
      .filter((e) => e.type === "arrow" && (e as ExcalidrawArrow).endArrowhead)
      .map((e) => e.id);

    const parts: string[] = [
      `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="${minX} ${minY} ${viewWidth} ${viewHeight}" width="${viewWidth}" height="${viewHeight}">`,
      `<rect x="${minX}" y="${minY}" width="${viewWidth}" height="${viewHeight}" fill="${this.scene.appState.viewBackgroundColor}" />`,
    ];

    // Emit a single <defs> block with all arrowhead markers
    if (arrowIds.length > 0) {
      parts.push("<defs>");
      for (const id of arrowIds) {
        const el = elements.find((e) => e.id === id)!;
        parts.push(
          `<marker id="ah-${id}" markerWidth="10" markerHeight="7" refX="10" refY="3.5" orient="auto"><polygon points="0 0, 10 3.5, 0 7" fill="${el.strokeColor}" /></marker>`,
        );
      }
      parts.push("</defs>");
    }

    for (const el of elements) {
      parts.push(this.elementToSvg(el));
    }

    parts.push("</svg>");
    return parts.join("\n");
  }

  private elementToSvg(el: ExcalidrawElement): string {
    const opacity = el.opacity / 100;
    const dashArray = el.strokeStyle === "dashed" ? 'stroke-dasharray="8 4"' : el.strokeStyle === "dotted" ? 'stroke-dasharray="2 4"' : "";

    switch (el.type) {
      case "rectangle":
        return `<rect x="${el.x}" y="${el.y}" width="${el.width}" height="${el.height}" fill="${el.backgroundColor}" stroke="${el.strokeColor}" stroke-width="${el.strokeWidth}" ${dashArray} opacity="${opacity}" rx="8" />`;

      case "ellipse":
        return `<ellipse cx="${el.x + el.width / 2}" cy="${el.y + el.height / 2}" rx="${el.width / 2}" ry="${el.height / 2}" fill="${el.backgroundColor}" stroke="${el.strokeColor}" stroke-width="${el.strokeWidth}" ${dashArray} opacity="${opacity}" />`;

      case "diamond": {
        const cx = el.x + el.width / 2;
        const cy = el.y + el.height / 2;
        const points = `${cx},${el.y} ${el.x + el.width},${cy} ${cx},${el.y + el.height} ${el.x},${cy}`;
        return `<polygon points="${points}" fill="${el.backgroundColor}" stroke="${el.strokeColor}" stroke-width="${el.strokeWidth}" ${dashArray} opacity="${opacity}" />`;
      }

      case "text": {
        const textEl = el as ExcalidrawText;
        const anchor = textEl.textAlign === "center" ? "middle" : textEl.textAlign === "right" ? "end" : "start";
        const textX = textEl.textAlign === "center" ? el.x + el.width / 2 : el.x;
        return `<text x="${textX}" y="${el.y + textEl.fontSize}" font-size="${textEl.fontSize}" fill="${el.strokeColor}" text-anchor="${anchor}" opacity="${opacity}">${this.escapeXml(textEl.text)}</text>`;
      }

      case "arrow": {
        const arrowEl = el as ExcalidrawArrow;
        const pointsStr = arrowEl.points
          .map(([px, py]) => `${el.x + px},${el.y + py}`)
          .join(" ");
        const marker = arrowEl.endArrowhead ? ` marker-end="url(#ah-${el.id})"` : "";
        return `<polyline points="${pointsStr}" fill="none" stroke="${el.strokeColor}" stroke-width="${el.strokeWidth}" ${dashArray}${marker} opacity="${opacity}" />`;
      }

      case "image": {
        const imgEl = el as ExcalidrawImage;
        const fileEntry = this.scene.files[imgEl.fileId];
        if (fileEntry) {
          return `<image x="${el.x}" y="${el.y}" width="${el.width}" height="${el.height}" href="${fileEntry.dataURL}" opacity="${opacity}" />`;
        }
        return `<rect x="${el.x}" y="${el.y}" width="${el.width}" height="${el.height}" fill="#eee" stroke="#ccc" />`;
      }

      case "frame":
        return `<rect x="${el.x}" y="${el.y}" width="${el.width}" height="${el.height}" fill="none" stroke="#bbb" stroke-width="1" stroke-dasharray="4 4" opacity="${opacity}" />`;

      case "line":
        return "";

      default:
        return "";
    }
  }

  private escapeXml(text: string): string {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // ── Internal Helpers ────────────────────────────────────

  private findElement(id: string): ExcalidrawElement {
    const el = this.scene.elements.find((e) => e.id === id && !e.isDeleted);
    if (!el) throw new Error(`Element not found: ${id}`);
    return el;
  }

  private getBoundText(element: ExcalidrawElement): string | undefined {
    const textBound = element.boundElements?.find((b) => b.type === "text");
    if (!textBound) return undefined;
    const textEl = this.scene.elements.find((e) => e.id === textBound.id && !e.isDeleted) as ExcalidrawText | undefined;
    return textEl?.text;
  }

  private addBoundText(
    parent: ExcalidrawElement,
    text: string,
    fontSize?: number,
  ): ExcalidrawText {
    const size = fontSize ?? DEFAULT_FONT_SIZE;
    const lines = text.split("\n");
    const estimatedWidth = Math.max(...lines.map((l) => l.length)) * size * 0.6;
    const estimatedHeight = lines.length * size * DEFAULT_LINE_HEIGHT;

    const base = makeBase(
      "text",
      parent.x + (parent.width - estimatedWidth) / 2,
      parent.y + (parent.height - estimatedHeight) / 2,
      estimatedWidth,
      estimatedHeight,
      {
        strokeColor: DEFAULT_STROKE_COLOR,
        backgroundColor: "transparent",
        roundness: null,
      },
    );

    const textElement: ExcalidrawText = {
      ...base,
      type: "text",
      text,
      fontSize: size,
      fontFamily: DEFAULT_FONT_FAMILY,
      textAlign: "center",
      verticalAlign: "middle",
      containerId: parent.id,
      originalText: text,
      autoResize: true,
      lineHeight: DEFAULT_LINE_HEIGHT,
    };

    this.scene.elements.push(textElement);
    this.addBoundElementRef(parent, textElement.id, "text");

    return textElement;
  }

  private addBoundElementRef(element: ExcalidrawElement, refId: string, refType: BoundElement["type"]): void {
    if (!element.boundElements) {
      element.boundElements = [];
    }
    element.boundElements.push({ id: refId, type: refType });
  }

  private removeBoundElementRef(elementId: string, refId: string): void {
    const el = this.scene.elements.find((e) => e.id === elementId && !e.isDeleted);
    if (el?.boundElements) {
      el.boundElements = el.boundElements.filter((b) => b.id !== refId);
    }
  }

  private toElementResult(element: ExcalidrawElement, text?: string): ElementResult {
    return {
      id: element.id,
      type: element.type,
      x: element.x,
      y: element.y,
      width: element.width,
      height: element.height,
      text,
    };
  }
}
