// Types for the Visio MCP Server

export interface ShapeResult {
  id: number;
  name: string;
  text: string;
  x: number;
  y: number;
  width?: number;
  height?: number;
  azure_service?: string;
  stencil?: string;
  master?: string;
}

export interface ConnectorResult {
  id: number;
  name: string;
  from_id: number;
  to_id: number;
  label: string;
}

export interface PageResult {
  index: number;
  name: string;
}

export interface DocumentResult {
  index: number;
  name: string;
  path: string;
  pages: number;
}

export interface StencilMapEntry {
  stencil: string;
  master: string;
}

export type StencilMap = Record<string, StencilMapEntry>;

export interface AzureColor {
  r: number;
  g: number;
  b: number;
}

export const AZURE_COLORS: Record<string, AzureColor> = {
  azure_blue: { r: 0, g: 120, b: 215 },
  dark_blue: { r: 0, g: 78, b: 152 },
  teal: { r: 0, g: 178, b: 148 },
  orange: { r: 255, g: 140, b: 0 },
  purple: { r: 135, g: 100, b: 184 },
  green: { r: 122, g: 184, b: 0 },
  red: { r: 232, g: 17, b: 35 },
};
