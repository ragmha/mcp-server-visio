#!/usr/bin/env node

/**
 * Excalidraw MCP Server
 *
 * Exposes Excalidraw diagram operations as MCP tools via stdio transport.
 * Run with: npx mcp-server-excalidraw
 *
 * Designed for GitHub Copilot CLI and VS Code Agent Mode.
 * All tools follow STYLE_GUIDE.md conventions automatically.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { AZURE_ICONS } from "./azure-icons.js";
import { ExcalidrawEngine } from "./excalidraw-engine.js";

const server = new McpServer({
  name: "Excalidraw Diagram Server",
  version: "1.0.0",
});

const engine = new ExcalidrawEngine();

// ── Document Management ──────────────────────────────────────

server.tool(
  "create_diagram",
  `Create a new Excalidraw diagram. Initializes an empty canvas ready for elements.`,
  {},
  async () => {
    try {
      const msg = engine.createDiagram();
      return { content: [{ type: "text", text: msg }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "save_diagram",
  "Save the current diagram to an .excalidraw file. Can be opened in excalidraw.com or the Excalidraw VS Code extension.",
  {
    file_path: z
      .string()
      .describe('Full path to save (e.g. "~/diagrams/arch.excalidraw"). Extension .excalidraw is added if missing.'),
  },
  async ({ file_path }) => {
    try {
      const path = engine.saveDiagram(file_path);
      return { content: [{ type: "text", text: `Saved to: ${path}` }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "export_diagram",
  'Export the current diagram as an image file. Supports .excalidraw, .svg, .png, and .jpg formats. The format is determined by the file extension.',
  {
    file_path: z
      .string()
      .describe('Output path with extension (e.g. "output.svg", "arch.png", "diagram.excalidraw").'),
  },
  async ({ file_path }) => {
    try {
      const path = await engine.exportDiagram(file_path);
      return { content: [{ type: "text", text: `Exported to: ${path}` }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "get_diagram_info",
  "Get summary information about the current diagram: element count, arrow count, and bounding box.",
  {},
  async () => {
    try {
      const info = engine.getDiagramInfo();
      return {
        content: [{ type: "text", text: JSON.stringify(info, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Element Operations ───────────────────────────────────────

server.tool(
  "add_element",
  `Add a shape element to the diagram.
Shapes are styled with clean lines (roughness=0), solid fills.
Use add_azure_icon for Azure service icons — it embeds real SVG icons.`,
  {
    element_type: z
      .enum(["rectangle", "ellipse", "diamond"])
      .describe("Type of shape to add."),
    x: z.number().describe("X position in pixels from left edge."),
    y: z.number().describe("Y position in pixels from top edge."),
    width: z.number().describe("Width in pixels."),
    height: z.number().describe("Height in pixels."),
    text: z.string().optional().describe("Label text to display inside the shape."),
    stroke_color: z
      .string()
      .optional()
      .describe('Stroke color as hex (e.g. "#0078D7") or named: azure_blue, dark_blue, teal, orange, purple, green, red.'),
    background_color: z
      .string()
      .optional()
      .describe("Fill color (hex or named Azure color). Use 'transparent' for no fill."),
    fill_style: z
      .enum(["solid", "hachure", "cross-hatch"])
      .optional()
      .describe("Fill pattern style. Defaults to solid."),
    stroke_style: z
      .enum(["solid", "dashed", "dotted"])
      .optional()
      .describe("Stroke line style. Defaults to solid."),
    stroke_width: z.number().optional().describe("Stroke width in pixels. Defaults to 2."),
    opacity: z.number().min(0).max(100).optional().describe("Opacity from 0 (invisible) to 100 (fully opaque). Defaults to 100."),
  },
  async ({ element_type, x, y, width, height, text, stroke_color, background_color, fill_style, stroke_style, stroke_width, opacity }) => {
    try {
      const result = engine.addElement(
        element_type,
        x,
        y,
        width,
        height,
        text,
        stroke_color,
        background_color,
        fill_style,
        stroke_style,
        stroke_width,
        opacity,
      );
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "add_azure_icon",
  `Add an Azure service icon to the diagram.
Embeds the official Azure SVG icon as an image element.
Use list_azure_services to see all available service keys.`,
  {
    service: z
      .string()
      .describe('Azure service key (e.g. "azure/front-door", "azure/sql-database", "azure/kubernetes-services"). Use list_azure_services to see all available keys.'),
    x: z.number().describe("X position in pixels from left edge."),
    y: z.number().describe("Y position in pixels from top edge."),
    label: z.string().optional().describe("Optional text label below the icon (defaults to the service name)."),
    width: z.number().optional().describe("Icon width in pixels (default 48)."),
    height: z.number().optional().describe("Icon height in pixels (default 48)."),
  },
  async ({ service, x, y, label, width, height }) => {
    try {
      const iconEntry = Object.hasOwn(AZURE_ICONS, service) ? AZURE_ICONS[service] : undefined;
      if (!iconEntry) {
        const available = Object.keys(AZURE_ICONS).sort();
        throw new Error(
          `Unknown Azure service: "${service}". Use list_azure_services to see available keys. Closest matches: ${available.filter((k) => k.includes(service.split("/").pop() ?? "")).join(", ") || "none"}`,
        );
      }

      const displayLabel = label ?? service.split("/").pop()?.replace(/-/g, " ") ?? service;
      const result = engine.addAzureIcon(
        service,
        iconEntry.svgDataUrl,
        x,
        y,
        width ?? 48,
        height ?? 48,
        displayLabel,
      );
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "add_text",
  "Add a floating text label at the given position.",
  {
    x: z.number().describe("X position in pixels."),
    y: z.number().describe("Y position in pixels."),
    text: z.string().describe("The text to display."),
    font_size: z.number().optional().describe("Font size in pixels (default 16)."),
    text_align: z
      .enum(["left", "center", "right"])
      .optional()
      .describe("Text alignment. Defaults to left."),
  },
  async ({ x, y, text, font_size, text_align }) => {
    try {
      const result = engine.addText(x, y, text, font_size, text_align);
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "modify_element",
  "Modify properties of an existing element. Only provided properties are changed; others keep their current values.",
  {
    element_id: z.string().describe("The element's unique ID (from add_element, add_azure_icon, or list_elements)."),
    x: z.number().optional().describe("New X position (or omit to keep current)."),
    y: z.number().optional().describe("New Y position (or omit to keep current)."),
    width: z.number().optional().describe("New width (or omit to keep current)."),
    height: z.number().optional().describe("New height (or omit to keep current)."),
    text: z.string().optional().describe("New text label (or omit to keep current)."),
    stroke_color: z.string().optional().describe("New stroke color (hex or named Azure color)."),
    background_color: z.string().optional().describe("New fill color (hex or named Azure color)."),
    fill_style: z.enum(["solid", "hachure", "cross-hatch"]).optional().describe("New fill pattern style."),
    stroke_style: z.enum(["solid", "dashed", "dotted"]).optional().describe("New stroke line style."),
    stroke_width: z.number().optional().describe("New stroke width."),
    opacity: z.number().min(0).max(100).optional().describe("New opacity (0-100)."),
  },
  async ({ element_id, x, y, width, height, text, stroke_color, background_color, fill_style, stroke_style, stroke_width, opacity }) => {
    try {
      const result = engine.modifyElement(element_id, {
        x,
        y,
        width,
        height,
        text,
        strokeColor: stroke_color,
        backgroundColor: background_color,
        fillStyle: fill_style,
        strokeStyle: stroke_style,
        strokeWidth: stroke_width,
        opacity,
      });
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "remove_element",
  "Remove an element from the diagram by its ID.",
  {
    element_id: z.string().describe("The element's unique ID."),
  },
  async ({ element_id }) => {
    try {
      const msg = engine.removeElement(element_id);
      return { content: [{ type: "text", text: msg }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "list_elements",
  "List all elements on the diagram. Returns JSON array with each element's id, type, position, size, and text.",
  {},
  async () => {
    try {
      const elements = engine.listElements();
      if (elements.length === 0)
        return {
          content: [{ type: "text", text: "No elements in the diagram." }],
        };
      return {
        content: [{ type: "text", text: JSON.stringify(elements, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Connections ──────────────────────────────────────────────

server.tool(
  "add_arrow",
  `Connect two elements with an arrow.
Arrows automatically bind to the source and target elements.
By default draws a one-directional arrow (source → target).`,
  {
    from_id: z.string().describe("ID of the source element."),
    to_id: z.string().describe("ID of the target element."),
    label: z.string().optional().describe('Optional text label on the arrow (e.g. "HTTPS", "gRPC").'),
    stroke_color: z.string().optional().describe("Arrow color (hex or named Azure color)."),
    stroke_style: z
      .enum(["solid", "dashed", "dotted"])
      .optional()
      .describe("Line style. Use dashed for failover/replication paths."),
    bidirectional: z
      .boolean()
      .optional()
      .describe("If true, arrows on both ends (for replication links)."),
  },
  async ({ from_id, to_id, label, stroke_color, stroke_style, bidirectional }) => {
    try {
      const result = engine.addArrow(
        from_id,
        to_id,
        label,
        stroke_color,
        stroke_style,
        bidirectional ? "arrow" : undefined,
        "arrow",
      );
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "remove_arrow",
  "Remove an arrow by its ID. Also unbinds from connected elements.",
  {
    arrow_id: z.string().describe("The arrow's unique ID."),
  },
  async ({ arrow_id }) => {
    try {
      const msg = engine.removeArrow(arrow_id);
      return { content: [{ type: "text", text: msg }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Grouping & Layout ───────────────────────────────────────

server.tool(
  "add_frame",
  `Add an Excalidraw frame for visually grouping elements.
Frames provide a labeled boundary that can contain other elements.`,
  {
    x: z.number().describe("X position in pixels."),
    y: z.number().describe("Y position in pixels."),
    width: z.number().describe("Width in pixels."),
    height: z.number().describe("Height in pixels."),
    name: z.string().optional().describe("Frame label/title."),
  },
  async ({ x, y, width, height, name }) => {
    try {
      const result = engine.addFrame(x, y, width, height, name);
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

server.tool(
  "add_container",
  `Add a styled container rectangle for visually grouping shapes.
Styled per guide: dashed border, semi-transparent fill, with a label.
Use for grouping related services (e.g. "Web Tier", "Availability Zone 1").`,
  {
    x: z.number().describe("X position in pixels."),
    y: z.number().describe("Y position in pixels."),
    width: z.number().describe("Width in pixels."),
    height: z.number().describe("Height in pixels."),
    label: z.string().optional().describe("Title text displayed at the top of the container."),
    background_color: z
      .string()
      .optional()
      .describe('Fill color (hex or named Azure color). Defaults to light blue "#e6f3ff".'),
    opacity: z
      .number()
      .min(0)
      .max(100)
      .optional()
      .describe("Fill opacity from 0 (invisible) to 100 (fully opaque). Defaults to 40."),
  },
  async ({ x, y, width, height, label, background_color, opacity }) => {
    try {
      const result = engine.addContainer(x, y, width, height, label, background_color, opacity);
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Discovery ───────────────────────────────────────────────

server.tool(
  "list_azure_services",
  "List all available Azure service keys that can be used with add_azure_icon. Returns a sorted JSON array of service identifiers.",
  {},
  async () => {
    const services = engine.listAzureServices(AZURE_ICONS);
    return {
      content: [{ type: "text", text: JSON.stringify(services, null, 2) }],
    };
  },
);

// ── Start Server ────────────────────────────────────────────

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Excalidraw MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
