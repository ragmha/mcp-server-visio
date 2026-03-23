#!/usr/bin/env node

/**
 * Visio MCP Server
 *
 * Exposes Microsoft Visio diagram operations as MCP tools via stdio transport.
 * Run with: npx mcp-server-visio
 *
 * Designed for GitHub Copilot CLI and VS Code Agent Mode.
 * All tools follow STYLE_GUIDE.md conventions automatically.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { VisioClient } from "./visio-client.js";

const server = new McpServer({
  name: "Visio Diagram Server",
  version: "1.0.0",
});

const visio = new VisioClient();

// ── Document Management ──────────────────────────────────────

server.tool(
  "create_diagram",
  `Create a new Visio diagram. The page is automatically set to landscape (11 × 8.5 in) per the style guide.`,
  {
    template: z
      .string()
      .optional()
      .describe(
        'Optional Visio template name or path (e.g. "Basic Diagram.vstx"). Leave empty for a blank drawing.',
      ),
  },
  async ({ template }) => {
    try {
      const name = visio.createDiagram(template);
      return { content: [{ type: "text", text: `Created diagram: ${name}` }] };
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
  "Save the active Visio diagram to a file.",
  {
    file_path: z
      .string()
      .describe(
        'Full path to save (e.g. "C:/Users/me/diagrams/arch.vsdx")',
      ),
  },
  async ({ file_path }) => {
    try {
      const path = visio.saveDiagram(file_path);
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
  "close_diagram",
  "Close the active Visio diagram without saving.",
  {},
  async () => {
    try {
      const msg = visio.closeDiagram();
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
  "list_open_diagrams",
  "List all open Visio documents with their page counts.",
  {},
  async () => {
    try {
      const docs = visio.listOpenDiagrams();
      if (docs.length === 0)
        return { content: [{ type: "text", text: "No open documents." }] };
      return {
        content: [{ type: "text", text: JSON.stringify(docs, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Shape Operations ─────────────────────────────────────────

server.tool(
  "add_shape",
  `Add a basic shape to the active Visio page.
Shapes are automatically styled: rounded corners (0.06 in), semi-transparent fills (15%).
For Azure service icons use add_azure_shape instead — it drops real stencil masters.`,
  {
    shape_type: z
      .string()
      .describe(
        "Type of shape: rectangle, square, ellipse, circle, diamond, triangle, rounded_rectangle, star. Or any master name from an open stencil.",
      ),
    x: z.number().describe("Horizontal position in inches from left edge."),
    y: z.number().describe("Vertical position in inches from bottom edge."),
    text: z.string().optional().describe("Label text to display inside the shape."),
    width: z.number().optional().describe("Width in inches (0 = default)."),
    height: z.number().optional().describe("Height in inches (0 = default)."),
    fill_color: z
      .string()
      .optional()
      .describe(
        'Fill color as hex RGB (e.g. "FF0000") or named: azure_blue, dark_blue, teal, orange, purple, green, red.',
      ),
  },
  async ({ shape_type, x, y, text, width, height, fill_color }) => {
    try {
      const result = visio.addShape(
        shape_type,
        x,
        y,
        text ?? "",
        width ?? 0,
        height ?? 0,
        fill_color,
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
  "add_azure_shape",
  `Add an Azure service icon from the official Azure Visio stencils.
ALWAYS prefer this over add_shape for Azure architecture diagrams.
The tool automatically opens the correct stencil file and drops the real Azure icon master.
Shapes get rounded corners and semi-transparent fills.`,
  {
    service: z
      .string()
      .describe(
        'Azure service name. Supports exact keys like "azure/front-door" or fuzzy names like "Front Door", "SQL Database", "VM Scale Sets". Use list_azure_services to see all 206 available services.',
      ),
    x: z.number().describe("Horizontal position in inches from left edge."),
    y: z.number().describe("Vertical position in inches from bottom edge."),
    text: z
      .string()
      .optional()
      .describe("Optional label (defaults to the master shape name)."),
    width: z
      .number()
      .optional()
      .describe("Width in inches (0 = default stencil size)."),
    height: z
      .number()
      .optional()
      .describe("Height in inches (0 = default stencil size)."),
    fill_color: z
      .string()
      .optional()
      .describe("Optional fill color override (hex RGB or named Azure color)."),
  },
  async ({ service, x, y, text, width, height, fill_color }) => {
    try {
      const result = visio.addAzureShape(
        service,
        x,
        y,
        text ?? "",
        width ?? 0,
        height ?? 0,
        fill_color,
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
  "remove_shape",
  "Remove a shape from the active page by its ID.",
  {
    shape_id: z
      .number()
      .describe("The numeric ID of the shape (from add_shape or list_shapes)."),
  },
  async ({ shape_id }) => {
    try {
      const msg = visio.removeShape(shape_id);
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
  "modify_shape",
  "Modify properties of an existing shape.",
  {
    shape_id: z.number().describe("The shape's numeric ID."),
    text: z
      .string()
      .optional()
      .describe("New text label (or omit to keep current)."),
    x: z
      .number()
      .optional()
      .describe("New X position in inches (or omit to keep)."),
    y: z
      .number()
      .optional()
      .describe("New Y position in inches (or omit to keep)."),
    width: z
      .number()
      .optional()
      .describe("New width in inches (or omit to keep)."),
    height: z
      .number()
      .optional()
      .describe("New height in inches (or omit to keep)."),
    fill_color: z
      .string()
      .optional()
      .describe(
        'Fill color as hex RGB (e.g. "FF0000") or named Azure color.',
      ),
  },
  async ({ shape_id, text, x, y, width, height, fill_color }) => {
    try {
      const result = visio.modifyShape(
        shape_id,
        text,
        x,
        y,
        width,
        height,
        fill_color,
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
  "list_shapes",
  "List all shapes on the active Visio page. Returns JSON array with each shape's id, name, text, position, and size.",
  {},
  async () => {
    try {
      const shapes = visio.listShapes();
      if (shapes.length === 0)
        return {
          content: [{ type: "text", text: "No shapes on the active page." }],
        };
      return {
        content: [{ type: "text", text: JSON.stringify(shapes, null, 2) }],
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
  "connect_shapes",
  `Connect two shapes with a styled connector line.
Automatically applies style-guide rules: filled triangle arrowheads, 1 pt line weight, 0.15 in rounding, 7 pt label font.`,
  {
    from_shape_id: z.number().describe("ID of the source shape."),
    to_shape_id: z.number().describe("ID of the target shape."),
    label: z
      .string()
      .optional()
      .describe('Optional text label on the connector (e.g. "HTTPS", "TDS").'),
    connector_style: z
      .enum(["straight", "curved", "right_angle"])
      .optional()
      .describe('One of "straight", "curved", or "right_angle".'),
    dashed: z
      .boolean()
      .optional()
      .describe(
        "If true, uses dashed line (for failover, replication, secondary paths).",
      ),
    bidirectional: z
      .boolean()
      .optional()
      .describe(
        "If true, arrows on both ends (for replication links).",
      ),
  },
  async ({
    from_shape_id,
    to_shape_id,
    label,
    connector_style,
    dashed,
    bidirectional,
  }) => {
    try {
      const result = visio.connectShapes(
        from_shape_id,
        to_shape_id,
        label ?? "",
        connector_style ?? "straight",
        dashed ?? false,
        bidirectional ?? false,
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
  "remove_connection",
  "Remove a connector by its ID.",
  {
    connector_id: z
      .number()
      .describe("The numeric ID of the connector shape."),
  },
  async ({ connector_id }) => {
    try {
      const msg = visio.removeConnection(connector_id);
      return { content: [{ type: "text", text: msg }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Architecture Helpers ─────────────────────────────────────

server.tool(
  "add_container",
  `Add a container/boundary rectangle for visually grouping shapes.
Styled per guide: dashed border, 60% transparent, 9 pt color-matched label at top.`,
  {
    x: z.number().describe("Center X position in inches."),
    y: z.number().describe("Center Y position in inches."),
    width: z.number().describe("Width in inches."),
    height: z.number().describe("Height in inches."),
    label: z
      .string()
      .optional()
      .describe("Title text displayed at the top of the container."),
    fill_color: z
      .string()
      .optional()
      .describe(
        'Fill color as hex RGB or named Azure color. Defaults to "E6F3FF" (light blue).',
      ),
    transparency: z
      .number()
      .optional()
      .describe(
        "Fill transparency from 0.0 (opaque) to 100 (invisible). Defaults to 60.",
      ),
    rounding: z
      .number()
      .optional()
      .describe("Corner rounding in inches. Defaults to 0 (sharp corners)."),
  },
  async ({ x, y, width, height, label, fill_color, transparency, rounding }) => {
    try {
      const result = visio.addContainer(
        x,
        y,
        width,
        height,
        label ?? "",
        fill_color ?? "E6F3FF",
        transparency,
        rounding,
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
  "add_tier_band",
  `Add a horizontal tier band spanning the full page width.
Styled per guide: 70% transparent, no border, bold 8 pt label on left margin.
Use for separating architecture tiers (e.g. "Web Tier", "App Tier", "Data Tier").`,
  {
    y: z.number().describe("Center Y position in inches."),
    height: z.number().describe("Band height in inches."),
    label: z
      .string()
      .optional()
      .describe('Tier label text (e.g. "Ingress", "Compute", "Data").'),
    fill_color: z
      .string()
      .optional()
      .describe("Fill color as hex RGB or named Azure color."),
    transparency: z
      .number()
      .optional()
      .describe("Fill transparency (0-100). Defaults to 70."),
    rounding: z
      .number()
      .optional()
      .describe("Corner rounding in inches. Defaults to 0."),
  },
  async ({ y, height, label, fill_color, transparency, rounding }) => {
    try {
      const result = visio.addTierBand(
        y,
        height,
        label ?? "",
        fill_color ?? "E6F3FF",
        transparency,
        rounding,
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
  "add_text_label",
  "Add a floating text label (no border, no fill) at the given position.",
  {
    x: z.number().describe("X position in inches."),
    y: z.number().describe("Y position in inches."),
    text: z.string().describe("The text to display."),
    font_size: z
      .number()
      .optional()
      .describe("Font size in points (default 10)."),
  },
  async ({ x, y, text, font_size }) => {
    try {
      const result = visio.addTextLabel(x, y, text, font_size ?? 10);
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

// ── Stencil Discovery ───────────────────────────────────────

server.tool(
  "list_azure_services",
  `List all 206 available Azure service keys that can be used with add_azure_shape.
Returns a sorted JSON array of service identifiers like "azure/front-door", "azure/sql-database", etc.`,
  {},
  async () => {
    const services = visio.listAzureServices();
    return {
      content: [{ type: "text", text: JSON.stringify(services, null, 2) }],
    };
  },
);

server.tool(
  "list_stencil_masters",
  "List all master shape names in a given Azure stencil. Use this to discover what icons are available in a specific stencil.",
  {
    stencil_name: z
      .string()
      .describe(
        'Stencil logical name (e.g. "Azure-Databases", "Azure-Compute", "Azure-Networking", "Azure-Web").',
      ),
  },
  async ({ stencil_name }) => {
    try {
      const masters = visio.listStencilMasters(stencil_name);
      return {
        content: [{ type: "text", text: JSON.stringify(masters, null, 2) }],
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
  "open_stencil",
  "Open an Azure stencil by name so its masters become available.",
  {
    stencil_name: z
      .string()
      .describe('Stencil logical name (e.g. "Azure-Databases").'),
  },
  async ({ stencil_name }) => {
    try {
      const name = visio.openStencil(stencil_name);
      return {
        content: [{ type: "text", text: `Opened stencil: ${name}` }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Page Operations ──────────────────────────────────────────

server.tool(
  "add_page",
  "Add a new page to the active Visio document.",
  {
    name: z
      .string()
      .optional()
      .describe('Optional name for the new page (e.g. "Network Layer").'),
  },
  async ({ name }) => {
    try {
      const result = visio.addPage(name ?? "");
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
  "set_active_page",
  "Switch to a specific page by its index (1-based).",
  {
    page_index: z
      .number()
      .describe("The page number to activate (1 = first page)."),
  },
  async ({ page_index }) => {
    try {
      const msg = visio.setActivePage(page_index);
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
  "list_pages",
  "List all pages in the active Visio document.",
  {},
  async () => {
    try {
      const pages = visio.listPages();
      return {
        content: [{ type: "text", text: JSON.stringify(pages, null, 2) }],
      };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Export ───────────────────────────────────────────────────

server.tool(
  "export_page",
  'Export the active page as an image (PNG, SVG, JPG, etc.). The format is determined by the file extension.',
  {
    file_path: z
      .string()
      .describe(
        'Output path, e.g. "C:/Users/me/diagram.png" or "output.svg".',
      ),
  },
  async ({ file_path }) => {
    try {
      const path = visio.exportPage(file_path);
      return { content: [{ type: "text", text: `Exported to: ${path}` }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: `Error: ${(e as Error).message}` }],
        isError: true,
      };
    }
  },
);

// ── Start Server ────────────────────────────────────────────

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Visio MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
