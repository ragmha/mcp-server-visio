# Visio MCP Server (npx)

A Node.js MCP server that exposes Microsoft Visio diagram operations as tools — generate production-grade Azure architecture diagrams from text descriptions.

Run with a single command:

```bash
npx mcp-server-visio
```

Built for **GitHub Copilot CLI** and **VS Code Agent Mode**, but works with any MCP client.

## Features

- **Azure service icons** — 206 Azure services from official Visio stencils (auto-discovered)
- **Architecture helpers** — tier bands, containers, connectors with style-guide compliance
- **Shape operations** — add, modify, remove, connect, and list shapes
- **Multi-page support** — add pages, switch between them
- **Export** — PNG, SVG, JPG output
- **Zero native deps** — uses PowerShell COM interop, no compilation required
- **npx-ready** — just `npx mcp-server-visio` and go

## Prerequisites

- **Windows** with **Microsoft Visio Professional** (installed and licensed)
- **Node.js 18+**
- **PowerShell** (comes with Windows)
- **Azure Visio stencils** — download from [Microsoft Azure Architecture Icons](https://learn.microsoft.com/en-us/azure/architecture/icons/) and extract to your `My Shapes` folder

## Installation

### Option 1: npx (recommended)

No installation needed — just configure and run:

```bash
npx mcp-server-visio
```

### Option 2: Global install

```bash
npm install -g mcp-server-visio
mcp-server-visio
```

### Option 3: Local development

```bash
git clone <repo-url>
cd mcp-server-visio
npm install
npm run build
npm start
```

## Configuration

### Copilot CLI

Add to `~/.copilot/mcp-config.json`:

```json
{
  "mcpServers": {
    "visio": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "mcp-server-visio"]
    }
  }
}
```

### VS Code

Add to `.vscode/mcp.json` or user settings:

```json
{
  "mcpServers": {
    "visio": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "mcp-server-visio"]
    }
  }
}
```

### Whitelisting Tools

Auto-approve all Visio tools:

```bash
copilot --allow-tool "visio"
```

Or persist in `~/.copilot/config.json`:

```json
{
  "allowedTools": ["visio"]
}
```

## Available Tools

| Tool | Description |
|---|---|
| `create_diagram` | Create a new Visio diagram (landscape, 11×8.5 in) |
| `save_diagram` | Save to `.vsdx` file |
| `close_diagram` | Close without saving |
| `list_open_diagrams` | List all open documents |
| `add_shape` | Add basic shapes (rectangle, ellipse, diamond, etc.) |
| `add_azure_shape` | Add Azure service icons from official stencils |
| `remove_shape` | Remove a shape by ID |
| `modify_shape` | Change text, position, size, or color |
| `list_shapes` | List all shapes on the active page |
| `connect_shapes` | Connect two shapes with styled connectors |
| `remove_connection` | Remove a connector |
| `add_container` | Add a grouping boundary rectangle |
| `add_tier_band` | Add a full-width horizontal tier band |
| `add_text_label` | Add a floating text label |
| `list_azure_services` | List all 206 available Azure service keys |
| `list_stencil_masters` | List masters in a specific stencil |
| `open_stencil` | Open an Azure stencil by name |
| `add_page` | Add a new page |
| `set_active_page` | Switch to a page by index |
| `list_pages` | List all pages |
| `export_page` | Export page as image (PNG, SVG, JPG) |

## Style Guide

All shapes and connectors are automatically styled per `STYLE_GUIDE.md`:

- **Shapes**: rounded corners (0.06 in), 15% transparent fills
- **Connectors**: filled triangle arrowheads, 1 pt weight, 7 pt label font
- **Containers**: dashed border, 60% transparent, 9 pt label
- **Tier bands**: 70% transparent, bold 8 pt label

## Example

```
Create a 3-tier Azure architecture with Front Door, VM Scale Sets in 2 availability zones, and Azure SQL with replication
```

The server will create a professional Visio diagram with proper Azure icons, tier bands, containers, and styled connectors.

## Architecture

Unlike the Python version which uses `win32com` directly, this Node.js version uses **PowerShell COM interop** — each Visio operation translates to a PowerShell script executed via `child_process`. This gives us:

- **Zero native Node.js addons** — no `node-gyp`, no compilation
- **Works with npx** — download and run instantly
- **Same COM control** — full access to Visio's COM object model
- **Robust error handling** — structured JSON responses from PowerShell

## License

MIT
