# Excalidraw Architecture Diagram — Style Guide

## Icons

- Always use **`add_azure_icon`** for Azure services — it embeds official SVG icons
- Use `list_azure_services` to discover available service keys
- Never use generic rectangles when an Azure icon is available

## Shapes

- **Clean lines**: `roughness = 0` (architect mode — no hand-drawn effect)
- **Solid fills**: `fillStyle = "solid"`
- **Stroke width**: `2px` default
- Use the **Azure brand color palette**:

| Color           | Hex         | Use For                          |
|-----------------|-------------|----------------------------------|
| Azure Blue      | `#0078D7`   | Load Balancers, SQL, general     |
| Dark Blue       | `#004E98`   | Internal LBs, Private Endpoints  |
| Teal            | `#00B294`   | Front Door, CDN                  |
| Orange          | `#FF8C00`   | Web tier, public-facing compute  |
| Purple          | `#8764B8`   | App tier, middleware             |
| Green           | `#7AB800`   | Storage, data lake               |
| Red             | `#E81123`   | Alerts, errors, critical paths   |

## Arrows

- **End arrowhead**: `"arrow"` (default)
- **Stroke width**: `2px`
- **Dashed lines** for failover, replication, or secondary paths: `strokeStyle = "dashed"`
- **Bidirectional arrows** for replication links (set both `startArrowhead` and `endArrowhead`)
- **Label**: 14px font, placed near midpoint

## Containers / Zones

- **Dashed border**: `strokeStyle = "dashed"`, `strokeWidth = 2`
- **Semi-transparent fill**: `opacity = 40` (out of 100)
- **Label**: 14px, positioned at top-left inside the container
- Color-match the border and fill to the zone's purpose (e.g., blue for network zones, green for data)

## Frames

- Use Excalidraw **frames** for logical grouping with a named boundary
- Frames have a thin gray dashed border
- Name frames descriptively (e.g., "Web Tier", "Data Layer")

## Page Layout

- **Coordinate system**: pixels from top-left origin (0,0)
- **Suggested canvas**: ~1056×816 pixels (landscape, equivalent to 11×8.5 in at 96dpi)
- **Content centered** with ~50px margins on all sides
- **Top-to-bottom flow**: users → frontend → web → app → data
- **Left/right symmetry** for availability zones or redundant paths
- **Icon spacing**: ~100px between Azure icons, ~200px between tiers
