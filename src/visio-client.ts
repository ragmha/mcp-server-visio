/**
 * Visio COM Automation Client
 *
 * Controls Microsoft Visio via PowerShell COM interop.
 * This approach requires zero native Node.js addons, making npx work seamlessly.
 * All styling follows STYLE_GUIDE.md conventions automatically.
 */

import { execSync, type ExecSyncOptionsWithStringEncoding } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { STENCIL_MAP } from "./stencil-map.js";
import {
  AZURE_COLORS,
  type AzureColor,
  type ShapeResult,
  type ConnectorResult,
  type PageResult,
  type DocumentResult,
} from "./types.js";

// Style constants from STYLE_GUIDE.md
const SHAPE_ROUNDING = 0.06;
const SHAPE_FILL_TRANSPARENCY = 15;
const CONNECTOR_END_ARROW = 4;
const CONNECTOR_ARROW_SIZE = 2;
const CONNECTOR_ROUNDING = 0.15;
const CONNECTOR_LINE_WEIGHT = "1 pt";
const CONNECTOR_LABEL_FONT = 7;
const CONNECTOR_DASHED_PATTERN = 2;
const CONTAINER_LINE_PATTERN = 2;
const CONTAINER_FILL_TRANS = 60;
const CONTAINER_LABEL_FONT = 9;
const TIER_BAND_FILL_TRANS = 70;
const TIER_BAND_LABEL_FONT = 8;

// Visio shape type → DrawXxx method or stencil master
const DRAW_METHODS: Record<string, string> = {
  rectangle: "DrawRectangle",
  square: "DrawRectangle",
  rounded_rectangle: "DrawRectangle",
  ellipse: "DrawOval",
  circle: "DrawOval",
};

// Connector routing style constants (Visio RouteStyle cell values)
const CONNECTOR_ROUTE_STYLES: Record<string, number> = {
  straight: 16,     // visLORouteStraight
  right_angle: 1,   // visLORouteRightAngle
  curved: 2,        // visLORouteCurve
};

export class VisioClient {
  private _tmpDir: string | null = null;

  /** Get or create a private temp directory for script files. */
  private get tmpDir(): string {
    if (!this._tmpDir || !fs.existsSync(this._tmpDir)) {
      this._tmpDir = fs.mkdtempSync(
        path.join(fs.realpathSync(process.env.TEMP || process.env.TMP || "/tmp"), "visio-mcp-"),
      );
    }
    return this._tmpDir;
  }

  /**
   * Execute a PowerShell script via a temp file in a private directory.
   * Catches execSync errors and extracts JSON error payloads from stdout.
   */
  private run(script: string): unknown {
    const nonce = crypto.randomBytes(8).toString("hex");
    const tmpFile = path.join(this.tmpDir, `${nonce}.ps1`);

    const fullScript = `
$ErrorActionPreference = 'Stop'
try {
${script}
} catch {
  @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
}
`;

    fs.writeFileSync(tmpFile, fullScript, { encoding: "utf-8", mode: 0o600 });
    try {
      const execOpts: ExecSyncOptionsWithStringEncoding = {
        encoding: "utf-8",
        timeout: 30000,
        maxBuffer: 10 * 1024 * 1024,
      };

      let output: string;
      try {
        output = execSync(
          `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "${tmpFile}"`,
          execOpts,
        );
      } catch (execErr: unknown) {
        // execSync throws on non-zero exit; extract stdout for our JSON error payload
        const err = execErr as { stdout?: string; stderr?: string; message?: string };
        const stdout = (err.stdout ?? "").trim();
        if (stdout) {
          try {
            const parsed = JSON.parse(stdout);
            if (parsed?.error) throw new Error(parsed.error);
          } catch (parseErr) {
            if (parseErr instanceof Error && parseErr.message !== stdout) throw parseErr;
          }
        }
        throw new Error(`PowerShell execution failed: ${err.stderr || err.message || "unknown error"}`);
      }

      const trimmed = output.trim();
      if (!trimmed) return null;

      try {
        const parsed = JSON.parse(trimmed);
        if (parsed?.error) throw new Error(parsed.error);
        return parsed;
      } catch (parseErr) {
        if (parseErr instanceof Error && parseErr.message !== trimmed) throw parseErr;
        return trimmed;
      }
    } finally {
      try {
        fs.unlinkSync(tmpFile);
      } catch {
        // ignore cleanup errors
      }
    }
  }

  // ── Document Management ──────────────────────────────────

  createDiagram(template: string = ""): string {
    const result = this.run(`
try { $visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application') }
catch { $visio = New-Object -ComObject Visio.Application; $visio.Visible = $true }
${template ? `$doc = $visio.Documents.Add("${this.escapePS(template)}")` : '$doc = $visio.Documents.Add("")'}
$page = $doc.Pages.Item(1)
$page.PageSheet.Cells("PageWidth").FormulaU = "11 in"
$page.PageSheet.Cells("PageHeight").FormulaU = "8.5 in"
$page.PageSheet.Cells("PrintPageOrientation").FormulaU = "2"
@{ name = $doc.Name } | ConvertTo-Json -Compress
`);
    const data = result as { name: string; error?: string };
    if (data?.error) throw new Error(data.error);
    return data.name;
  }

  saveDiagram(filePath: string): string {
    const absPath = filePath.replace(/\//g, "\\");
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$doc = $visio.ActiveDocument
if (-not $doc) { throw "No active document to save." }
$doc.SaveAs("${this.escapePS(absPath)}")
@{ path = $doc.FullName } | ConvertTo-Json -Compress
`);
    const data = result as { path: string; error?: string };
    if (data?.error) throw new Error(data.error);
    return data.path;
  }

  closeDiagram(): string {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$doc = $visio.ActiveDocument
if (-not $doc) { @{ message = "No active document." } | ConvertTo-Json -Compress; return }
$name = $doc.Name
$doc.Close()
@{ message = "Closed '$name'." } | ConvertTo-Json -Compress
`);
    const data = result as { message: string; error?: string };
    if (data?.error) throw new Error(data.error);
    return data.message;
  }

  listOpenDiagrams(): DocumentResult[] {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$docs = @()
for ($i = 1; $i -le $visio.Documents.Count; $i++) {
  $doc = $visio.Documents.Item($i)
  $docs += @{
    index = $i
    name = $doc.Name
    path = $doc.FullName
    pages = $doc.Pages.Count
  }
}
$docs | ConvertTo-Json -Compress
`);
    if (!result) return [];
    return (Array.isArray(result) ? result : [result]) as DocumentResult[];
  }

  // ── Shape Operations ─────────────────────────────────────

  addShape(
    shapeType: string,
    x: number,
    y: number,
    text: string = "",
    width: number = 0,
    height: number = 0,
    fillColor?: string,
  ): ShapeResult {
    const colorCmd = this.buildColorCommand(fillColor);
    const drawMethod = DRAW_METHODS[shapeType.toLowerCase()] ?? "DrawRectangle";
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page. Create a diagram first." }

# Drop shape using appropriate draw method
$shape = $page.${drawMethod}(${x - 0.75}, ${y - 0.375}, ${x + 0.75}, ${y + 0.375})

${text ? `$shape.Text = "${this.escapePS(text)}"` : ""}
${width > 0 ? `$shape.Cells("Width").ResultIU = ${width}` : ""}
${height > 0 ? `$shape.Cells("Height").ResultIU = ${height}` : ""}

# Style guide: rounded corners, semi-transparent fill
try { $shape.Cells("Rounding").FormulaU = "${SHAPE_ROUNDING} in" } catch {}
try { $shape.Cells("FillForegndTrans").FormulaU = "${SHAPE_FILL_TRANSPARENCY}%" } catch {}
${colorCmd}

@{
  id = $shape.ID
  name = $shape.Name
  text = "${this.escapePS(text)}"
  x = ${x}
  y = ${y}
} | ConvertTo-Json -Compress
`);
    const data = result as ShapeResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  addAzureShape(
    service: string,
    x: number,
    y: number,
    text: string = "",
    width: number = 0,
    height: number = 0,
    fillColor?: string,
  ): ShapeResult {
    const resolved = this.resolveAzureService(service);
    if (!resolved) {
      throw new Error(
        `Unknown Azure service '${service}'. Use list_azure_services to see available services.`,
      );
    }

    const { stencil: stencilName, master: masterName } = resolved;
    const stencilFile = `${stencilName}.vssx`;
    const colorCmd = this.buildColorCommand(fillColor);

    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page. Create a diagram first." }

# Open stencil if not already open
$stencil = $null
for ($i = 1; $i -le $visio.Documents.Count; $i++) {
  $doc = $visio.Documents.Item($i)
  if ($doc.Name -eq "${stencilFile}") { $stencil = $doc; break }
}
if (-not $stencil) {
  $stencil = $visio.Documents.OpenEx("${stencilFile}", 4)  # visOpenDocked
}

$master = $stencil.Masters.Item("${this.escapePS(masterName)}")
$shape = $page.Drop($master, ${x}, ${y})

${text ? `$shape.Text = "${this.escapePS(text)}"` : ""}
${width > 0 ? `$shape.Cells("Width").ResultIU = ${width}` : ""}
${height > 0 ? `$shape.Cells("Height").ResultIU = ${height}` : ""}

try { $shape.Cells("Rounding").FormulaU = "${SHAPE_ROUNDING} in" } catch {}
try { $shape.Cells("FillForegndTrans").FormulaU = "${SHAPE_FILL_TRANSPARENCY}%" } catch {}
${colorCmd}

@{
  id = $shape.ID
  name = $shape.Name
  text = if ("${this.escapePS(text)}") { "${this.escapePS(text)}" } else { "${this.escapePS(masterName)}" }
  x = ${x}
  y = ${y}
  azure_service = "azure/${this.fuzzyKey(service).split("/").pop()}"
  stencil = "${stencilName}"
  master = "${this.escapePS(masterName)}"
} | ConvertTo-Json -Compress
`);
    const data = result as ShapeResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  removeShape(shapeId: number): string {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }
$found = $false
for ($i = 1; $i -le $page.Shapes.Count; $i++) {
  $shape = $page.Shapes.Item($i)
  if ($shape.ID -eq ${shapeId}) {
    $name = $shape.Name
    $shape.Delete()
    @{ message = "Deleted shape '$name' (ID=${shapeId})." } | ConvertTo-Json -Compress
    $found = $true
    break
  }
}
if (-not $found) { throw "Shape with ID ${shapeId} not found on active page." }
`);
    const data = result as { message: string; error?: string };
    if (data?.error) throw new Error(data.error);
    return data.message;
  }

  modifyShape(
    shapeId: number,
    text?: string,
    x?: number,
    y?: number,
    width?: number,
    height?: number,
    fillColor?: string,
  ): ShapeResult {
    const colorCmd = fillColor ? this.buildColorCommand(fillColor) : "";
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }
$shape = $null
for ($i = 1; $i -le $page.Shapes.Count; $i++) {
  $s = $page.Shapes.Item($i)
  if ($s.ID -eq ${shapeId}) { $shape = $s; break }
}
if (-not $shape) { throw "Shape with ID ${shapeId} not found." }

${text !== undefined ? `$shape.Text = "${this.escapePS(text ?? "")}"` : ""}
${x !== undefined ? `$shape.Cells("PinX").ResultIU = ${x}` : ""}
${y !== undefined ? `$shape.Cells("PinY").ResultIU = ${y}` : ""}
${width !== undefined ? `$shape.Cells("Width").ResultIU = ${width}` : ""}
${height !== undefined ? `$shape.Cells("Height").ResultIU = ${height}` : ""}
${colorCmd}

@{
  id = $shape.ID
  name = $shape.Name
  text = $shape.Text
} | ConvertTo-Json -Compress
`);
    const data = result as ShapeResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  listShapes(): ShapeResult[] {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }
$shapes = @()
for ($i = 1; $i -le $page.Shapes.Count; $i++) {
  $shape = $page.Shapes.Item($i)
  try {
    $shapes += @{
      id = $shape.ID
      name = $shape.Name
      text = $shape.Text
      x = [math]::Round($shape.Cells("PinX").ResultIU, 2)
      y = [math]::Round($shape.Cells("PinY").ResultIU, 2)
      width = [math]::Round($shape.Cells("Width").ResultIU, 2)
      height = [math]::Round($shape.Cells("Height").ResultIU, 2)
    }
  } catch {}
}
$shapes | ConvertTo-Json -Compress
`);
    if (!result) return [];
    return (Array.isArray(result) ? result : [result]) as ShapeResult[];
  }

  // ── Connections ──────────────────────────────────────────

  connectShapes(
    fromShapeId: number,
    toShapeId: number,
    label: string = "",
    connectorStyle: string = "straight",
    dashed: boolean = false,
    bidirectional: boolean = false,
  ): ConnectorResult {
    const linePattern = dashed ? CONNECTOR_DASHED_PATTERN : 1;
    const routeStyle = CONNECTOR_ROUTE_STYLES[connectorStyle] ?? CONNECTOR_ROUTE_STYLES.straight;
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }

$fromShape = $null; $toShape = $null
for ($i = 1; $i -le $page.Shapes.Count; $i++) {
  $s = $page.Shapes.Item($i)
  if ($s.ID -eq ${fromShapeId}) { $fromShape = $s }
  if ($s.ID -eq ${toShapeId}) { $toShape = $s }
}
if (-not $fromShape) { throw "Source shape ID ${fromShapeId} not found." }
if (-not $toShape) { throw "Target shape ID ${toShapeId} not found." }

$connector = $page.Drop($visio.ConnectorToolDataObject, 0, 0)
$connector.Cells("BeginX").GlueTo($fromShape.Cells("PinX"))
$connector.Cells("EndX").GlueTo($toShape.Cells("PinX"))

# Style guide: filled triangle arrowheads, 1pt weight, rounded bends
$connector.Cells("EndArrow").FormulaU = "${CONNECTOR_END_ARROW}"
$connector.Cells("EndArrowSize").FormulaU = "${CONNECTOR_ARROW_SIZE}"
$connector.Cells("Rounding").FormulaU = "${CONNECTOR_ROUNDING} in"
$connector.Cells("LineWeight").FormulaU = "${CONNECTOR_LINE_WEIGHT}"
$connector.Cells("LinePattern").FormulaU = "${linePattern}"
$connector.Cells("ShapeRouteStyle").FormulaU = "${routeStyle}"

${bidirectional ? `$connector.Cells("BeginArrow").FormulaU = "${CONNECTOR_END_ARROW}"` : ""}
${label ? `$connector.Text = "${this.escapePS(label)}"; try { $connector.Cells("Char.Size").FormulaU = "${CONNECTOR_LABEL_FONT} pt" } catch {}` : ""}

@{
  id = $connector.ID
  name = $connector.Name
  from_id = ${fromShapeId}
  to_id = ${toShapeId}
  label = "${this.escapePS(label)}"
} | ConvertTo-Json -Compress
`);
    const data = result as ConnectorResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  removeConnection(connectorId: number): string {
    return this.removeShape(connectorId);
  }

  // ── Architecture Helpers ─────────────────────────────────

  addContainer(
    x: number,
    y: number,
    width: number,
    height: number,
    label: string = "",
    fillColor: string = "E6F3FF",
    transparency?: number,
    rounding?: number,
  ): ShapeResult {
    const trans = transparency ?? CONTAINER_FILL_TRANS;
    const round = rounding ?? 0;
    const { r, g, b } = this.parseColor(fillColor);

    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }

$shape = $page.DrawRectangle(${x - width / 2}, ${y - height / 2}, ${x + width / 2}, ${y + height / 2})
$shape.Cells("FillForegnd").FormulaU = "RGB(${r},${g},${b})"
$shape.Cells("FillForegndTrans").FormulaU = "${trans}%"
$shape.Cells("LinePattern").FormulaU = "${CONTAINER_LINE_PATTERN}"
$shape.Cells("LineWeight").FormulaU = "1 pt"
${round > 0 ? `$shape.Cells("Rounding").FormulaU = "${round} in"` : ""}

${label ? `
$shape.Text = "${this.escapePS(label)}"
try {
  $shape.Cells("Char.Size").FormulaU = "${CONTAINER_LABEL_FONT} pt"
  $shape.Cells("TxtPinY").FormulaU = "Height*0.96"
  $shape.Cells("Char.Color").FormulaU = "RGB(${r},${g},${b})"
} catch {}
` : ""}

# Send to back so shapes appear on top
try { $shape.SendToBack() } catch {}

@{
  id = $shape.ID
  name = $shape.Name
  text = "${this.escapePS(label)}"
  x = ${x}
  y = ${y}
  width = ${width}
  height = ${height}
} | ConvertTo-Json -Compress
`);
    const data = result as ShapeResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  addTierBand(
    y: number,
    height: number,
    label: string = "",
    fillColor: string = "E6F3FF",
    transparency?: number,
    rounding?: number,
  ): ShapeResult {
    const trans = transparency ?? TIER_BAND_FILL_TRANS;
    const round = rounding ?? 0;
    const { r, g, b } = this.parseColor(fillColor);

    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }

$pageWidth = $page.PageSheet.Cells("PageWidth").ResultIU
$shape = $page.DrawRectangle(0, ${y - height / 2}, $pageWidth, ${y + height / 2})
$shape.Cells("FillForegnd").FormulaU = "RGB(${r},${g},${b})"
$shape.Cells("FillForegndTrans").FormulaU = "${trans}%"
$shape.Cells("LinePattern").FormulaU = "0"
${round > 0 ? `$shape.Cells("Rounding").FormulaU = "${round} in"` : ""}

${label ? `
$shape.Text = "${this.escapePS(label)}"
try {
  $shape.Cells("Char.Size").FormulaU = "${TIER_BAND_LABEL_FONT} pt"
  $shape.Cells("Char.Style").FormulaU = "1"
  $shape.Cells("TxtPinX").FormulaU = "0.5 in"
  $shape.Cells("TxtWidth").FormulaU = "2 in"
  $shape.Cells("Para.HorzAlign").FormulaU = "0"
} catch {}
` : ""}

try { $shape.SendToBack() } catch {}

@{
  id = $shape.ID
  name = $shape.Name
  text = "${this.escapePS(label)}"
  x = [math]::Round($pageWidth / 2, 2)
  y = ${y}
  width = [math]::Round($pageWidth, 2)
  height = ${height}
} | ConvertTo-Json -Compress
`);
    const data = result as ShapeResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  addTextLabel(
    x: number,
    y: number,
    text: string,
    fontSize: number = 10,
  ): ShapeResult {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }

$shape = $page.DrawRectangle(${x - 0.5}, ${y - 0.15}, ${x + 0.5}, ${y + 0.15})
$shape.Text = "${this.escapePS(text)}"
$shape.Cells("LinePattern").FormulaU = "0"
$shape.Cells("FillPattern").FormulaU = "0"
try { $shape.Cells("Char.Size").FormulaU = "${fontSize} pt" } catch {}

@{
  id = $shape.ID
  name = $shape.Name
  text = "${this.escapePS(text)}"
  x = ${x}
  y = ${y}
} | ConvertTo-Json -Compress
`);
    const data = result as ShapeResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  // ── Stencil Discovery ────────────────────────────────────

  listAzureServices(): string[] {
    return Object.keys(STENCIL_MAP).sort();
  }

  listStencilMasters(stencilName: string): string[] {
    const escapedName = this.escapePS(stencilName);
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')

$stencilFile = "${escapedName}.vssx"
$stencil = $null
for ($i = 1; $i -le $visio.Documents.Count; $i++) {
  $doc = $visio.Documents.Item($i)
  if ($doc.Name -eq $stencilFile) { $stencil = $doc; break }
}
if (-not $stencil) {
  $stencil = $visio.Documents.OpenEx($stencilFile, 4)
}

$masters = @()
for ($i = 1; $i -le $stencil.Masters.Count; $i++) {
  $masters += $stencil.Masters.Item($i).Name
}
$masters | ConvertTo-Json -Compress
`);
    if (!result) return [];
    return (Array.isArray(result) ? result : [result]) as string[];
  }

  openStencil(stencilName: string): string {
    const escapedName = this.escapePS(stencilName);
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$stencilFile = "${escapedName}.vssx"
$stencil = $null
for ($i = 1; $i -le $visio.Documents.Count; $i++) {
  $doc = $visio.Documents.Item($i)
  if ($doc.Name -eq $stencilFile) { $stencil = $doc; break }
}
if (-not $stencil) {
  $stencil = $visio.Documents.OpenEx($stencilFile, 4)
}
@{ name = $stencil.Name } | ConvertTo-Json -Compress
`);
    const data = result as { name: string; error?: string };
    if (data?.error) throw new Error(data.error);
    return data.name;
  }

  // ── Page Operations ──────────────────────────────────────

  addPage(name: string = ""): PageResult {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$doc = $visio.ActiveDocument
if (-not $doc) { throw "No active document." }
$page = $doc.Pages.Add()
${name ? `$page.Name = "${this.escapePS(name)}"` : ""}

# Set landscape per style guide
$page.PageSheet.Cells("PageWidth").FormulaU = "11 in"
$page.PageSheet.Cells("PageHeight").FormulaU = "8.5 in"
$page.PageSheet.Cells("PrintPageOrientation").FormulaU = "2"

@{
  index = $page.Index
  name = $page.Name
} | ConvertTo-Json -Compress
`);
    const data = result as PageResult & { error?: string };
    if (data?.error) throw new Error(data.error as string);
    return data;
  }

  setActivePage(pageIndex: number): string {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$doc = $visio.ActiveDocument
if (-not $doc) { throw "No active document." }
$page = $doc.Pages.Item(${pageIndex})
$visio.ActiveWindow.Page = $page
@{ message = "Switched to page $($page.Index): $($page.Name)" } | ConvertTo-Json -Compress
`);
    const data = result as { message: string; error?: string };
    if (data?.error) throw new Error(data.error);
    return data.message;
  }

  listPages(): PageResult[] {
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$doc = $visio.ActiveDocument
if (-not $doc) { throw "No active document." }
$pages = @()
for ($i = 1; $i -le $doc.Pages.Count; $i++) {
  $page = $doc.Pages.Item($i)
  $pages += @{
    index = $page.Index
    name = $page.Name
  }
}
$pages | ConvertTo-Json -Compress
`);
    if (!result) return [];
    return (Array.isArray(result) ? result : [result]) as PageResult[];
  }

  // ── Export ───────────────────────────────────────────────

  exportPage(filePath: string): string {
    const absPath = filePath.replace(/\//g, "\\");
    const escapedPath = this.escapePS(absPath);
    const result = this.run(`
$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$page = $visio.ActivePage
if (-not $page) { throw "No active page." }
$page.Export("${escapedPath}")
@{ path = "${escapedPath}" } | ConvertTo-Json -Compress
`);
    const data = result as { path: string; error?: string };
    if (data?.error) throw new Error(data.error);
    return data.path;
  }

  // ── Helpers ──────────────────────────────────────────────

  private resolveAzureService(
    service: string,
  ): { stencil: string; master: string } | null {
    // Try exact key
    if (STENCIL_MAP[service]) return STENCIL_MAP[service];

    // Try fuzzy key
    const key = this.fuzzyKey(service);
    if (STENCIL_MAP[key]) return STENCIL_MAP[key];

    return null;
  }

  private fuzzyKey(name: string): string {
    let normalized = name
      .toLowerCase()
      .replace(/ /g, "-")
      .replace(/_/g, "-")
      .replace(/^azure-/, "");
    if (!normalized.startsWith("azure/")) {
      normalized = `azure/${normalized}`;
    }
    return normalized;
  }

  private parseColor(color: string): AzureColor {
    const lower = color.toLowerCase();
    if (AZURE_COLORS[lower]) return AZURE_COLORS[lower];
    const hex = color.replace(/^#/, "");
    if (hex.length === 6) {
      return {
        r: parseInt(hex.substring(0, 2), 16),
        g: parseInt(hex.substring(2, 4), 16),
        b: parseInt(hex.substring(4, 6), 16),
      };
    }
    throw new Error(
      `Invalid color '${color}'. Use hex RGB or: ${Object.keys(AZURE_COLORS).join(", ")}`,
    );
  }

  private buildColorCommand(fillColor?: string): string {
    if (!fillColor) return "";
    const { r, g, b } = this.parseColor(fillColor);
    return `try { $shape.Cells("FillForegnd").FormulaU = "RGB(${r},${g},${b})" } catch {}`;
  }

  /**
   * Escape a string for safe embedding inside a PowerShell double-quoted string.
   * Backtick is PowerShell's escape char and must be escaped first.
   */
  private escapePS(s: string): string {
    return s
      .replace(/`/g, "``")    // backtick must be first (PS escape char)
      .replace(/"/g, '`"')    // double quote
      .replace(/\$/g, "`$")   // dollar sign (prevents variable expansion)
      .replace(/\0/g, "`0");  // null byte
  }
}
