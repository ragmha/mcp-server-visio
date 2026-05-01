#!/usr/bin/env tsx
/**
 * Build Azure Icons
 *
 * Downloads the official Microsoft Azure Architecture Icons pack,
 * extracts SVGs, base64-encodes them, and generates src/azure-icons.ts.
 *
 * Source: https://arch-center.azureedge.net/icons/Azure_Public_Service_Icons_V22.zip
 * Same source used by github.com/dwarfered/azure-architecture-icons-for-drawio
 *
 * Usage: npx tsx scripts/build-azure-icons.ts
 */

import { execSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const ICON_ZIP_URL = "https://arch-center.azureedge.net/icons/Azure_Public_Service_Icons_V22.zip";
const TMP_DIR = path.join(process.cwd(), "tmp-azure-icons");
const OUTPUT_FILE = path.join(process.cwd(), "src", "azure-icons.ts");

// Map of Azure service keys to expected SVG filename patterns.
// Each entry maps a service key to a search term used to find the matching SVG.
const SERVICE_TO_SVG: Record<string, string> = {
  "azure/front-door": "Front Door",
  "azure/cdn": "CDN",
  "azure/app-services": "App Services",
  "azure/static-apps": "Static Apps",
  "azure/api-management": "API Management",
  "azure/cosmos-db": "Cosmos DB",
  "azure/sql-database": "SQL Database",
  "azure/sql-server": "SQL Server",
  "azure/sql-managed-instance": "SQL Managed Instance",
  "azure/redis-cache": "Cache Redis",
  "azure/storage-accounts": "Storage Accounts",
  "azure/virtual-networks": "Virtual Networks",
  "azure/load-balancers": "Load Balancers",
  "azure/application-gateway": "Application Gateways",
  "azure/firewall": "Firewalls",
  "azure/dns-zones": "DNS Zones",
  "azure/virtual-machines": "Virtual Machine",
  "azure/vm-scale-sets": "VM Scale Sets",
  "azure/function-apps": "Function Apps",
  "azure/container-instances": "Container Instances",
  "azure/container-registries": "Container Registries",
  "azure/kubernetes-services": "Kubernetes Services",
  "azure/key-vault": "Key Vaults",
  "azure/security-center": "Defender For Cloud",
  "azure/sentinel": "Sentinel",
  "azure/logic-apps": "Logic Apps",
  "azure/service-bus": "Service Bus",
  "azure/event-hubs": "Event Hubs",
  "azure/event-grid": "Event Grid",
  "azure/monitor": "Monitor",
  "azure/application-insights": "Application Insights",
  "azure/databricks": "Databricks",
  "azure/synapse-analytics": "Synapse Analytics",
  "azure/data-factory": "Data Factory",
  "azure/machine-learning": "Machine Learning",
  "azure/openai": "OpenAI",
  "azure/iot-hub": "IOT Hub",
  "azure/vpn-gateway": "Virtual Network Gateways",
  "azure/expressroute": "ExpressRoute",
  "azure/bastion": "Bastions",
  "azure/private-endpoint": "Private Endpoints",
  "azure/nat-gateway": "NAT",
  "azure/container-apps": "Container App",
  "azure/batch": "Batch Accounts",
  "azure/service-fabric": "Service Fabric",
  "azure/notification-hubs": "Notification Hubs",
  "azure/signalr": "SignalR",
  "azure/cognitive-services": "Cognitive Services",
  "azure/bot-services": "Bot Services",
  "azure/devops": "Azure DevOps",
  "azure/policy": "Policy",
  "azure/advisor": "Advisor",
  "azure/cost-management": "Cost Management",
  "azure/azure-arc": "Azure Arc",
  "azure/managed-identities": "Managed Identities",
  "azure/entra-connect": "Entra Connect",
};

async function main() {
  console.log("🔽 Downloading Azure Architecture Icons...");

  // Clean up
  if (fs.existsSync(TMP_DIR)) {
    fs.rmSync(TMP_DIR, { recursive: true });
  }
  fs.mkdirSync(TMP_DIR, { recursive: true });

  const zipPath = path.join(TMP_DIR, "icons.zip");

  // Download
  execSync(`curl -sL "${ICON_ZIP_URL}" -o "${zipPath}"`, { stdio: "inherit" });

  // Extract
  console.log("📦 Extracting...");
  execSync(`unzip -q "${zipPath}" -d "${TMP_DIR}"`, { stdio: "inherit" });

  // Find SVG files
  console.log("🔍 Finding SVGs...");
  const svgFiles = findSvgFiles(TMP_DIR);
  console.log(`Found ${svgFiles.length} SVG files`);

  // Build the mapping
  const entries: string[] = [];
  let matched = 0;

  for (const [serviceKey, searchTerm] of Object.entries(SERVICE_TO_SVG)) {
    const match = svgFiles.find((f) => {
      const basename = path.basename(f, ".svg").toLowerCase().replace(/[\s_-]+/g, "");
      const searchTokens = searchTerm.toLowerCase().split(/\s+/);
      // All tokens from the search term must appear in the filename
      return searchTokens.every((token) => basename.includes(token));
    });

    if (match) {
      const svgContent = fs.readFileSync(match, "utf-8");
      const base64 = Buffer.from(svgContent, "utf-8").toString("base64");
      const dataUrl = `data:image/svg+xml;base64,${base64}`;

      // Determine category from parent folder name
      const parentDir = path.basename(path.dirname(match));
      const category = parentDir.replace(/^\d+[-_\s]*/, "");

      entries.push(`  "${serviceKey}": { category: "${escapeString(category)}", svgDataUrl: "${dataUrl}" },`);
      matched++;
    } else {
      console.warn(`⚠️  No SVG match for: ${serviceKey} (searched: "${searchTerm}")`);
    }
  }

  console.log(`✅ Matched ${matched}/${Object.keys(SERVICE_TO_SVG).length} services`);

  // Generate TypeScript file
  const tsContent = `// AUTO-GENERATED by scripts/build-azure-icons.ts — do not edit manually.
// Source: ${ICON_ZIP_URL}

import type { AzureIconMap } from "./types.js";

/**
 * Azure service key → embedded SVG data URL mapping.
 * Generated from the official Microsoft Azure Architecture Icons pack.
 */
export const AZURE_ICONS: AzureIconMap = {
${entries.join("\n")}
};
`;

  fs.writeFileSync(OUTPUT_FILE, tsContent, "utf-8");
  console.log(`📝 Generated ${OUTPUT_FILE}`);

  // Cleanup
  fs.rmSync(TMP_DIR, { recursive: true });
  console.log("🧹 Cleaned up temp files");
}

function findSvgFiles(dir: string): string[] {
  const results: string[] = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...findSvgFiles(fullPath));
    } else if (entry.name.endsWith(".svg")) {
      results.push(fullPath);
    }
  }
  return results;
}

function escapeString(s: string): string {
  return s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

main().catch((err) => {
  console.error("❌ Error:", err);
  process.exit(1);
});
