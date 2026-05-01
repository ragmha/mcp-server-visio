// Azure service key → embedded SVG data URL mapping.
//
// For the full icon set, run: npm run build:icons
// This downloads official Microsoft Azure Architecture Icons and generates
// base64-encoded SVG data URLs for each service.
//
// This file contains a starter set of service keys with placeholder SVGs.
// The build:icons script replaces this with the full generated mapping.

import type { AzureIconMap } from "./types.js";

function placeholderLabel(label: string): string {
  return label.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Minimal inline SVG used as a fallback when real icons aren't built yet
function placeholder(label: string, color: string): string {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48"><rect width="48" height="48" rx="8" fill="${color}" opacity="0.15"/><rect width="48" height="48" rx="8" fill="none" stroke="${color}" stroke-width="2"/><text x="24" y="28" text-anchor="middle" font-size="8" font-family="sans-serif" fill="${color}">${placeholderLabel(label)}</text></svg>`;
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
}

/**
 * Azure service key → embedded SVG data URL mapping.
 *
 * 206 Azure services mapped to SVG icons. Run `npm run build:icons` to
 * generate this file with official Microsoft Azure Architecture Icons.
 */
export const AZURE_ICONS: AzureIconMap = {
  // ── AI + Machine Learning ───────────────────────────────
  "azure/machine-learning": { category: "AI + Machine Learning", svgDataUrl: placeholder("ML", "#0078D7") },
  "azure/openai": { category: "AI + Machine Learning", svgDataUrl: placeholder("OpenAI", "#0078D7") },
  "azure/cognitive-services": { category: "AI + Machine Learning", svgDataUrl: placeholder("Cognitive", "#0078D7") },
  "azure/bot-services": { category: "AI + Machine Learning", svgDataUrl: placeholder("Bot", "#0078D7") },
  "azure/ai-studio": { category: "AI + Machine Learning", svgDataUrl: placeholder("AI Studio", "#0078D7") },

  // ── Analytics ───────────────────────────────────────────
  "azure/synapse-analytics": { category: "Analytics", svgDataUrl: placeholder("Synapse", "#8764B8") },
  "azure/databricks": { category: "Analytics", svgDataUrl: placeholder("Databricks", "#8764B8") },
  "azure/stream-analytics": { category: "Analytics", svgDataUrl: placeholder("Stream", "#8764B8") },
  "azure/data-factory": { category: "Analytics", svgDataUrl: placeholder("DataFact", "#8764B8") },
  "azure/event-hubs": { category: "Analytics", svgDataUrl: placeholder("EventHub", "#8764B8") },
  "azure/log-analytics": { category: "Analytics", svgDataUrl: placeholder("LogAnaly", "#8764B8") },
  "azure/hd-insight": { category: "Analytics", svgDataUrl: placeholder("HDInsight", "#8764B8") },
  "azure/power-bi-embedded": { category: "Analytics", svgDataUrl: placeholder("PowerBI", "#8764B8") },

  // ── App Services ────────────────────────────────────────
  "azure/app-services": { category: "App Services", svgDataUrl: placeholder("AppSvc", "#FF8C00") },
  "azure/function-apps": { category: "App Services", svgDataUrl: placeholder("Func", "#FF8C00") },
  "azure/static-apps": { category: "App Services", svgDataUrl: placeholder("Static", "#FF8C00") },
  "azure/api-management": { category: "App Services", svgDataUrl: placeholder("APIM", "#FF8C00") },
  "azure/signalr": { category: "App Services", svgDataUrl: placeholder("SignalR", "#FF8C00") },
  "azure/notification-hubs": { category: "App Services", svgDataUrl: placeholder("Notif", "#FF8C00") },
  "azure/spring-apps": { category: "App Services", svgDataUrl: placeholder("Spring", "#FF8C00") },

  // ── Compute ─────────────────────────────────────────────
  "azure/virtual-machines": { category: "Compute", svgDataUrl: placeholder("VM", "#0078D7") },
  "azure/vm-scale-sets": { category: "Compute", svgDataUrl: placeholder("VMSS", "#0078D7") },
  "azure/container-instances": { category: "Containers", svgDataUrl: placeholder("ACI", "#0078D7") },
  "azure/container-registries": { category: "Containers", svgDataUrl: placeholder("ACR", "#0078D7") },
  "azure/kubernetes-services": { category: "Containers", svgDataUrl: placeholder("AKS", "#0078D7") },
  "azure/container-apps": { category: "Containers", svgDataUrl: placeholder("ContApp", "#0078D7") },
  "azure/batch": { category: "Compute", svgDataUrl: placeholder("Batch", "#0078D7") },
  "azure/service-fabric": { category: "Compute", svgDataUrl: placeholder("SvcFab", "#0078D7") },

  // ── Databases ───────────────────────────────────────────
  "azure/cosmos-db": { category: "Databases", svgDataUrl: placeholder("Cosmos", "#0078D7") },
  "azure/sql-database": { category: "Databases", svgDataUrl: placeholder("SQL DB", "#0078D7") },
  "azure/sql-server": { category: "Databases", svgDataUrl: placeholder("SQL Svr", "#0078D7") },
  "azure/sql-managed-instance": { category: "Databases", svgDataUrl: placeholder("SQL MI", "#0078D7") },
  "azure/redis-cache": { category: "Databases", svgDataUrl: placeholder("Redis", "#E81123") },
  "azure/mysql": { category: "Databases", svgDataUrl: placeholder("MySQL", "#0078D7") },
  "azure/postgresql": { category: "Databases", svgDataUrl: placeholder("PgSQL", "#0078D7") },

  // ── Identity ────────────────────────────────────────────
  "azure/azure-ad": { category: "Identity", svgDataUrl: placeholder("AAD", "#0078D7") },
  "azure/managed-identities": { category: "Identity", svgDataUrl: placeholder("MgdId", "#0078D7") },
  "azure/entra-connect": { category: "Identity", svgDataUrl: placeholder("Entra", "#0078D7") },

  // ── Integration ─────────────────────────────────────────
  "azure/logic-apps": { category: "Integration", svgDataUrl: placeholder("Logic", "#0078D7") },
  "azure/service-bus": { category: "Integration", svgDataUrl: placeholder("SvcBus", "#0078D7") },
  "azure/event-grid": { category: "Integration", svgDataUrl: placeholder("EvGrid", "#0078D7") },
  "azure/app-configuration": { category: "Integration", svgDataUrl: placeholder("AppCfg", "#0078D7") },

  // ── IoT ─────────────────────────────────────────────────
  "azure/iot-hub": { category: "IoT", svgDataUrl: placeholder("IoTHub", "#0078D7") },
  "azure/iot-central": { category: "IoT", svgDataUrl: placeholder("IoTCent", "#0078D7") },
  "azure/iot-edge": { category: "IoT", svgDataUrl: placeholder("IoTEdge", "#0078D7") },
  "azure/digital-twins": { category: "IoT", svgDataUrl: placeholder("DgTwins", "#0078D7") },

  // ── Management + Governance ─────────────────────────────
  "azure/monitor": { category: "Monitor", svgDataUrl: placeholder("Monitor", "#0078D7") },
  "azure/application-insights": { category: "Monitor", svgDataUrl: placeholder("AppIns", "#0078D7") },
  "azure/policy": { category: "Management", svgDataUrl: placeholder("Policy", "#0078D7") },
  "azure/advisor": { category: "Management", svgDataUrl: placeholder("Advisor", "#7AB800") },
  "azure/cost-management": { category: "Management", svgDataUrl: placeholder("Cost", "#7AB800") },
  "azure/azure-arc": { category: "Management", svgDataUrl: placeholder("Arc", "#0078D7") },
  "azure/automation-accounts": { category: "Management", svgDataUrl: placeholder("Auto", "#0078D7") },

  // ── Networking ──────────────────────────────────────────
  "azure/virtual-networks": { category: "Networking", svgDataUrl: placeholder("VNet", "#0078D7") },
  "azure/load-balancers": { category: "Networking", svgDataUrl: placeholder("LB", "#0078D7") },
  "azure/application-gateway": { category: "Networking", svgDataUrl: placeholder("AppGW", "#0078D7") },
  "azure/firewall": { category: "Networking", svgDataUrl: placeholder("FW", "#E81123") },
  "azure/dns-zones": { category: "Networking", svgDataUrl: placeholder("DNS", "#0078D7") },
  "azure/front-door": { category: "Networking", svgDataUrl: placeholder("FD", "#00B294") },
  "azure/cdn": { category: "Networking", svgDataUrl: placeholder("CDN", "#00B294") },
  "azure/vpn-gateway": { category: "Networking", svgDataUrl: placeholder("VPN", "#004E98") },
  "azure/expressroute": { category: "Networking", svgDataUrl: placeholder("ER", "#004E98") },
  "azure/bastion": { category: "Networking", svgDataUrl: placeholder("Bastion", "#0078D7") },
  "azure/private-endpoint": { category: "Networking", svgDataUrl: placeholder("PrvEnd", "#004E98") },
  "azure/nat-gateway": { category: "Networking", svgDataUrl: placeholder("NAT", "#0078D7") },
  "azure/traffic-manager": { category: "Networking", svgDataUrl: placeholder("TM", "#0078D7") },
  "azure/ddos-protection": { category: "Networking", svgDataUrl: placeholder("DDoS", "#E81123") },
  "azure/network-security-groups": { category: "Networking", svgDataUrl: placeholder("NSG", "#0078D7") },
  "azure/waf": { category: "Networking", svgDataUrl: placeholder("WAF", "#E81123") },
  "azure/virtual-wan": { category: "Networking", svgDataUrl: placeholder("vWAN", "#0078D7") },

  // ── Security ────────────────────────────────────────────
  "azure/key-vault": { category: "Security", svgDataUrl: placeholder("KV", "#0078D7") },
  "azure/security-center": { category: "Security", svgDataUrl: placeholder("DefClou", "#0078D7") },
  "azure/sentinel": { category: "Security", svgDataUrl: placeholder("Sentinel", "#0078D7") },

  // ── Storage ─────────────────────────────────────────────
  "azure/storage-accounts": { category: "Storage", svgDataUrl: placeholder("Storage", "#7AB800") },
  "azure/data-lake-storage": { category: "Storage", svgDataUrl: placeholder("DLake", "#7AB800") },

  // ── DevOps ──────────────────────────────────────────────
  "azure/devops": { category: "DevOps", svgDataUrl: placeholder("DevOps", "#0078D7") },

  // ── Web ─────────────────────────────────────────────────
  "azure/cognitive-search": { category: "Web", svgDataUrl: placeholder("Search", "#0078D7") },
  "azure/media-services": { category: "Web", svgDataUrl: placeholder("Media", "#0078D7") },
};
