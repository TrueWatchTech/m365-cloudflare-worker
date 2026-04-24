const USER_AGENT = "M365LicenseCollector/1.0";
const INTERNAL_PING_URL = "https://datakit.internal/v1/ping";
const INTERNAL_POST_URL = "https://datakit.internal/v1/write/metric";

const MAX_POST_BYTES = 4_500_000;
const MAX_LINES_PER_POST = 10_000;
const POST_TIMEOUT_MS = 30_000;
const PING_TIMEOUT_MS = 4_000;
const PING_WAIT_MS = 1_500;

// SKUs that have artificially high seat counts and are not meaningful to monitor
const SYSTEM_SKUS = new Set([
  "POWERAPPS_DEV",
  "WINDOWS_STORE",
  "FLOW_FREE",
  "TEAMS_FREE",
  "FLOW_FREE_OOTB",
]);

// Microsoft SKU part number → human-readable product name
// Source: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
const SKU_NAMES: Record<string, string> = {
  // Microsoft 365 Enterprise
  SPE_E1: "Microsoft 365 E1",
  SPE_E3: "Microsoft 365 E3",
  SPE_E5: "Microsoft 365 E5",
  SPE_E3_USGOV_GCCHIGH: "Microsoft 365 E3 (GCC High)",
  SPE_E5_USGOV_GCCHIGH: "Microsoft 365 E5 (GCC High)",
  DEVELOPERPACK_E5: "Microsoft 365 E5 Developer",

  // Microsoft 365 Frontline
  SPE_F1: "Microsoft 365 F1",
  SPE_F3: "Microsoft 365 F3",
  DESKLESSPACK: "Office 365 F3",

  // Microsoft 365 Business
  O365_BUSINESS_ESSENTIALS: "Microsoft 365 Business Basic",
  O365_BUSINESS_PREMIUM: "Microsoft 365 Business Standard",
  SPB: "Microsoft 365 Business Premium",
  O365_BUSINESS: "Microsoft 365 Apps for Business",
  MCOEV_FACULTY: "Microsoft 365 Business Voice",

  // Office 365 Enterprise (legacy)
  STANDARDPACK: "Office 365 E1",
  ENTERPRISEPACK: "Office 365 E3",
  ENTERPRISEPREMIUM: "Office 365 E5",
  ENTERPRISEPREMIUM_NOPSTNCONF: "Office 365 E5 without Audio Conferencing",
  DEVELOPERPACK: "Office 365 E3 Developer",

  // Microsoft Teams
  "M365_TEAMS_PREMIUM": "Microsoft Teams Premium",
  "Teams_Premium_(for_Departments)": "Microsoft Teams Premium (for Departments)",
  MCOEV: "Microsoft Teams Phone Standard",
  MCOMEETADV: "Microsoft Teams Audio Conferencing",
  MCOPSTN1: "Microsoft Teams Domestic Calling Plan",
  MCOPSTN2: "Microsoft Teams International Calling Plan",
  MCOPSTN_5: "Microsoft Teams Domestic Calling Plan (120 min)",
  TEAMS_EXPLORATORY: "Microsoft Teams Exploratory",
  TEAMS_FREE: "Microsoft Teams Free",

  // Security & Compliance
  IDENTITY_THREAT_PROTECTION: "Microsoft 365 E5 Security",
  IDENTITY_THREAT_PROTECTION_FOR_EMS_E5: "Microsoft 365 E5 Security for EMS E5",
  ATP_ENTERPRISE: "Microsoft Defender for Office 365 Plan 2",
  THREAT_INTELLIGENCE: "Microsoft Defender for Office 365 Plan 2 (standalone)",
  ATA: "Microsoft Defender for Identity",
  WINDEFATP: "Microsoft Defender for Endpoint P2",
  MDE_LITE: "Microsoft Defender for Endpoint P1",
  ADALLOM_STANDALONE: "Microsoft Defender for Cloud Apps",
  CAYSMS: "Microsoft Defender for Cloud Apps (GCC)",
  M365_COMPLIANCE_SUITE: "Microsoft 365 E5 Compliance",
  INFORMATION_PROTECTION_COMPLIANCE: "Microsoft 365 E5 Compliance",

  // Enterprise Mobility + Security / Entra / Intune
  EMS: "Enterprise Mobility + Security E3",
  EMSPREMIUM: "Enterprise Mobility + Security E5",
  AAD_PREMIUM: "Microsoft Entra ID P1",
  AAD_PREMIUM_P2: "Microsoft Entra ID P2",
  INTUNE_A: "Microsoft Intune Plan 1",
  INTUNE_A_D: "Microsoft Intune Plan 1 for Device",
  INTUNE_SMB: "Microsoft Intune for Small and Medium Business",

  // Power Platform
  POWERAPPS_DEV: "Power Apps Developer Plan",
  POWERAPPS_PER_USER: "Power Apps Premium",
  POWERAPPS_PER_APP: "Power Apps per App Plan",
  POWER_BI_PRO: "Power BI Pro",
  POWER_BI_PREMIUM_PER_USER: "Power BI Premium Per User",
  POWER_BI_STANDARD: "Power BI (free)",
  FLOW_FREE: "Power Automate Free",
  FLOW_PER_USER: "Power Automate Premium",
  FLOW_PER_USER_DEPT: "Power Automate per User with Attended RPA Plan",

  // Project & Visio
  PROJECTPROFESSIONAL: "Microsoft Project Plan 3",
  PROJECTPREMIUM: "Microsoft Project Plan 5",
  PROJECT_PLAN1: "Microsoft Project Plan 1",
  VISIOCLIENT: "Microsoft Visio Plan 2",
  VISIOONLINE_PLAN1: "Microsoft Visio Plan 1",

  // Windows
  WIN10_PRO_ENT_SUB: "Windows 10/11 Enterprise E3",
  WIN_ENT_E5: "Windows 10/11 Enterprise E5",
  WINDOWS_STORE: "Microsoft Store for Business",

  // Azure Information Protection
  RMS_S_PREMIUM: "Azure Information Protection Premium P1",
  RMS_S_PREMIUM2: "Azure Information Protection Premium P2",

  // Dynamics 365
  DYN365_ENTERPRISE_PLAN1: "Dynamics 365 Customer Engagement Plan",
  DYN365_ENTERPRISE_SALES: "Dynamics 365 Sales Enterprise",
  DYN365_ENTERPRISE_CUSTOMER_SERVICE: "Dynamics 365 Customer Service Enterprise",
  DYN365_ENTERPRISE_FIELD_SERVICE: "Dynamics 365 Field Service",
  DYN365_BUSINESS_EDITION: "Dynamics 365 for Sales, Business Edition",
};

interface Env {
  DATAKIT: Fetcher;
  LOGS_BUCKET: R2Bucket;
}

interface SubscribedSku {
  skuPartNumber: string;
  accountName: string;
  appliesTo: string;
  capabilityStatus: string;
  consumedUnits: number;
  prepaidUnits: {
    enabled: number;
    suspended: number;
    warning: number;
    lockedOut: number;
  };
}

type R2Notification = {
  object?: {
    key?: string;
  };
};

export default {
  async fetch(): Promise<Response> {
    return new Response("", {
      status: 404,
      headers: {
        "Cache-Control": "no-store",
        "X-Content-Type-Options": "nosniff",
      },
    });
  },

  async queue(batch: MessageBatch<R2Notification>, env: Env): Promise<void> {
    for (const message of batch.messages) {
      try {
        const objectKey = message.body?.object?.key;
        if (!objectKey) {
          console.log("Skipping message without object.key");
          message.ack();
          continue;
        }

        if (!objectKey.startsWith("m365/licenses/")) {
          console.log("Skipping unknown prefix:", objectKey);
          message.ack();
          continue;
        }

        console.log("===========================================");
        console.log(`Processing R2 file: ${objectKey}`);

        const tsNs = parseTimestampNsFromObjectKey(objectKey);
        if (!tsNs) {
          console.error(`Could not parse timestamp from object key: ${objectKey}`);
          message.ack();
          continue;
        }

        const object = await env.LOGS_BUCKET.get(objectKey);
        if (!object) {
          console.error(`File not found: ${objectKey}`);
          message.retry();
          continue;
        }

        let skusArray: SubscribedSku[];
        try {
          skusArray = JSON.parse(await object.text()) as SubscribedSku[];
        } catch {
          console.error("JSON parsing failed");
          message.ack();
          continue;
        }

        if (!Array.isArray(skusArray)) {
          console.error("File is not array");
          message.ack();
          continue;
        }

        console.log(`SKUs in file: ${skusArray.length}`);

        const lines: string[] = [];
        let considered = 0;
        let emitted = 0;
        let skippedSystem = 0;

        for (const sku of skusArray) {
          considered += 1;

          const skuPartNumber = sku?.skuPartNumber;
          if (!skuPartNumber) continue;

          // Drop known system/free SKUs with inflated seat counts
          if (SYSTEM_SKUS.has(skuPartNumber)) {
            skippedSystem += 1;
            console.log(`Skipping system SKU: ${skuPartNumber}`);
            continue;
          }

          const total = sku.prepaidUnits?.enabled ?? 0;
          const consumed = sku.consumedUnits ?? 0;
          const available = total - consumed;
          const suspended = sku.prepaidUnits?.suspended ?? 0;
          const warning = sku.prepaidUnits?.warning ?? 0;
          const lockedOut = sku.prepaidUnits?.lockedOut ?? 0;

          const skuName = SKU_NAMES[skuPartNumber] ?? skuPartNumber;
          const tenant = influxEscapeTagValue(sku.accountName ?? "unknown");
          const tagSku = influxEscapeTagValue(skuPartNumber);
          const tagSkuName = influxEscapeTagValue(skuName);
          const tagAppliesTo = influxEscapeTagValue(sku.appliesTo ?? "unknown");
          const tagStatus = influxEscapeTagValue(sku.capabilityStatus ?? "unknown");

          lines.push(
            `ms365_license,tenant=${tenant},sku=${tagSku},sku_name=${tagSkuName},applies_to=${tagAppliesTo},status=${tagStatus}` +
            ` total=${total}i,consumed=${consumed}i,available=${available}i,suspended=${suspended}i,warning=${warning}i,locked_out=${lockedOut}i` +
            ` ${tsNs}`,
          );
          emitted += 1;
        }

        console.log(
          `Considered=${considered}, emitted=${emitted}, skippedSystem=${skippedSystem}`,
        );

        if (lines.length === 0) {
          console.log("No eligible metrics to send");
          console.log("===========================================");
          message.ack();
          continue;
        }

        await pingTrueWatchWithRetries(env, "m365", "ms365_license");

        const chunks = chunkLines(lines, {
          maxBytes: MAX_POST_BYTES,
          maxLines: MAX_LINES_PER_POST,
        });

        for (const chunk of chunks) {
          const payload = `${chunk.join("\n")}\n`;
          const resp = await sendToTrueWatchMetrics(payload, env);
          console.log(`TrueWatch response: ${JSON.stringify(resp)}`);
        }

        console.log(`Successfully sent ${lines.length} metric lines`);
        console.log("===========================================");
        message.ack();
      } catch (error) {
        console.error("Fatal error:", error);
        message.retry();
      }
    }
  },
};

function parseTimestampNsFromObjectKey(objectKey: string): number | null {
  const filename = objectKey.split("/").pop() || "";
  const match = filename.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})-(\d{3})Z(?:\.\w+)?$/,
  );
  if (!match) return null;

  const utcMs = Date.UTC(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4]),
    Number(match[5]),
    Number(match[6]),
    Number(match[7]),
  );

  if (!Number.isFinite(utcMs)) return null;
  return utcMs * 1_000_000;
}

async function pingTrueWatchWithRetries(env: Env, source: string, service: string) {
  const maxAttempts = 5;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), PING_TIMEOUT_MS);

    try {
      const url =
        `${INTERNAL_PING_URL}?source=${encodeURIComponent(source)}` +
        `&service=${encodeURIComponent(service)}` +
        `&attempt=${attempt}`;

      const res = await env.DATAKIT.fetch(
        new Request(url, {
          method: "GET",
          headers: { "User-Agent": USER_AGENT },
          signal: controller.signal,
        }),
      );

      clearTimeout(timeout);
      if (res.ok) {
        console.log(`Ping succeeded (attempt ${attempt})`);
        return true;
      }

      console.warn(`Ping failed attempt ${attempt}`);
    } catch (error) {
      console.warn(`Ping error attempt ${attempt}: ${String((error as Error)?.message || error)}`);
    }

    await sleep(PING_WAIT_MS);
  }

  console.warn("Ping failed after retries");
  return false;
}

async function sendToTrueWatchMetrics(lineProtocolContent: string, env: Env) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), POST_TIMEOUT_MS);

  const res = await env.DATAKIT.fetch(
    new Request(INTERNAL_POST_URL, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain",
        "User-Agent": USER_AGENT,
      },
      body: lineProtocolContent,
      signal: controller.signal,
    }),
  );

  clearTimeout(timeout);

  if (!res.ok) {
    const body = await res.text().catch(() => "<no-body>");
    throw new Error(`DataKit POST failed (${res.status}): ${body}`);
  }

  return res.json().catch(() => ({}));
}

function chunkLines(lines: string[], { maxBytes, maxLines }: { maxBytes: number; maxLines: number }) {
  const chunks: string[][] = [];
  let buf: string[] = [];
  let bytes = 0;

  for (const line of lines) {
    const lineBytes = line.length + 1;

    if (lineBytes > maxBytes) {
      if (buf.length) chunks.push(buf);
      chunks.push([line]);
      buf = [];
      bytes = 0;
      continue;
    }

    if (buf.length + 1 > maxLines || bytes + lineBytes > maxBytes) {
      chunks.push(buf);
      buf = [line];
      bytes = lineBytes;
      continue;
    }

    buf.push(line);
    bytes += lineBytes;
  }

  if (buf.length) chunks.push(buf);
  return chunks;
}

function influxEscapeTagValue(v: unknown) {
  return String(v)
    .replace(/\\/g, "\\\\")
    .replace(/,/g, "\\,")
    .replace(/ /g, "\\ ")
    .replace(/=/g, "\\=");
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
