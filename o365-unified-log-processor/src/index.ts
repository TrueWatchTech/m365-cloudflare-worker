const DEDUP_TTL_SECONDS = 48 * 3600;
const USER_AGENT = "M365LogCollector/1.0";
const INTERNAL_PING_URL = "https://datakit.internal/v1/ping";
const INTERNAL_POST_URL = "https://datakit.internal/v1/write/logstreaming";

// Update tenantLabel and tenantPrefix to match your TENANT_LABEL value in wrangler.jsonc
const ROUTE_CONFIG = [
  {
    prefix: "m365/exchange/",
    tenantLabel: "YOUR_TENANT_LABEL",
    tenantPrefix: "YOUR_TENANT_LABEL",
    source: "m365",
    service: "o365_exchange",
  },
  {
    prefix: "m365/sharepoint/",
    tenantLabel: "YOUR_TENANT_LABEL",
    tenantPrefix: "YOUR_TENANT_LABEL",
    source: "m365",
    service: "o365_sharepoint",
  },
  {
    prefix: "m365/general/",
    tenantLabel: "YOUR_TENANT_LABEL",
    tenantPrefix: "YOUR_TENANT_LABEL",
    source: "m365",
    service: "o365_general",
  },
  {
    prefix: "m365/azureactivedirectory/",
    tenantLabel: "YOUR_TENANT_LABEL",
    tenantPrefix: "YOUR_TENANT_LABEL",
    source: "m365",
    service: "o365_activedirectory",
  },
  {
    prefix: "m365/dlp/",
    tenantLabel: "YOUR_TENANT_LABEL",
    tenantPrefix: "YOUR_TENANT_LABEL",
    source: "m365",
    service: "o365_dlp",
  },
] as const;

type R2Notification = {
  object?: {
    key?: string;
  };
};

type RouteConfig = (typeof ROUTE_CONFIG)[number];

interface Env {
  DATAKIT: Fetcher;
  LOG_IDS: KVNamespace;
  LOGS_BUCKET: R2Bucket;
}

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
        const notification = message.body;
        const objectKey = notification?.object?.key;
        if (!objectKey) {
          console.log("Skipping message without object.key:", JSON.stringify(notification));
          message.ack();
          continue;
        }

        const route = findRouteConfig(objectKey);
        if (!route) {
          console.log("Skipping unknown prefix:", objectKey);
          message.ack();
          continue;
        }

        console.log("===========================================");
        console.log(`[${route.tenantLabel}] Processing R2 log file: ${objectKey}`);

        const object = await env.LOGS_BUCKET.get(objectKey);
        if (!object) {
          console.error(`[${route.tenantLabel}] File not found: ${objectKey}`);
          message.retry();
          continue;
        }

        let logsArray: Array<Record<string, unknown>>;
        try {
          logsArray = JSON.parse(await object.text()) as Array<Record<string, unknown>>;
        } catch (error) {
          console.error(`[${route.tenantLabel}] JSON parsing failed:`, error);
          message.ack();
          continue;
        }

        if (!Array.isArray(logsArray)) {
          console.error(`[${route.tenantLabel}] File is not a log array`);
          message.ack();
          continue;
        }

        console.log(`[${route.tenantLabel}] Logs in file: ${logsArray.length}`);

        const { uniqueLogs, idsToMark, duplicateCount } = await deduplicateLogsReadOnly(
          logsArray,
          env.LOG_IDS,
          route.tenantPrefix,
          route.service,
          route.tenantLabel,
        );

        if (uniqueLogs.length === 0) {
          console.log(`[${route.tenantLabel}] All logs were duplicates; nothing to send`);
          console.log("===========================================");
          message.ack();
          continue;
        }

        console.log(
          `[${route.tenantLabel}] Unique logs: ${uniqueLogs.length} (removed ${duplicateCount} duplicates)`,
        );

        uniqueLogs.sort(compareByCreationTimeAsc);
        const ndjsonContent = uniqueLogs.map((log) => JSON.stringify(log)).join("\n");
        console.log(
          `[${route.tenantLabel}] Sending logs: source=${route.source}, service=${route.service}`,
        );

        const truewatchResponse = await sendToTrueWatch(
          ndjsonContent,
          route.source,
          route.service,
          env,
        );

        console.log(`[${route.tenantLabel}] TrueWatch response: ${JSON.stringify(truewatchResponse)}`);

        await markLogsProcessed(
          idsToMark,
          env.LOG_IDS,
          route.tenantPrefix,
          route.service,
          route.tenantLabel,
        );

        console.log(`[${route.tenantLabel}] Successfully sent and marked ${uniqueLogs.length} logs`);
        console.log("===========================================");
        message.ack();
      } catch (error) {
        console.error("Fatal error:", error);
        message.retry();
      }
    }
  },
};

function findRouteConfig(objectKey: string): RouteConfig | null {
  return ROUTE_CONFIG.find((route) => objectKey.startsWith(route.prefix)) || null;
}

async function deduplicateLogsReadOnly(
  logsArray: Array<Record<string, unknown>>,
  kv: KVNamespace,
  tenantPrefix: string,
  service: string,
  tenantLabel: string,
) {
  const uniqueLogs: Array<Record<string, unknown>> = [];
  const idsToMark: string[] = [];
  let duplicateCount = 0;

  for (const log of logsArray) {
    const logId = typeof log.Id === "string" ? log.Id : null;
    if (!logId) {
      uniqueLogs.push(log);
      continue;
    }

    const kvKey = buildDedupKey(tenantPrefix, service, logId);
    const seen = await kv.get(kvKey);
    if (!seen) {
      uniqueLogs.push(log);
      idsToMark.push(logId);
    } else {
      duplicateCount += 1;
    }
  }

  if (duplicateCount > 0) {
    console.log(`[${tenantLabel}] Dedup skipped ${duplicateCount} logs already marked in KV`);
  }

  return { uniqueLogs, idsToMark, duplicateCount };
}

async function markLogsProcessed(
  idsToMark: string[],
  kv: KVNamespace,
  tenantPrefix: string,
  service: string,
  tenantLabel: string,
) {
  let count = 0;

  for (const id of idsToMark) {
    const kvKey = buildDedupKey(tenantPrefix, service, id);
    try {
      await kv.put(kvKey, "1", { expirationTtl: DEDUP_TTL_SECONDS });
      count += 1;
    } catch (error) {
      console.error(`[${tenantLabel}] KV put failed (${kvKey}):`, error);
    }
  }

  console.log(`[${tenantLabel}] KV marked ${count}/${idsToMark.length} IDs`);
}

function buildDedupKey(tenantPrefix: string, service: string, id: string) {
  return `${tenantPrefix}:${service}:${id}`;
}

async function pingTrueWatchWithRetries(env: Env, source: string, service: string) {
  const maxAttempts = 5;
  const intervalMs = 1500;
  const timeoutMs = 4000;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

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

      if (!res.ok) {
        console.warn(`WARN: Ping attempt ${attempt} failed (HTTP ${res.status})`);
      } else {
        let metadata = "";
        try {
          const body = (await res.json()) as { content?: { version?: string; uptime?: string } };
          if (body?.content) {
            metadata = ` (version: ${body.content.version}, uptime: ${body.content.uptime})`;
          }
        } catch {
        }

        console.log(`INFO: Ping succeeded on attempt ${attempt} (HTTP ${res.status})${metadata}`);
        return true;
      }
    } catch (error) {
      console.warn(`WARN: Ping attempt ${attempt} error: ${String((error as Error)?.message || error)}`);
    }

    await sleep(intervalMs);
  }

  console.warn(`WARN: Ping did not succeed after ${maxAttempts} attempts; proceeding to POST once anyway`);
  return false;
}

async function sendToTrueWatch(ndjsonContent: string, source: string, service: string, env: Env) {
  await pingTrueWatchWithRetries(env, source, service);

  console.log(`Posting ${ndjsonContent.length} bytes to DataKit...`);
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);

  const res = await env.DATAKIT.fetch(
    new Request(
      `${INTERNAL_POST_URL}?source=${encodeURIComponent(source)}&service=${encodeURIComponent(service)}`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-ndjson",
          "User-Agent": USER_AGENT,
        },
        body: ndjsonContent,
        signal: controller.signal,
      },
    ),
  );

  clearTimeout(timeout);

  if (!res.ok) {
    const body = await res.text().catch(() => "<no-body>");
    throw new Error(`DataKit POST failed (${res.status}): ${body}`);
  }

  return res.json().catch(() => ({ message: "<non-json-response>" }));
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function compareByCreationTimeAsc(a: Record<string, unknown>, b: Record<string, unknown>) {
  const aTime = parseCreationTime(a);
  const bTime = parseCreationTime(b);

  if (aTime === null && bTime === null) return 0;
  if (aTime === null) return 1;
  if (bTime === null) return -1;
  return aTime - bTime;
}

function parseCreationTime(log: Record<string, unknown>) {
  const value = log.CreationTime;
  if (typeof value !== "string") return null;
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? null : parsed;
}
