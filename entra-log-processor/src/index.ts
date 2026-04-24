const DEDUP_TTL_SECONDS = 24 * 3600;
const USER_AGENT = "M365LogCollector/1.0";
const INTERNAL_PING_URL = "https://datakit.internal/v1/ping";
const INTERNAL_POST_URL = "https://datakit.internal/v1/write/logstreaming";
const OBJECT_PREFIX = "m365/entra/";
const SOURCE = "m365";
const SERVICE_BASE = "entra";

type R2Notification = {
  object?: {
    key?: string;
  };
};

type RawEnvelope = {
  tenantLabel: string;
  logType: "audit" | "signin";
  logs: Array<Record<string, unknown>>;
};

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
        const objectKey = message.body?.object?.key;
        if (!objectKey) {
          console.error("Skipping queue message without object key");
          message.ack();
          continue;
        }

        if (!objectKey.startsWith(OBJECT_PREFIX)) {
          console.log(`Skipping unknown prefix: ${objectKey}`);
          message.ack();
          continue;
        }

        const object = await env.LOGS_BUCKET.get(objectKey);
        if (!object) {
          console.error(`R2 object not found: ${objectKey}`);
          message.retry();
          continue;
        }

        let payload: RawEnvelope;
        try {
          payload = JSON.parse(await object.text()) as RawEnvelope;
        } catch (error) {
          console.error(`Failed to parse R2 object ${objectKey}:`, error);
          message.ack();
          continue;
        }

        const logsArray = Array.isArray(payload.logs) ? payload.logs : [];
        const service = getServiceForObjectKey(objectKey);
        if (!service) {
          console.error(`[${payload.tenantLabel}] Unknown log type in path: ${objectKey}`);
          message.ack();
          continue;
        }

        console.log("===========================================");
        console.log(`[${payload.tenantLabel}] Processing ${payload.logType} R2 object ${objectKey} with ${logsArray.length} logs`);

        if (logsArray.length === 0) {
          console.log(`[${payload.tenantLabel}] Empty batch; acknowledging`);
          console.log("===========================================");
          message.ack();
          continue;
        }

        const { uniqueLogs, idsToMark } = await deduplicateLogsReadOnly(
          logsArray,
          env.LOG_IDS,
          payload.tenantLabel,
        );

        if (uniqueLogs.length === 0) {
          console.log(`[${payload.tenantLabel}] All logs were duplicates; nothing to send`);
          console.log("===========================================");
          message.ack();
          continue;
        }

        console.log(
          `[${payload.tenantLabel}] Unique logs: ${uniqueLogs.length} (removed ${logsArray.length - uniqueLogs.length} duplicates)`,
        );

        const ndjsonContent = uniqueLogs.map((log) => JSON.stringify(log)).join("\n");
        console.log(
          `[${payload.tenantLabel}] Sending logs: source=${SOURCE}, service=${service}`,
        );

        const truewatchResponse = await sendToTrueWatch(
          ndjsonContent,
          SOURCE,
          service,
          env,
        );

        console.log(`[${payload.tenantLabel}] TrueWatch response: ${JSON.stringify(truewatchResponse)}`);

        await markLogsProcessed(idsToMark, env.LOG_IDS, payload.tenantLabel);

        console.log(`[${payload.tenantLabel}] Successfully sent and marked ${uniqueLogs.length} logs`);
        console.log("===========================================");
        message.ack();
      } catch (error) {
        console.error("Fatal error:", error);
        message.retry();
      }
    }
  },
};

async function deduplicateLogsReadOnly(
  logsArray: Array<Record<string, unknown>>,
  kv: KVNamespace,
  tenantLabel: string,
) {
  const uniqueLogs: Array<Record<string, unknown>> = [];
  const idsToMark: string[] = [];

  for (const log of logsArray) {
    const logId = typeof log.id === "string" ? log.id : null;
    if (!logId) {
      uniqueLogs.push(log);
      continue;
    }

    const kvKey = `${tenantLabel}:${logId}`;
    const seen = await kv.get(kvKey);

    if (!seen) {
      uniqueLogs.push(log);
      idsToMark.push(logId);
    }
  }

  return { uniqueLogs, idsToMark };
}

async function markLogsProcessed(idsToMark: string[], kv: KVNamespace, tenantLabel: string) {
  let count = 0;

  for (const id of idsToMark) {
    const kvKey = `${tenantLabel}:${id}`;
    try {
      await kv.put(kvKey, "1", { expirationTtl: DEDUP_TTL_SECONDS });
      count += 1;
    } catch (error) {
      console.error(`[${tenantLabel}] KV put failed (${kvKey}):`, error);
    }
  }

  console.log(`[${tenantLabel}] KV marked ${count}/${idsToMark.length} IDs`);
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
        console.log(`INFO: Ping succeeded on attempt ${attempt} (HTTP ${res.status})`);
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

async function sendToTrueWatch(
  ndjsonContent: string,
  source: string,
  service: string,
  env: Env,
) {
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

function getServiceForObjectKey(objectKey: string): string | null {
  if (objectKey.includes("/audit/")) {
    return `${SERVICE_BASE}_audit`;
  }
  if (objectKey.includes("/signin/")) {
    return `${SERVICE_BASE}_signin`;
  }
  return null;
}
