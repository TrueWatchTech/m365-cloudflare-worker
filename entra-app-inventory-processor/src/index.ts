const USER_AGENT = "M365LogCollector/1.0";
const INTERNAL_PING_URL = "https://datakit.internal/v1/ping";
const INTERNAL_POST_URL = "https://datakit.internal/v1/write/metric";

const MAX_POST_BYTES = 4_500_000;
const MAX_LINES_PER_POST = 10_000;
const POST_TIMEOUT_MS = 30000;
const PING_TIMEOUT_MS = 4000;
const PING_WAIT_MS = 1500;

interface Env {
  DATAKIT: Fetcher;
  LOGS_BUCKET: R2Bucket;
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

        if (!objectKey.startsWith("m365/ms_applications/")) {
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

        const collectedMs = Math.floor(tsNs / 1_000_000);
        const object = await env.LOGS_BUCKET.get(objectKey);
        if (!object) {
          console.error(`File not found: ${objectKey}`);
          message.retry();
          continue;
        }

        let appsArray: Array<Record<string, unknown>>;
        try {
          appsArray = JSON.parse(await object.text()) as Array<Record<string, unknown>>;
        } catch {
          console.error("JSON parsing failed");
          message.ack();
          continue;
        }

        if (!Array.isArray(appsArray)) {
          console.error("File is not array");
          message.ack();
          continue;
        }

        console.log(`Apps in file: ${appsArray.length}`);

        const lines: string[] = [];
        let considered = 0;
        let emitted = 0;
        let skippedNoSecrets = 0;
        let skippedExpired = 0;
        let skippedBadDate = 0;

        for (const app of appsArray) {
          const displayName = app?.displayName;
          const passwordCredentials = app?.passwordCredentials;
          considered += 1;

          if (!displayName || !Array.isArray(passwordCredentials) || passwordCredentials.length === 0) {
            skippedNoSecrets += 1;
            continue;
          }

          const latest = pickLatestPasswordCredential(passwordCredentials);
          if (!latest?.endDateTime) {
            skippedBadDate += 1;
            continue;
          }

          const expiryMs = Date.parse(latest.endDateTime);
          if (!Number.isFinite(expiryMs)) {
            skippedBadDate += 1;
            continue;
          }

          if (expiryMs <= collectedMs) {
            skippedExpired += 1;
            continue;
          }

          const expiryTimeSeconds = Math.floor((expiryMs - collectedMs) / 1000);
          const tagApp = influxEscapeTagValue(displayName);
          lines.push(`entra_app_secret_expiry,app=${tagApp} expiry_time_seconds=${expiryTimeSeconds}i ${tsNs}`);
          emitted += 1;
        }

        console.log(
          `Considered=${considered}, emitted=${emitted}, skippedNoSecrets=${skippedNoSecrets}, skippedExpired=${skippedExpired}, skippedBadDate=${skippedBadDate}`,
        );

        if (lines.length === 0) {
          console.log("No eligible metrics to send");
          console.log("===========================================");
          message.ack();
          continue;
        }

        const source = "m365";
        const service = "entra_secret_expiry";

        await pingTrueWatchWithRetries(env, source, service);

        const chunks = chunkLines(lines, {
          maxBytes: MAX_POST_BYTES,
          maxLines: MAX_LINES_PER_POST,
        });

        for (const chunk of chunks) {
          const payload = `${chunk.join("\n")}\n`;
          const resp = await sendToTrueWatchMetrics(payload, env);
          console.log(`TrueWatch response: ${JSON.stringify(resp)}`);
        }

        console.log(`Successfully sent ${lines.length} metrics lines`);
        console.log("===========================================");
        message.ack();
      } catch (error) {
        console.error("Fatal error:", error);
        message.retry();
      }
    }
  },
};

function pickLatestPasswordCredential(passwordCredentials: unknown[]) {
  let best: Record<string, unknown> | null = null;
  let bestMs = -Infinity;

  for (const credential of passwordCredentials) {
    const end = (credential as { endDateTime?: string })?.endDateTime;
    if (!end) continue;

    const ms = Date.parse(end);
    if (!Number.isFinite(ms)) continue;

    if (ms > bestMs) {
      bestMs = ms;
      best = credential as Record<string, unknown>;
    }
  }

  return best as { endDateTime?: string } | null;
}

function parseTimestampNsFromObjectKey(objectKey: string) {
  const filename = objectKey.split("/").pop() || "";
  const match = filename.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})-(\d{3})Z(?:\.\w+)?$/);
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
