// =========================
// CONFIGURATION
// =========================

const BUFFER_MINUTES = 15;
const INITIAL_LOOKBACK_MINUTES = 60;
const FETCH_TIMEOUT_MS = 25_000;
const FETCH_RETRIES = 3;
const PAGE_TOP = 1000;
const USER_AGENT = "M365LogCollector/1.0";
const R2_PREFIX = "m365/entra";

type RawEnvelope = {
  tenantLabel: string;
  logType: "audit" | "signin";
  logs: unknown[];
};

interface Env {
  CLIENT_SECRET: string;
  CLIENT_ID: string;
  LOG_CURSORS: KVNamespace;
  LOGS_BUCKET: R2Bucket;
  R2_BATCH_SIZE: string;
  TENANT_ID: string;
  TENANT_LABEL: string;
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

  async scheduled(event: ScheduledController, env: Env): Promise<void> {
    console.log("Starting Entra ID log collection...");
    console.log(`cron=${event.cron} scheduledTime=${new Date(event.scheduledTime).toISOString()}`);

    try {
      await collectTenantLogs({
        tenantLabel: env.TENANT_LABEL,
        tenantId: env.TENANT_ID,
        clientId: env.CLIENT_ID,
        clientSecret: env.CLIENT_SECRET,
        cursorsKV: env.LOG_CURSORS,
        logsBucket: env.LOGS_BUCKET,
        objectBatchSize: Number(env.R2_BATCH_SIZE || "1000"),
      });
      console.log("Entra ID log collection completed successfully");
    } catch (error) {
      console.error("Entra ID log collection failed:", error);
      throw error;
    }
  },
};

async function collectTenantLogs({
  tenantLabel,
  tenantId,
  clientId,
  clientSecret,
  cursorsKV,
  logsBucket,
  objectBatchSize,
}: {
  tenantLabel: string;
  tenantId: string;
  clientId: string;
  clientSecret: string;
  cursorsKV: KVNamespace;
  logsBucket: R2Bucket;
  objectBatchSize: number;
}) {
  console.log(`=== Collecting logs for ${tenantLabel} (${tenantId}) ===`);

  const token = await getAccessToken({ tenantId, clientId, clientSecret, tenantLabel });

  const audit = await fetchLogsWithCursor({
    token,
    baseUrl: "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits",
    dateProperty: "activityDateTime",
    cursorKey: `${tenantLabel}_entra_audit_cursor`,
    cursorsKV,
    tenantLabel,
    logType: "audit",
  });

  if (audit.logs.length) {
    await storeLogs(logsBucket, objectBatchSize, {
      tenantLabel,
      logType: "audit",
      logs: audit.logs,
    });
    await updateCursorIfPresent(cursorsKV, audit.cursorKey, audit.newest, tenantLabel, "audit");
  } else {
    console.log(`[${tenantLabel}/audit] No logs to enqueue; cursor unchanged.`);
  }

  const signin = await fetchLogsWithCursor({
    token,
    baseUrl: "https://graph.microsoft.com/v1.0/auditLogs/signIns",
    dateProperty: "createdDateTime",
    cursorKey: `${tenantLabel}_entra_signin_cursor`,
    cursorsKV,
    tenantLabel,
    logType: "signin",
  });

  if (signin.logs.length) {
    await storeLogs(logsBucket, objectBatchSize, {
      tenantLabel,
      logType: "signin",
      logs: signin.logs,
    });
    await updateCursorIfPresent(cursorsKV, signin.cursorKey, signin.newest, tenantLabel, "signin");
  } else {
    console.log(`[${tenantLabel}/signin] No logs to enqueue; cursor unchanged.`);
  }

  console.log(`=== Completed ${tenantLabel} ===`);
}

async function fetchWithRetry(
  url: string,
  options: RequestInit = {},
  { label = "fetch", retries = FETCH_RETRIES, timeoutMs = FETCH_TIMEOUT_MS } = {},
) {
  const host = (() => {
    try {
      return new URL(url).host;
    } catch {
      return "invalid-url";
    }
  })();

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const start = Date.now();

    try {
      console.log(`[${label}] -> ${host} (attempt=${attempt})`);
      const res = await fetch(url, { ...options, signal: controller.signal });
      clearTimeout(timer);

      const dur = Date.now() - start;
      console.log(`[${label}] <- ${host} status=${res.status} ms=${dur}`);

      if (res.status === 429 && attempt < retries) {
        const ra = Number(res.headers.get("Retry-After") || "0");
        const waitMs = ra > 0 ? ra * 1000 : 500 * 2 ** attempt + Math.floor(Math.random() * 250);
        console.warn(`[${label}] 429 throttled. waiting ${waitMs}ms then retry...`);
        await sleep(waitMs);
        continue;
      }

      if ([500, 502, 503, 504].includes(res.status) && attempt < retries) {
        const waitMs = 500 * 2 ** attempt + Math.floor(Math.random() * 250);
        console.warn(`[${label}] ${res.status} transient. waiting ${waitMs}ms then retry...`);
        await sleep(waitMs);
        continue;
      }

      return res;
    } catch (error) {
      clearTimeout(timer);
      const dur = Date.now() - start;

      if (attempt >= retries) {
        console.error(`[${label}] !! ${host} ms=${dur} final_error=${String((error as Error)?.message || error)}`);
        throw error;
      }

      const waitMs = 500 * 2 ** attempt + Math.floor(Math.random() * 250);
      console.warn(
        `[${label}] !! ${host} ms=${dur} error=${String((error as Error)?.message || error)}. waiting ${waitMs}ms then retry...`,
      );
      await sleep(waitMs);
    }
  }

  throw new Error(`${label}: exhausted retries unexpectedly`);
}

async function getAccessToken({
  tenantId,
  clientId,
  clientSecret,
  tenantLabel,
}: {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  tenantLabel: string;
}) {
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const res = await fetchWithRetry(
    tokenUrl,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "User-Agent": USER_AGENT,
      },
      body,
    },
    { label: `${tenantLabel}/token` },
  );

  if (!res.ok) {
    const err = await res.text().catch(() => "<no-body>");
    throw new Error(`Token request failed for tenant ${tenantId}: HTTP ${res.status} ${err}`);
  }

  const json = (await res.json()) as { access_token?: string };
  if (!json.access_token) {
    throw new Error(`Token response missing access_token for tenant ${tenantId}`);
  }

  return json.access_token;
}

async function fetchLogsWithCursor({
  token,
  baseUrl,
  dateProperty,
  cursorKey,
  cursorsKV,
  tenantLabel,
  logType,
}: {
  token: string;
  baseUrl: string;
  dateProperty: string;
  cursorKey: string;
  cursorsKV: KVNamespace;
  tenantLabel: string;
  logType: "audit" | "signin";
}) {
  const now = Date.now();
  const cursor = await cursorsKV.get(cursorKey);

  const fromTime = !cursor
    ? new Date(now - INITIAL_LOOKBACK_MINUTES * 60 * 1000).toISOString()
    : new Date(new Date(cursor).getTime() - BUFFER_MINUTES * 60 * 1000).toISOString();

  if (!cursor) {
    console.log(`[${tenantLabel}/${logType}] No cursor. Using initial lookback from ${fromTime}`);
  } else {
    console.log(`[${tenantLabel}/${logType}] Using cursor ${cursor}, fetching from ${fromTime}`);
  }

  const logs = await fetchLogs(token, baseUrl, dateProperty, fromTime, `${tenantLabel}/${logType}`);
  if (!logs.length) {
    console.log(`[${tenantLabel}/${logType}] No logs returned from Graph.`);
    return { logs: [], newest: null, cursorKey };
  }

  const newest = (logs[logs.length - 1] as Record<string, unknown> | undefined)?.[dateProperty] ?? null;
  return { logs, newest: typeof newest === "string" ? newest : null, cursorKey };
}

async function updateCursorIfPresent(
  cursorsKV: KVNamespace,
  cursorKey: string,
  newest: string | null,
  tenantLabel: string,
  logType: string,
) {
  if (!newest) return;
  await cursorsKV.put(cursorKey, newest);
  console.log(`[${tenantLabel}/${logType}] Cursor updated -> ${newest}`);
}

async function fetchLogs(
  token: string,
  baseUrl: string,
  dateProperty: string,
  fromTime: string,
  labelBase: string,
) {
  let nextLink: string | null = buildInitialGraphUrl(baseUrl, dateProperty, fromTime);
  const collected: unknown[] = [];
  let page = 0;

  console.log(`[${labelBase}] Fetching from ${nextLink}`);

  while (nextLink) {
    page += 1;

    const res = await fetchWithRetry(
      nextLink,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          "User-Agent": USER_AGENT,
        },
      },
      { label: `${labelBase}/graph_page_${page}` },
    );

    if (!res.ok) {
      const body = await res.text().catch(() => "<no-body>");
      throw new Error(`Graph fetch error HTTP ${res.status}: ${body}`);
    }

    const data = (await res.json()) as { value?: unknown[]; "@odata.nextLink"?: string };
    if (Array.isArray(data.value)) {
      collected.push(...data.value);
    }

    nextLink = data["@odata.nextLink"] || null;
  }

  collected.sort((a, b) => {
    const A = String((a as Record<string, unknown>)[dateProperty] || "");
    const B = String((b as Record<string, unknown>)[dateProperty] || "");
    return A < B ? -1 : A > B ? 1 : 0;
  });

  console.log(`[${labelBase}] Fetched ${collected.length} logs from Graph.`);
  return collected;
}

function buildInitialGraphUrl(baseUrl: string, dateProperty: string, fromTime: string) {
  const u = new URL(baseUrl);
  u.searchParams.set("$filter", `${dateProperty} ge ${fromTime}`);
  u.searchParams.set("$top", String(PAGE_TOP));
  return u.toString();
}

async function storeLogs(
  bucket: R2Bucket,
  objectBatchSize: number,
  envelope: RawEnvelope,
) {
  const chunks = chunkArray(envelope.logs, objectBatchSize);
  console.log(
    `[${envelope.tenantLabel}/${envelope.logType}] Writing ${envelope.logs.length} logs to R2 in ${chunks.length} object(s)`,
  );

  for (let index = 0; index < chunks.length; index += 1) {
    const logs = chunks[index];
    const objectKey = buildObjectKey(envelope.logType, index);
    await bucket.put(objectKey, JSON.stringify({
      ...envelope,
      logs,
      collectedAt: new Date().toISOString(),
      objectKey,
    }));
    console.log(
      `[${envelope.tenantLabel}/${envelope.logType}] Stored ${logs.length} logs at ${objectKey}`,
    );
  }
}

function buildObjectKey(logType: "audit" | "signin", index: number) {
  const timestamp = new Date().toISOString().replaceAll(":", "-");
  return `${R2_PREFIX}/${logType}/${timestamp}-part-${String(index).padStart(4, "0")}.json`;
}

function chunkArray<T>(value: T[], size: number): T[][] {
  const result: T[][] = [];
  for (let index = 0; index < value.length; index += size) {
    result.push(value.slice(index, index + size));
  }
  return result;
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
