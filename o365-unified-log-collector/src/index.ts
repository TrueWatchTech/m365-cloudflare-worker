const CONTENT_LOOKBACK_MINUTES = 120;
const FETCH_TIMEOUT_MS = 25_000;
const FETCH_RETRIES = 3;
const MAX_CONTENT_ITEMS_PER_RUN = 1_000;
const CONTENT_ID_TTL_SECONDS = 48 * 3600;
const USER_AGENT = "M365LogCollector/1.0";

const CONTENT_TYPES = [
  { api: "Audit.Exchange", folder: "exchange" },
  { api: "Audit.SharePoint", folder: "sharepoint" },
  { api: "Audit.General", folder: "general" },
  { api: "Audit.AzureActiveDirectory", folder: "azureactivedirectory" },
  { api: "DLP.All", folder: "dlp" },
] as const;

interface Env {
  CLIENT_ID: string;
  CLIENT_SECRET: string;
  LOGS_BUCKET: R2Bucket;
  LOG_CONTENT_IDS: KVNamespace;
  PUBLISHER_ID: string;
  TENANT_ID: string;
  TENANT_LABEL: string;
}

type ContentItem = {
  contentId: string;
  contentUri: string;
  contentCreated?: string;
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

  async scheduled(event: ScheduledController, env: Env): Promise<void> {
    console.log("Starting M365 Unified Audit log collection...");
    console.log(`cron=${event.cron} scheduledTime=${new Date(event.scheduledTime).toISOString()}`);

    const now = new Date(event.scheduledTime || Date.now());
    const windowEnd = now.toISOString();
    const windowStart = new Date(now.getTime() - CONTENT_LOOKBACK_MINUTES * 60 * 1000).toISOString();

    try {
      await collectTenantLogs({
        tenantLabel: env.TENANT_LABEL,
        tenantId: env.TENANT_ID,
        clientId: env.CLIENT_ID,
        clientSecret: env.CLIENT_SECRET,
        bucket: env.LOGS_BUCKET,
        stateKV: env.LOG_CONTENT_IDS,
        publisherId: env.PUBLISHER_ID,
        windowStart,
        windowEnd,
      });
      console.log("M365 Unified Audit collection completed successfully");
    } catch (error) {
      console.error("M365 Unified Audit collection failed:", error);
      throw error;
    }
  },
};

async function collectTenantLogs({
  tenantLabel,
  tenantId,
  clientId,
  clientSecret,
  bucket,
  stateKV,
  publisherId,
  windowStart,
  windowEnd,
}: {
  tenantLabel: string;
  tenantId: string;
  clientId: string;
  clientSecret: string;
  bucket: R2Bucket;
  stateKV: KVNamespace;
  publisherId: string;
  windowStart: string;
  windowEnd: string;
}) {
  console.log(`=== Collecting Unified Audit logs for ${tenantLabel} (${tenantId}) ===`);
  const token = await getAccessToken({ tenantId, clientId, clientSecret, tenantLabel });

  for (const contentType of CONTENT_TYPES) {
    await ensureSubscription({
      token,
      tenantId,
      tenantLabel,
      contentType: contentType.api,
      publisherId,
    });

    const contentItems = await listAvailableContent({
      token,
      tenantId,
      tenantLabel,
      contentType: contentType.api,
      windowStart,
      windowEnd,
      publisherId,
    });

    if (!contentItems.length) {
      console.log(`[${tenantLabel}/${contentType.folder}] No content available.`);
      continue;
    }

    let storedCount = 0;
    let skippedCount = 0;

    for (const item of contentItems.slice(0, MAX_CONTENT_ITEMS_PER_RUN)) {
      const dedupKey = buildContentKey(tenantLabel, contentType.api, item.contentId);
      const alreadyProcessed = await stateKV.get(dedupKey);
      if (alreadyProcessed) {
        skippedCount += 1;
        console.log(
          `[${tenantLabel}/${contentType.folder}] Skipping existing blob contentId=${item.contentId}`,
        );
        continue;
      }

      const content = await fetchContentBlob({
        token,
        tenantLabel,
        contentType: contentType.api,
        contentUri: item.contentUri,
        publisherId,
      });

      const filename = buildBlobFilename({
        folder: contentType.folder,
        contentId: item.contentId,
        contentCreated: item.contentCreated,
      });

      await bucket.put(filename, JSON.stringify(content), {
        httpMetadata: { contentType: "application/json" },
        customMetadata: {
          source: "o365-unified-audit",
          tenant: tenantLabel.toLowerCase(),
          contentType: contentType.api,
          contentId: item.contentId,
          contentCreated: item.contentCreated || "",
        },
      });

      await stateKV.put(dedupKey, "1", { expirationTtl: CONTENT_ID_TTL_SECONDS });
      storedCount += 1;
      console.log(
        `[${tenantLabel}/${contentType.folder}] Stored blob contentId=${item.contentId} -> ${filename}`,
      );
    }

    console.log(
      `[${tenantLabel}/${contentType.folder}] stored=${storedCount} skipped=${skippedCount} listed=${contentItems.length}`,
    );
  }

  console.log(`=== Completed Unified Audit collection for ${tenantLabel} ===`);
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

      if ((res.status === 429 || [500, 502, 503, 504].includes(res.status)) && attempt < retries) {
        const ra = Number(res.headers.get("Retry-After") || "0");
        const waitMs = ra > 0 ? ra * 1000 : 500 * 2 ** attempt + Math.floor(Math.random() * 250);
        console.warn(`[${label}] retryable status=${res.status}; waiting ${waitMs}ms`);
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
    scope: "https://manage.office.com/.default",
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
  if (!json?.access_token) {
    throw new Error(`Token response missing access_token for tenant ${tenantId}`);
  }

  return json.access_token;
}

async function ensureSubscription({
  token,
  tenantId,
  tenantLabel,
  contentType,
  publisherId,
}: {
  token: string;
  tenantId: string;
  tenantLabel: string;
  contentType: string;
  publisherId: string;
}) {
  const url =
    `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/start` +
    `?contentType=${encodeURIComponent(contentType)}` +
    `&PublisherIdentifier=${encodeURIComponent(publisherId)}`;

  const res = await fetchWithRetry(
    url,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "User-Agent": USER_AGENT,
      },
    },
    { label: `${tenantLabel}/${contentType}/subscribe` },
  );

  if (res.status === 200 || res.status === 201) {
    console.log(`[${tenantLabel}/${contentType}] Subscription ensured.`);
    return;
  }

  if (res.status === 400) {
    const text = await res.text().catch(() => "");
    if (text.includes("already enabled") || text.includes("already exists")) {
      console.log(`[${tenantLabel}/${contentType}] Subscription already active.`);
      return;
    }
    throw new Error(`[${tenantLabel}/${contentType}] Subscribe failed: HTTP 400 ${text}`);
  }

  if (!res.ok) {
    const text = await res.text().catch(() => "<no-body>");
    throw new Error(`[${tenantLabel}/${contentType}] Subscribe failed: HTTP ${res.status} ${text}`);
  }
}

async function listAvailableContent({
  token,
  tenantId,
  tenantLabel,
  contentType,
  windowStart,
  windowEnd,
  publisherId,
}: {
  token: string;
  tenantId: string;
  tenantLabel: string;
  contentType: string;
  windowStart: string;
  windowEnd: string;
  publisherId: string;
}) {
  const url =
    `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/content` +
    `?contentType=${encodeURIComponent(contentType)}` +
    `&startTime=${encodeURIComponent(windowStart)}` +
    `&endTime=${encodeURIComponent(windowEnd)}` +
    `&PublisherIdentifier=${encodeURIComponent(publisherId)}`;

  const res = await fetchWithRetry(
    url,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        "User-Agent": USER_AGENT,
      },
    },
    { label: `${tenantLabel}/${contentType}/content_list` },
  );

  if (!res.ok) {
    const text = await res.text().catch(() => "<no-body>");
    throw new Error(`[${tenantLabel}/${contentType}] List content failed: HTTP ${res.status} ${text}`);
  }

  const data = (await res.json()) as unknown;
  return Array.isArray(data) ? (data as ContentItem[]) : [];
}

async function fetchContentBlob({
  token,
  tenantLabel,
  contentType,
  contentUri,
  publisherId,
}: {
  token: string;
  tenantLabel: string;
  contentType: string;
  contentUri: string;
  publisherId: string;
}) {
  const blobUrl = new URL(contentUri);
  blobUrl.searchParams.set("PublisherIdentifier", publisherId);

  const res = await fetchWithRetry(
    blobUrl.toString(),
    {
      headers: {
        Authorization: `Bearer ${token}`,
        "User-Agent": USER_AGENT,
      },
    },
    { label: `${tenantLabel}/${contentType}/content_blob` },
  );

  if (!res.ok) {
    const text = await res.text().catch(() => "<no-body>");
    throw new Error(`[${tenantLabel}/${contentType}] Download blob failed: HTTP ${res.status} ${text}`);
  }

  const data = (await res.json()) as unknown;
  return Array.isArray(data) ? data : [data];
}

function buildContentKey(tenantLabel: string, contentType: string, contentId: string) {
  return `${tenantLabel}_unified_content_${contentType}:${contentId}`;
}

function buildBlobFilename({
  folder,
  contentId,
  contentCreated,
}: {
  folder: string;
  contentId: string;
  contentCreated?: string;
}) {
  const safeCreated = (contentCreated || new Date().toISOString()).replace(/[:.]/g, "-");
  const safeId = contentId.replace(/[^A-Za-z0-9._-]/g, "_");
  return `m365/${folder}/${safeCreated}_${safeId}.json`;
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
