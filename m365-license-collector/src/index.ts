const FETCH_TIMEOUT_MS = 25_000;
const FETCH_RETRIES = 3;
const USER_AGENT = "M365LicenseCollector/1.0";

interface Env {
  CLIENT_ID: string;
  CLIENT_SECRET: string;
  LOGS_BUCKET: R2Bucket;
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
    console.log("Starting M365 license collection...");
    console.log(`cron=${event.cron} scheduledTime=${new Date(event.scheduledTime).toISOString()}`);

    const timestamp = new Date().toISOString();

    try {
      await collectLicenses({
        tenantLabel: env.TENANT_LABEL,
        tenantId: env.TENANT_ID,
        clientId: env.CLIENT_ID,
        clientSecret: env.CLIENT_SECRET,
        bucket: env.LOGS_BUCKET,
        timestamp,
      });
      console.log("M365 license collection completed successfully");
    } catch (error) {
      console.error("M365 license collection failed:", error);
      throw error;
    }
  },
};

async function collectLicenses({
  tenantLabel,
  tenantId,
  clientId,
  clientSecret,
  bucket,
  timestamp,
}: {
  tenantLabel: string;
  tenantId: string;
  clientId: string;
  clientSecret: string;
  bucket: R2Bucket;
  timestamp: string;
}) {
  console.log(`=== Collecting ms_licenses for ${tenantLabel} (${tenantId}) ===`);

  const token = await getAccessToken({ tenantId, clientId, clientSecret, tenantLabel });

  const { skus } = await fetchSubscribedSkus({ token, tenantLabel });

  if (skus.length) {
    await storeLogs(bucket, timestamp, skus);
  } else {
    console.log(`[${tenantLabel}/ms_licenses] No records to store.`);
  }

  console.log(`=== Completed ${tenantLabel} ===`);
}

async function fetchSubscribedSkus({
  token,
  tenantLabel,
}: {
  token: string;
  tenantLabel: string;
}) {
  // subscribedSkus is typically <100 records, but handle nextLink just in case
  let nextLink: string | null = "https://graph.microsoft.com/v1.0/subscribedSkus";
  const collected: unknown[] = [];
  let page = 0;

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
      { label: `${tenantLabel}/ms_licenses/page_${page}` },
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

  console.log(`[${tenantLabel}/ms_licenses] Fetched ${collected.length} SKUs from Graph.`);
  return { skus: collected };
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
  if (!json?.access_token) {
    throw new Error(`Token response missing access_token for tenant ${tenantId}`);
  }

  return json.access_token;
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

async function storeLogs(bucket: R2Bucket, timestamp: string, skus: unknown[]) {
  const safeTs = timestamp.replace(/[:.]/g, "-");
  const filename = `m365/licenses/${safeTs}.json`;

  await bucket.put(filename, JSON.stringify(skus), {
    httpMetadata: { contentType: "application/json" },
  });

  console.log(`Stored ${skus.length} SKUs -> ${filename}`);
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
