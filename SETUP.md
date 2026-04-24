# TrueWatch M365 Cloudflare Workers — Setup Guide

A set of Cloudflare Workers that collect Microsoft 365 logs and metrics and forward them to TrueWatch via DataKit.

## Architecture

```
[Entra Log Collector]         ---> R2 ---> Queue ---> [Entra Log Processor]         ---> DataKit ---> TrueWatch
[Entra App Inventory Collector] -> R2 ---> Queue ---> [Entra App Inventory Processor] -> DataKit ---> TrueWatch
[O365 Unified Log Collector]  ---> R2 ---> Queue ---> [O365 Unified Log Processor]   ---> DataKit ---> TrueWatch
[M365 License Collector]      ---> R2 ---> Queue ---> [M365 License Processor]        -> DataKit ---> TrueWatch
```

## Prerequisites

- Cloudflare account with **Workers Paid plan** ($5/month minimum)
- Microsoft 365 tenant with an Entra ID App Registration (client credentials)
- TrueWatch account with a DataWay URL and token
- Docker Desktop installed and running (for building the DataKit container image)

## Step 1 — Install Dependencies and Log In

```bash
npm install
wrangler login
```

## Step 2 — Create R2 Bucket

```bash
wrangler r2 bucket create m365-logs
```

## Step 3 — Create KV Namespaces

Run each command and copy the returned ID into the corresponding `wrangler.jsonc`:

```bash
wrangler kv namespace create LOG_CURSORS
# → update entra-log-collector/wrangler.jsonc

wrangler kv namespace create LOG_IDS_ENTRA
# → update entra-log-processor/wrangler.jsonc

wrangler kv namespace create LOG_CONTENT_IDS
# → update o365-unified-log-collector/wrangler.jsonc

wrangler kv namespace create LOG_IDS_O365
# → update o365-unified-log-processor/wrangler.jsonc
```

## Step 4 — Create Queues

```bash
wrangler queues create entra-log-events
wrangler queues create entra-application-events
wrangler queues create o365-unified-log-events
wrangler queues create m365-license-events
```

## Step 5 — Set R2 Event Notification Rules (Cloudflare Dashboard)

Navigate to **R2 → m365-logs bucket → Settings → Event Notifications** and add:

| Prefix                   | Queue                       |
|--------------------------|-----------------------------|
| `m365/entra/`            | `entra-log-events`          |
| `m365/ms_applications/`  | `entra-application-events`  |
| `m365/exchange/`         | `o365-unified-log-events`   |
| `m365/sharepoint/`       | `o365-unified-log-events`   |
| `m365/general/`          | `o365-unified-log-events`   |
| `m365/azureactivedirectory/` | `o365-unified-log-events` |
| `m365/dlp/`              | `o365-unified-log-events`   |
| `m365/licenses/`         | `m365-license-events`       |

## Step 6 — Update wrangler.jsonc Placeholders

In each collector's `wrangler.jsonc`, replace:
- `YOUR_TENANT_LABEL` — a short label for your tenant (e.g. `CONTOSO`)
- `YOUR_TENANT_ID` — your Entra ID tenant ID (GUID)
- `YOUR_CLIENT_ID` — your App Registration client ID (GUID)
- `REPLACE_WITH_KV_NAMESPACE_ID` — the KV namespace ID from Step 3

In `o365-unified-log-processor/src/index.ts`, update `ROUTE_CONFIG` to match your `TENANT_LABEL`.

## Step 7 — Deploy DataKit

**7a — Update the DataKit version (optional)**

Open `datakit/Dockerfile` and update the version tag if needed. Get the latest version from your TrueWatch Platform workspace settings.

**7b — Update the DataWay URL**

1. Go to **TrueWatch Console → Integrations → DataKit**
2. Copy the DataWay address (format: `https://openway.truewatch.com?token=YOUR_TOKEN`)
3. Open `datakit/src/index.ts` and replace `YOUR_TRUEWATCH_DATAWAY_URL` with the copied address

**7c — Deploy**

```bash
cd datakit
wrangler deploy
cd ..
```

Wrangler automatically builds the local `Dockerfile` and pushes it to Cloudflare's container registry before deploying.

## Step 8 — Set Secrets

Run for each collector:

```bash
wrangler secret put CLIENT_SECRET --name entra-log-collector
wrangler secret put CLIENT_SECRET --name entra-app-inventory-collector
wrangler secret put CLIENT_SECRET --name o365-unified-log-collector
wrangler secret put CLIENT_SECRET --name m365-license-collector
```

## Step 9 — Deploy All Workers

```bash
cd entra-log-collector && wrangler deploy && cd ..
cd entra-log-processor && wrangler deploy && cd ..
cd entra-app-inventory-collector && wrangler deploy && cd ..
cd entra-app-inventory-processor && wrangler deploy && cd ..
cd o365-unified-log-collector && wrangler deploy && cd ..
cd o365-unified-log-processor && wrangler deploy && cd ..
cd m365-license-collector && wrangler deploy && cd ..
cd m365-license-processor && wrangler deploy && cd ..
```

## Microsoft Entra App Registration Permissions

### Creating the App Registration

1. Go to **Microsoft Entra admin center → App registrations → New registration**
2. Give it a name (e.g. `TrueWatch M365 Collector`) and register it
3. Note the **Tenant ID** and **Client ID** from the Overview page
4. Under **Certificates & secrets → Client secrets**, create a new secret and note the value
5. Under **API permissions**, add each permission below — select **Application permissions** (not Delegated)
6. Click **Grant admin consent** for all permissions

> ⚠️ All permissions must be **Application** type, not Delegated. Admin consent is required for all.

---

### Permissions by Worker

**entra-log-collector** — Microsoft Graph

| Permission | Description |
|---|---|
| `AuditLog.Read.All` | Read all audit log data |
| `Directory.Read.All` | Read directory data |

**entra-app-inventory-collector** — Microsoft Graph

| Permission | Description |
|---|---|
| `Application.Read.All` | Read all applications |

**o365-unified-log-collector** — Office 365 Management APIs

| Permission | Description |
|---|---|
| `ActivityFeed.Read` | Read activity data for your organization |
| `ActivityFeed.ReadDlp` | Read DLP policy events including detected sensitive data |

**m365-license-collector** — Microsoft Graph

| Permission | Description |
|---|---|
| `LicenseAssignment.Read.All` | Read license assignment information |

> **Note:** If `Directory.Read.All` is already granted (e.g. for entra-log-collector), `LicenseAssignment.Read.All` is not needed — `Directory.Read.All` covers this API as well.

## Cost Estimate (Cloudflare Workers Paid Plan)

| Component | Cost |
|-----------|------|
| Workers, R2, KV, Queues | Covered by $5/month tier |
| DataKit container (standard-2) | ~$40–$50 USD/month additional |
| **Total estimate** | **~$45–$55 USD/month** |

See [Cloudflare Containers Pricing](https://developers.cloudflare.com/containers/pricing/) for details.
