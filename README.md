# Policy Consistency Checker

> AI-powered SharePoint web part that automatically detects semantic contradictions between policy documents using Azure OpenAI, Cosmos DB vector search, and Microsoft 365 Copilot Chat API.

![SPFx](https://img.shields.io/badge/SPFx-1.22.2-green.svg)
![Azure Functions](https://img.shields.io/badge/Azure%20Functions-v4-blue.svg)
![Node.js](https://img.shields.io/badge/Node.js-20+-brightgreen.svg)
![TypeScript](https://img.shields.io/badge/TypeScript-5.x-blue.svg)
![Hackathon](https://img.shields.io/badge/Powell%20Hackathon-2026-purple.svg)

---

## What it does

When a document is uploaded to a monitored SharePoint library, the system automatically chunks and embeds it, performs vector similarity search against reference policy documents, sends semantically similar pairs to an LLM judge, and surfaces detected contradictions and ambiguities directly inside SharePoint — with Microsoft 365 Copilot integration for resolution assistance.

---

## Repository structure

This repository contains two projects:

```
/
├── policy-consistency-fn/     ← Azure Functions backend (AI pipeline + REST API)
└── policy-consistency-spfx/   ← SPFx web part (React frontend)
```

---

## Architecture

```
SharePoint Doc Library
        │ webhook
        ▼
fn-watcher (HTTP)  ──enqueue──►  sp-notifications queue
        │
        ▼
fn-processor (queue)
  reads IsReferenceDocument column via Graph delta API
        │
        ├── IsReference = true  ──►  ref-indexing queue  ──►  fn-indexer
        │                                                       (embed + store in Cosmos)
        │
        └── IsReference = false ──►  doc-analysis queue  ──►  fn-analyzer
                                                                (vector search + LLM judge → verdict)

Cosmos DB ◄─────────────────────────────────────────────────────────────────────┘
  chunks     (1536-dim embeddings, DiskANN)
  documents  (file metadata, change tokens)
  verdicts   (contradicts / ambiguous, resolution history)

Redis Cache  ←  verdict API cache (60s) + LLM result cache (7d)

HTTP API
  GET  /api/verdicts         fn-verdicts       (JWT-protected, paginated)
  PATCH /api/verdicts/{id}   fn-verdict-action (mark resolved)

Timer
  every 30d                  fn-subscription-renew (179-day webhook renewal)

SPFx Web Part (React · SharePoint Online)
  Dashboard · Detail Panel · PolicyQABar · DigestModal
        │
        ▼
M365 Copilot Chat API (Graph v1.0 · /copilot/conversations)
  - Ask about this conflict
  - Policy Q&A (enterprise search grounding)
  - Reconciliation suggestion
  - Weekly digest
```

---

## Tech stack

| Layer | Technology |
|---|---|
| Frontend | SPFx 1.22.2 · React 17 · Fluent UI v8 · CSS Modules |
| Auth (frontend) | `aadTokenProviderFactory` · `MSGraphClientV3` |
| Backend | Azure Functions v4 · Node.js 20+ · TypeScript 5.x |
| AI | Azure OpenAI — `text-embedding-3-small` · `gpt-4.1-mini` |
| Vector search | Cosmos DB NoSQL · DiskANN · cosine · 1536 dimensions |
| Data | Cosmos DB (3 containers) · Azure Cache for Redis |
| Messaging | Azure Storage Queues (3 queues) |
| Auth (backend) | Entra ID client credentials · JWKS JWT validation |
| Copilot | M365 Copilot Chat API · Graph v1.0 |

---

## Prerequisites

### Azure resources

| Resource | Notes |
|---|---|
| Azure Functions App | Node.js 20, v4 runtime |
| Azure Storage Account | Create 3 queues: `sp-notifications`, `ref-indexing`, `doc-analysis` |
| Azure Cosmos DB | Create database + 3 containers — see [Cosmos DB setup](#cosmos-db-setup) |
| Azure Cache for Redis | Any tier |
| Azure OpenAI | Deploy `text-embedding-3-small` and `gpt-4.1-mini` |
| Azure Entra ID App Registration | Client credentials · Graph `Sites.Read.All` · `AiEnterpriseInteraction.ReadWrite.All` |

### SharePoint

- A document library with an `IsReferenceDocument` Yes/No column
- A webhook subscription pointing to `fn-watcher`

> Both can be set up via `SharePointService.ensureIsReferenceDocumentColumn()` and `SharePointService.createListSubscription()`, or manually via Graph API.

### Local development

- Node.js `>=20.0.0` (backend) / `>=22.14.0 <23.0.0` (SPFx)
- Azure Functions Core Tools v4
- Azurite (local queue/blob emulation)
- Microsoft 365 tenant with Copilot license for Copilot features

---

## Getting started

### Backend — Azure Functions

```bash
cd policy-consistency-fn
npm install
cp local.settings.json.sample local.settings.json
# Fill in your values — see Environment variables section below
npm start   # builds TypeScript then runs func start on port 7072
```

### Frontend — SPFx

```bash
cd policy-consistency-spfx
npm install
npm run start   # local workbench in watch mode
npm run build   # production build + .sppkg packaging
```

> Set the `azureFunctionsBaseUrl` property in the SPFx Property Pane after deploying to SharePoint. The monitored site URL is auto-detected from `pageContext` and can be changed at runtime via the integrated SitePicker.

---

## Environment variables

### Azure Identity

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | Entra ID tenant ID |
| `AZURE_CLIENT_ID` | App registration client ID |
| `AZURE_CLIENT_SECRET` | App registration client secret |

### Azure OpenAI

| Variable | Description |
|---|---|
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI endpoint URL |
| `AZURE_OPENAI_API_KEY` | API key |
| `AZURE_OPENAI_EMBEDDING_DEPLOYMENT` | e.g. `text-embedding-3-small` |
| `AZURE_OPENAI_GPT_DEPLOYMENT` | e.g. `gpt-4.1-mini` |

### Cosmos DB

| Variable | Description |
|---|---|
| `COSMOS_ENDPOINT` | Cosmos DB account endpoint |
| `COSMOS_KEY` | Cosmos DB account key |
| `COSMOS_DATABASE` | Database name |
| `COSMOS_CONTAINER_CHUNKS` | `chunks` |
| `COSMOS_CONTAINER_VERDICTS` | `verdicts` |
| `COSMOS_CONTAINER_DOCUMENTS` | `documents` |

### Redis

| Variable | Description |
|---|---|
| `REDIS_CONNECTION_STRING` | `host:port,password=...,ssl=True` |

### Azure Storage

| Variable | Description |
|---|---|
| `STORAGE_CONNECTION_STRING` | Storage account connection string |
| `SP_NOTIFICATIONS_QUEUE` | `sp-notifications` |
| `REF_INDEXING_QUEUE` | `ref-indexing` |
| `DOC_ANALYSIS_QUEUE` | `doc-analysis` |

### SharePoint

| Variable | Description |
|---|---|
| `WEBHOOK_ENDPOINT` | Public HTTPS URL for `fn-watcher` |
| `WEBHOOK_CLIENT_STATE` | Secret token for webhook validation |
| `SP_REFERENCE_FIELD` | Column name (default: `IsReferenceDocument`) |

---

## Cosmos DB setup

Three containers, all partitioned by `/siteUrl` (= `/partitionKey`):

| Container | Key fields |
|---|---|
| `chunks` | `embedding[1536]` · DiskANN vector index · `contentHash` (SHA-256 dedup) |
| `verdicts` | `verdict` · `confidence` · `explanation` · `refExcerpt` · `newExcerpt` · `resolvedAt` |
| `documents` | `isReference` · `status` · `changeToken` |

> The DiskANN vector index on `chunks` is created automatically on first run via `CosmosService.initialize()`. **The vector policy on `chunks` is immutable** — create the container with the correct settings before running the app.

---

## API reference

### `GET /api/verdicts`

Returns paginated active verdicts. Requires Bearer token.

**Query parameters:** `siteUrl` (required) · `verdict` (contradicts|ambiguous) · `status` (active|resolved|all) · `from` · `to` · `newDocumentId` · `pageSize` · `continuationToken`

**Response:**
```json
{
  "verdicts": [{
    "id": "...",
    "verdict": "contradicts",
    "confidence": 0.92,
    "explanation": "...",
    "referenceFileName": "...",
    "newDocFileName": "...",
    "referenceExcerpt": "...",
    "newDocExcerpt": "...",
    "detectedAt": "2026-03-15T10:00:00Z",
    "isActive": true
  }],
  "continuationToken": "..."
}
```

### `PATCH /api/verdicts/{id}`

Marks a verdict as resolved. Requires Bearer token.

```json
{
  "action": "resolve",
  "siteUrl": "https://contoso.sharepoint.com/sites/policies",
  "resolvedBy": "user@contoso.com",
  "comment": "Updated section 4.2 to align with reference policy."
}
```

---

## Key design decisions

- **Similarity threshold** — Only Cosmos vector search hits with cosine similarity > 0.6 are sent to the LLM judge, reducing unnecessary API calls.
- **Verdict persistence** — Only `contradicts` and `ambiguous` verdicts are stored. `consistent` results are silently discarded.
- **Content deduplication** — SHA-256 hash check before re-indexing; unchanged documents are skipped entirely.
- **Chunking strategy** — 400 tokens/chunk with 60-token overlap (tiktoken). Splits on paragraph boundaries.
- **LLM result caching** — Judgment results cached in Redis for 7 days keyed on `(refChunkId, newChunkId)`.
- **Verdict API caching** — Query results cached in Redis for 60 seconds.
- **Webhook renewal** — SharePoint max subscription lifetime is 180 days. Timer runs every 30 days, extends by 179 days.
- **Multi-tenant isolation** — All Cosmos containers partition by `siteUrl`.

---

## Copilot features

All four features call the M365 Copilot Chat API directly from the SPFx web part via `MSGraphClientV3`, grounded on documents in the SharePoint tenant.

| Feature | Description |
|---|---|
| **Ask about this conflict** | Per-verdict analysis grounded on both documents — cites sources, recommends HR action |
| **Policy Q&A** | Free-form questions answered against the monitored site's policy documents |
| **Reconciliation suggestion** | Copilot drafts a reconciliation paragraph — editable before confirming resolution |
| **Weekly digest** | Executive summary of all active conflicts, ready to share by email |

> Requires `AiEnterpriseInteraction.ReadWrite.All` Graph permission with admin consent.

---

## Deployment

### Backend

```bash
cd policy-consistency-fn
npm run build
func azure functionapp publish <your-function-app-name>
```

### Frontend

```bash
cd policy-consistency-spfx
npm run build
# Upload the generated .sppkg from sharepoint/solution/ to your SharePoint App Catalog
# Approve the webApiPermissionRequests in the SharePoint Admin Center
```

**Required SPFx API permissions (package-solution.json):**

```json
"webApiPermissionRequests": [
  { "resource": "b950c535-bae8-408f-8ae1-81524b57d3ae", "scope": "Verdicts.Read" },
  { "resource": "Microsoft Graph", "scope": "AiEnterpriseInteraction.ReadWrite.All" }
]
```

---

## Built at Powell Hackathon 2026
