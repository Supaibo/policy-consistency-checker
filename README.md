# Policy Consistency Checker

> AI-powered SharePoint web part that automatically detects semantic contradictions between policy documents using Azure OpenAI, Cosmos DB vector search, and Microsoft 365 Copilot Chat API.

![SPFx](https://img.shields.io/badge/SPFx-1.22.2-green.svg)
![Azure Functions](https://img.shields.io/badge/Azure%20Functions-v4-blue.svg)
![Node.js](https://img.shields.io/badge/Node.js-20+-brightgreen.svg)
![TypeScript](https://img.shields.io/badge/TypeScript-5.x-blue.svg)
![Hackathon](https://img.shields.io/badge/Microsoft%20SharePoint%20Hackathon-2026-0078D4)

---

## The problem

Large organizations store hundreds of policy documents, procedures, and team guides in SharePoint — all living side by side, with no automated way to detect when they contradict each other. Compliance teams rely on manual audits. HR finds out too late.

**Policy Consistency Checker solves this.**

---

## What it does

When a document is uploaded to a monitored SharePoint library, the system automatically:

1. Chunks and embeds the document using `text-embedding-3-small`
2. Performs vector similarity search against indexed reference policy documents
3. Sends semantically similar chunk pairs to a `gpt-4.1-mini` LLM judge
4. Surfaces detected contradictions and ambiguities directly inside SharePoint
5. Provides Microsoft 365 Copilot-assisted resolution workflows

All of this runs within the Microsoft 365 trust boundary — no data egress, permissions preserved end to end.

---

## Demo

> 📹 [Watch the demo video](#) — 8 minutes

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

Cosmos DB
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
  - Ask about this conflict   (grounded on tenant docs)
  - Policy Q&A                (enterprise search grounding)
  - Reconciliation suggestion (editable, confirm & resolve)
  - Weekly digest             (email-ready executive summary)
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

## Features

- **Automatic detection** — documents are analyzed the moment they are uploaded, no manual trigger required
- **Dashboard** — contradiction and ambiguity verdicts with confidence scores, filterable by type and document
- **Side-by-side comparison** — highlighted excerpts from both documents in the detail panel
- **Verdict resolution** — acknowledge, override, or ignore with a comment and full audit trail
- **Copilot Ask** — per-conflict analysis grounded on the SharePoint tenant, with source citations
- **Policy Q&A** — free-form questions answered against the monitored site's policy documents
- **Reconciliation suggestion** — Copilot drafts a reconciliation paragraph, editable before confirming
- **Weekly digest** — executive summary of all active conflicts, shareable by email via Graph API
- **Site picker** — switch between SharePoint sites at runtime without reconfiguring the web part
- **Pagination** — continuation token-based pagination for large verdict sets

---

## Azure Functions — 7 functions

| Function | Trigger | Role |
|---|---|---|
| `fn-watcher` | HTTP POST | Receives SharePoint webhooks, validates handshake, enqueues notifications |
| `fn-processor` | Queue `sp-notifications` | Calls Graph delta API, routes to `ref-indexing` or `doc-analysis` |
| `fn-indexer` | Queue `ref-indexing` | Download → extract → chunk → embed → store reference docs in Cosmos |
| `fn-analyzer` | Queue `doc-analysis` | Download → chunk → vector search → LLM judge → store verdicts |
| `fn-verdicts` | HTTP GET | Returns paginated, JWT-protected verdicts with Redis cache (60s TTL) |
| `fn-verdict-action` | HTTP PATCH | Marks verdicts as resolved with metadata |
| `fn-subscription-renew` | Timer (30d) | Renews expiring SharePoint webhook subscriptions |

---

## Copilot integration

All four features call the M365 Copilot Chat API directly from the SPFx web part via `MSGraphClientV3`, grounded on documents in the SharePoint tenant. No custom RAG pipeline required.

| Feature | Copilot capability used |
|---|---|
| **Ask about this conflict** | Grounded analysis on both documents · cites sources · recommends HR action |
| **Policy Q&A** | Enterprise search grounding on the monitored SharePoint site |
| **Reconciliation suggestion** | Text generation grounded on the reference document |
| **Weekly digest** | Structured report generation from active verdict data |

> Requires `AiEnterpriseInteraction.ReadWrite.All` Graph permission with admin consent.

---

## Key design decisions

- **Similarity threshold** — only Cosmos vector search hits with cosine similarity > 0.6 are sent to the LLM judge, reducing unnecessary API calls
- **Verdict persistence** — only `contradicts` and `ambiguous` verdicts are stored; `consistent` results are silently discarded
- **Content deduplication** — SHA-256 hash check before re-indexing; unchanged documents are skipped entirely
- **Chunking strategy** — 400 tokens/chunk with 60-token overlap (tiktoken), splits on paragraph boundaries
- **LLM result caching** — judgment results cached in Redis for 7 days keyed on `(refChunkId, newChunkId)`
- **Verdict API caching** — query results cached in Redis for 60 seconds
- **Webhook renewal** — SharePoint max subscription lifetime is 180 days; timer runs every 30 days and extends by 179 days
- **Multi-tenant isolation** — all Cosmos containers partition by `siteUrl`

---

## Built for the Microsoft SharePoint Hackathon 2026

**Powell Software** — [powell-software.com](https://powell-software.com)
