# Cost Report Node API

This is the frontend-facing Node/TypeScript API and the only component that
connects to Supabase. Upload parsing, transformation, and Excel generation run
in-process in TypeScript.

## Setup

The Supabase database schema is already provisioned. No migration step is
required for this handoff.

1. Copy `.env.example` to `.env` and set the backend Supabase values.
2. Install, build, and start the service:

```powershell
npm ci
npm run build
npm start
```

Swagger is available at `http://127.0.0.1:3001/docs`.
See `API_AND_DATA_GUIDE.md` for the frontend API contract and stored data map.

## Runtime Ownership

- Node receives uploads and validates admin access.
- Node parses uploaded XLSX/CSV files and transforms them into canonical rows.
- Node upserts the cost and AI rows into Supabase and removes stale
  rows for the uploaded month/subscription scope.
- Frontend reads come from Node and the resolved Supabase views.
- For downloads, Node reads filtered rows and builds the workbook in memory.
- Uploads are temporary and generated Excel files are streamed without storage.

All Supabase client creation, database reads, writes, report persistence, and
settings mapping operations are contained in `src/database.ts`.

## Main Routes

- `POST /api/reports/ingest`
- `GET /api/reports`
- `GET /api/costs`
- `GET /api/ai/usage`
- `GET /api/overview/summary`
- `GET /api/reports/:reportMonth/excel`
- `GET|POST|PATCH /api/settings/project-applications`
- `GET|POST|PATCH|DELETE /api/settings/resource-mappings`
- `GET /api/settings/resource-catalog`
- `GET /api/settings/unmapped`

Cost and Excel endpoints accept exact-match filters such as `subscription`,
`resource_group`, `project`, `application`, `environment`, and
`resource_type`. Dates use a
monthly `report_month` value such as `2026-08-01` because the sample Azure
export has monthly rather than daily granularity.

All `/api` routes require a valid Supabase user access token in the
`Authorization: Bearer <token>` header. Mutation routes additionally require
the user to have `app_metadata.role = "admin"`. The service-role key remains
on the Node server and must never be placed in the frontend.

The frontend needs only the API base URL, the Supabase project URL, and a
publishable key for login. Supply the service-role key separately to whoever
runs this backend; never include a populated `.env` in a handoff archive.

## Tests

Run the local unit tests and TypeScript build with:

```powershell
npm test
npm run build
```

The live suite creates temporary Supabase users and settings records, exercises
every API route and method, verifies the Swagger path contract, tests the real
multipart upload and download flow, and removes its artifacts:

```powershell
npm run test:live -- "path\to\cost-report.xlsx" "path\to\ai-usage.csv"
```
