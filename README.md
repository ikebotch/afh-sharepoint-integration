# AFH Common SharePoint Utils

## Purpose
`AFH.Common.SharePointUtils` is the shared SharePoint integration SDK for the AFH platform.

It provides generic SharePoint capabilities for site/list discovery, list CRUD and querying, and document-library access. It is intentionally an SDK rather than a business-domain library; consuming services keep their own policies, orchestration, and persistence concerns.

## Repo Layout
- `src`
- `tests`

## Build And Test
- `dotnet build AFH.Common.SharePointUtils.sln --no-restore`
- `dotnet build src/AFH.Common.SharePointUtils.csproj --no-restore`
- `dotnet test tests/AFH.Common.SharePointUtils.Tests.csproj --no-restore`

## Verification Notes
- Unit-style SDK tests run locally.
- Tenant-backed integration tests are skipped unless the required SharePoint environment variables are available.
- For local integration setup, copy `tests/.env.example` to `tests/.env` or provide the same values through your shell environment.
