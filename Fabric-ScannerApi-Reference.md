# Fabric Scanner — Reference

## Overview
This PowerShell script scans Power BI / Fabric workspaces using the Admin Scanner APIs and exports enriched metadata to JSON and CSV files. It collects workspace, report, dataset, table, column, measure, datasource and lineage metadata, and optionally attempts to retrieve dataset refresh history (last refresh times and recent refresh attempts).

Primary goals:
- Produce a single merged JSON scan result and flat CSV exports for easy analysis.
- Improve column and datatype extraction with robust null checks and error handling.
- Optionally collect dataset refresh history and last refresh timestamps when supported.

## What the script does (high level)
1. Connects interactively to Power BI using `Connect-PowerBIServiceAccount`.
2. Verifies Admin API access (`/admin/groups`) and enumerates all workspaces (paged).
3. Optionally allows interactive workspace selection or wildcard filtering.
4. Submits workspace batches to the Admin `workspaces/getInfo` scanner API (up to 100 ids per request).
5. Waits for scanner job(s) to complete, fetches the scan result JSON, and validates basic quality checks.
6. Merges batch JSON payloads into one final JSON file and flattens that into multiple CSVs (workspaces, reports, datasets, tables, columns, measures, datasources, lineage, refresh history, etc.).

## Requirements
- PowerShell (Windows PowerShell or PowerShell Core / pwsh). The script runs on Windows but is compatible with pwsh.
- Modules (`Install-Module -Name MicrosoftPowerBIMgmt`):
  - `MicrosoftPowerBIMgmt.Profile` (for `Connect-PowerBIServiceAccount`)
  - `MicrosoftPowerBIMgmt.Admin` (Admin REST wrappers: listing groups, invoking admin endpoints)

- Permissions:
  - The account used to run the script must have Power BI / Fabric admin privileges to call Admin APIs (`/admin/groups`, `/admin/workspaces/*`). If you don't have admin privileges the script will fail when trying to list workspaces as an admin.
  - To retrieve dataset schema or expressions, tenant metadata-scanning features needs to be enabled in the Power BI Admin portal: Admin API Settings -> "Enhance admin APIs responses with detailed metadata"

Notes on service principals: the script is written for interactive sign-in (Entra ID user account). You can adapt it to use a service principal / app credentials, but that requires changes to the authentication flow and appropriate permissions.

## Inputs / Parameters
The script exposes several top-level script parameters (examples shown):

- `-ExportRoot <path>` (default: `$HOME\PowerBI_Metadata`): root folder where per-run timestamped folders are created.
- `-SelectWorkspaces` (switch): show an interactive selector (Out-GridView or text-based) to pick target workspaces.
- `-WorkspaceNameLike <string>`: wildcard filter (PowerShell `-like`) to narrow workspace choices by name.
- `-IncludeLineage` (switch, default true): include lineage edges in exports.
- `-IncludeDatasourceDetails` (switch, default true): include datasource connection details.
- `-IncludeDatasetSchema` (switch, default true): gather dataset tables/columns/measures.
- `-IncludeDatasetExpressions` (switch, default true): include DAX/M expressions when metadata-scanning is enabled in the tenant.
- `-IncludeRefreshHistory` (switch, default true): attempt to query dataset refresh history (note: not all dataset types support refresh history).
- `-MaxRetries <int>` (default 3): global retry attempts for transient API failures.
- `-RetryDelaySeconds <int>` (default 10): delay between retries (seconds).

## Output files (per-run)
When the scan completes the script creates a timestamped run folder under `ExportRoot` and writes the following files (examples):

- `scanner_result_<TIMESTAMP>.json` — merged scanner JSON (full result)
- `workspaces_<TIMESTAMP>.csv` — workspace list and properties
- `reports_<TIMESTAMP>.csv` — report id/name/dataset mapping
- `datasets_<TIMESTAMP>.csv` — datasets with basic properties and last refresh columns
- `dataset_tables_<TIMESTAMP>.csv` — tables per dataset
- `dataset_columns_<TIMESTAMP>.csv` — columns with data types and attributes
- `dataset_measures_<TIMESTAMP>.csv` — measures and expressions
- `dataset_datasources_<TIMESTAMP>.csv` — datasource connection detail (JSON in a column)
- `lineage_edges_<TIMESTAMP>.csv` — extracted upstream lineage edges
- `dataset_refresh_history_<TIMESTAMP>.csv` — detailed refresh records (last N refresh attempts) when available
- `scan_statistics_<TIMESTAMP>.json` — simple scan run metrics and any recorded processing errors

Note: CSV filenames are suffixed with the timestamp for the run, and CSVs are UTF8 encoded with no type information headers.

## Refresh history behavior
- The script calls the Power BI refreshes API for each dataset (`/groups/{groupId}/datasets/{datasetId}/refreshes`).
- Not all dataset types support the refreshes API. Fabric lakehouses/warehouses, some dataflows, and other non-semantic-model content may return 404 for this endpoint. The script detects 404s and records `NotSupported` or `NoHistory` in the dataset row rather than failing the entire run.
- When refresh history is available the script records fields such as start/end times, status, refresh type, request id and optional service exception JSON.

## How to run
From a PowerShell prompt (pwsh), cd into the script folder and run:

- Interactive default run:
  .\Fabric-ScannerApi-Interactive.ps1

- Run for a subset of workspaces (by name wildcard):
  .\Fabric-ScannerApi-Interactive.ps1 -WorkspaceNameLike "*Finance*"

- Use interactive selection (Out-GridView or console selection):
  .\Fabric-ScannerApi-Interactive.ps1 -SelectWorkspaces

- Disable refresh history collection:
  .\Fabric-ScannerApi-Interactive.ps1 -IncludeRefreshHistory:$false

- Increase retry attempts and delay:
  .\Fabric-ScannerApi-Interactive.ps1 -MaxRetries 6 -RetryDelaySeconds 20

Note: `Connect-PowerBIServiceAccount` will prompt for interactive sign-in. Run the script in an environment where an interactive browser-based login is acceptable.

## Best practices and caveats
- Admin privileges are required for Admin APIs. If you are not a tenant admin the script will fail at the Admin calls.
- Requesting dataset expressions or schema requires tenant metadata scanning to be enabled; otherwise those parts of the scan will be missing.
- Retry behavior: the script provides `-MaxRetries` and `-RetryDelaySeconds` to handle transient errors. Avoid enabling high retry counts for POST operations that start scanner jobs (they may create duplicate jobs on retries). If you need safe retries for non-idempotent POSTs, add idempotency checks or add a more granular retry count for those calls.
- Refresh history: treat missing refresh history as typical for some content types; the script logs and continues when refresh APIs return 404.

## Troubleshooting
- Missing modules error: install the required Power BI modules via `Install-Module MicrosoftPowerBIMgmt.Profile` and `Install-Module MicrosoftPowerBIMgmt.Admin` as admin.
- Admin API access error: ensure your account is assigned a Power BI/Fabric admin role.
- No columns in CSV: you may lack dataset schema permissions, or the workspaces may contain content types that don't expose schema.
- Persistent HTTP 404/403/401 errors: check the workspace/dataset access and API changes in the Power BI Admin Portal and ensure tenant settings allow metadata scanning.

## Extending the script
- To run non-interactively with a service principal, replace the interactive auth portion with an OAuth flow for app credentials and give the app the necessary privileges.
- Consider adding exponential backoff + jitter for retries and special handling of `429 Too Many Requests` responses.
- Optionally whitelist or blacklist specific workspace IDs to control scope.

---

