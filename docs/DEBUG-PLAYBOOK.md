# Debug Playbook

This playbook helps identify root causes for people connector issues. It focuses on profile source registration/prioritization, ingestion/export status, and auth/permissions.

## Quick checklist

Run the checklist command:

- npm run debug:checklist -- --tenant-id <TENANT_ID> --connection-id <CONNECTION_ID>

If the command is not available, see the Manual checks section below.

## Symptoms -> checks -> fixes

### 1) Profile source registration fails

Symptoms:
- PeopleAdmin BadRequest: "Profile source is in priority list"
- Repeated "ProfileSourceRegistrar: PostProfileSource failed"

Checks:
- node tools/debug/check-profile-source.mjs
- node tools/debug/check-people-connector-status.mjs

Likely fixes:
- Remove the profile source from the priority list, then delete/recreate it.
- Ensure the connection ID is correct and matches the profile source ID.

### 2) Ingestion/export is not progressing

Symptoms:
- Export batch sizes stay small or never complete
- Items appear in connector but not in People profiles

Checks:
- node tools/debug/verify-ingestion-progress.mjs
- node tools/debug/verify-items.mjs

Likely fixes:
- Confirm schema alignment (personAccount label, string/stringCollection types).
- Verify external items are present and mapped to existing users.

### 3) Auth/permissions failures

Symptoms:
- TSS Unauthorized in telemetry
- Token scope errors or Graph 401s

Checks:
- node tools/debug/check-app-permissions.mjs
- node tools/debug/check-token-scopes.mjs

Likely fixes:
- Validate app-only permissions for Graph connector and People Admin.
- Re-consent admin scopes and ensure correct tenant/application IDs.

## Telemetry CSV exports (Kusto)

Use this query to export connector pipeline telemetry (CSV matches debug/internal_debug.csv):

```kusto
Jarvis ## let _connectionid = '';
let _endTime = datetime(2026-02-02T13:16:02Z);
let _environment = dynamic(null);
let _startTime = datetime(2026-02-01T08:16:02Z);
let _tid = '18c5fde2-f22d-4171-b868-c0f1d347bd7b';
ConnectorChangeEvent_Global
| where component == "ProfileConnectorProcessors"
// | where subScenario == "TransformAsync"
| where env_time between (_startTime.._endTime)
| where env_cloud_environment in (_environment) or isempty(_environment)
| where isempty(_connectionid) or connectionId == _connectionid
| where isempty(_tid) or tenantId == _tid
| summarize Count = count() by tenantId, connectionId, message, env_time
| order by Count desc
```

## Manual checks

Use these when you want to run each check individually (connection ID required):

- node tools/debug/check-profile-source.mjs <CONNECTION_ID> [CSV_PATH]
- node tools/debug/check-people-connector-status.mjs <CONNECTION_ID>
- node tools/debug/verify-ingestion-progress.mjs
- node tools/debug/check-app-permissions.mjs <CONNECTION_ID>
- node tools/debug/check-token-scopes.mjs

## Notes

- Do not enable verbose/PII logging in production.
- Prefer app-only permissions for connector ingestion and People Admin calls.
