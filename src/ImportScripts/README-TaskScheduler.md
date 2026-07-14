# Task Scheduler example for the M365 group import

Use this example to create a scheduled task on the server that runs the import script automatically.

## Recommended schedule
- Run every 4 hours
- Keep the log retention at 30 days and max 14 files

## Example task configuration

### Action
- Program/script: `pwsh.exe`
- Add arguments:
  `-NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\M365Manager\src\ImportScripts\Import-M365GroupsToSql.ps1" -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -CertificateThumbprint "YOUR_CERT_THUMBPRINT" -SqlServer "YOUR_SQL_SERVER" -SqlDatabase "M365Manager" -SqlUser "YOUR_SQL_USER" -SqlPassword "YOUR_SQL_PASSWORD" -LogFile "C:\Temp\m365-group-import.log" -MaxLogFiles 14 -LogRetentionDays 30`

### Trigger
- Begin the task: On a schedule
- Settings: Daily
- Recur every: 4 hours
- Start: choose a sensible start time such as 00:00

### Conditions
- Leave default settings unless you want to restrict it to a specific network profile

### Settings
- Run whether user is logged on or not
- Run with highest privileges: optional
- If the task is already running, do not start a new instance

## Notes
- The task should run under a service account or dedicated admin account that can access the SQL Server and has the certificate available.
- The certificate must be installed for the account running the task.
- If you use integrated security instead of SQL credentials, remove `-SqlUser` and `-SqlPassword` and add `-UseIntegratedSecurity`.
