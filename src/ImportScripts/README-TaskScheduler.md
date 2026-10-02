# Task Scheduler example for the M365 group import

Use this example to create a scheduled task on the server that runs the import script automatically.

## Recommended schedule
- Run every 4 hours
- Keep the log retention at 30 days and max 14 files

## Example task configuration

### Action
- Program/script: `"C:\Program Files\PowerShell\7\pwsh.exe"` (quoted - unquoted, the task fails with
  2147942402 because it looks for `C:\Program`)
- Add arguments (do NOT repeat `pwsh.exe` here - pwsh then exits with 64):
  `-NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\M365Manager\src\ImportScripts\Import-M365GroupsToSql.ps1" -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -CertificateThumbprint "YOUR_CERT_THUMBPRINT" -SqlServer "YOUR_SQL_SERVER" -SqlDatabase "M365Manager" -LogFile "C:\Temp\m365-group-import.log" -MaxLogFiles 14 -LogRetentionDays 30`
- Start in: the script folder

## SQL credential file

The scripts never take a SQL password on the command line. They read user and password from a
credential file, by default `D:\SCRIPTS\Secure\m365manager-sql.xml` (override with
`-SqlCredentialPath`). Export-Clixml encrypts the password with DPAPI, so the file can only be
read by the Windows account that created it, on the machine it was created on. Create it **as the
task's account, on the task's server**:

```powershell
Get-Credential -UserName M365Manager -Message 'SQL login' | Export-Clixml D:\SCRIPTS\Secure\m365manager-sql.xml
```

After a password change, recreate the file the same way - the tasks themselves stay untouched.

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
- If you use integrated security instead of the SQL login, add `-UseIntegratedSecurity`; the
  credential file is then not read.
