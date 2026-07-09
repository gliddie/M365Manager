/**********************************************************************************************
    M365Manager - Database & Logging Schema
    ------------------------------------------------------------------------------------------
    Run this script in SQL Server Management Studio (SSMS) against your SQL Server.
    It creates the [M365Manager] database and the logging table that replaces the old
    text-file logs (LogFile + ReportFile) of the legacy PowerShell GUI.

    Safe to re-run: every step is guarded with IF NOT EXISTS.

    Notes:
      * Application settings and secrets (SQL password, M365 app details) are NOT stored here.
        They are kept locally per admin machine, encrypted with Windows DPAPI.
      * All timestamps are stored in UTC (DATETIME2). Convert to local time for display.
**********************************************************************************************/

/*------------------------------------------------------------------------------------------
  1. Create the database (only if it does not already exist)
------------------------------------------------------------------------------------------*/
IF DB_ID(N'M365Manager') IS NULL
BEGIN
    CREATE DATABASE [M365Manager];
END
GO

USE [M365Manager];
GO

/*------------------------------------------------------------------------------------------
  2. Logging table

     Maps the legacy tab-separated log lines to structured columns:
       Legacy audit line : Date  Time  MachineName  EventCode  Message
       Legacy report line: Date  Time  Message
     ...plus new fields useful for M365 admin auditing.
------------------------------------------------------------------------------------------*/
IF OBJECT_ID(N'dbo.LogEntries', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.LogEntries
    (
        Id             BIGINT           IDENTITY(1,1) NOT NULL,

        -- When the event happened (UTC). SYSUTCDATETIME() default for server-side inserts.
        TimestampUtc   DATETIME2(3)     NOT NULL CONSTRAINT DF_LogEntries_TimestampUtc DEFAULT (SYSUTCDATETIME()),

        -- Who / where
        UserUpn        NVARCHAR(256)    NULL,   -- signed-in M365 admin (the colleague performing the change)
        WindowsUser    NVARCHAR(128)    NULL,   -- Windows account running the app (whoami)
        MachineName    NVARCHAR(128)    NULL,   -- computer the action was launched from

        -- What
        Area           NVARCHAR(64)     NULL,   -- functional area, e.g. 'UserMailbox', 'Licensing', 'Groups'
        [Action]       NVARCHAR(128)    NULL,   -- operation / former script name (e.g. 'AddFolderPermissions')
        TargetObject   NVARCHAR(256)    NULL,   -- object acted on (mailbox/UPN/group/domain)

        -- Result classification
        EventCode      NVARCHAR(16)     NULL,   -- legacy 4-char codes: STAR, INFO, ADD, REMO, ERR, FAIL, PASS, WARN, UPDA, CREATE, DELE...
        Severity       TINYINT          NOT NULL CONSTRAINT DF_LogEntries_Severity DEFAULT (1),
                                                -- 0=Debug, 1=Info, 2=Success, 3=Warning, 4=Error

        -- Details
        [Message]      NVARCHAR(MAX)    NULL,

        -- Correlate all entries belonging to one operation/session
        CorrelationId  UNIQUEIDENTIFIER NULL,

        -- Row insertion time (server clock, UTC)
        CreatedAtUtc   DATETIME2(3)     NOT NULL CONSTRAINT DF_LogEntries_CreatedAtUtc DEFAULT (SYSUTCDATETIME()),

        CONSTRAINT PK_LogEntries PRIMARY KEY CLUSTERED (Id)
    );
END
GO

/*------------------------------------------------------------------------------------------
  3. Indexes to keep the log view / filters fast
------------------------------------------------------------------------------------------*/
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LogEntries_TimestampUtc' AND object_id = OBJECT_ID(N'dbo.LogEntries'))
    CREATE INDEX IX_LogEntries_TimestampUtc ON dbo.LogEntries (TimestampUtc DESC);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LogEntries_UserUpn' AND object_id = OBJECT_ID(N'dbo.LogEntries'))
    CREATE INDEX IX_LogEntries_UserUpn ON dbo.LogEntries (UserUpn) INCLUDE (TimestampUtc, Severity);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LogEntries_TargetObject' AND object_id = OBJECT_ID(N'dbo.LogEntries'))
    CREATE INDEX IX_LogEntries_TargetObject ON dbo.LogEntries (TargetObject) INCLUDE (TimestampUtc, [Action]);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LogEntries_Severity' AND object_id = OBJECT_ID(N'dbo.LogEntries'))
    CREATE INDEX IX_LogEntries_Severity ON dbo.LogEntries (Severity, TimestampUtc DESC);
GO

/*------------------------------------------------------------------------------------------
  4. Quick connectivity self-test (optional) - inserts and removes one row.
------------------------------------------------------------------------------------------*/
-- INSERT INTO dbo.LogEntries (UserUpn, WindowsUser, MachineName, Area, [Action], EventCode, Severity, [Message])
-- VALUES (N'test@ul.com', SUSER_SNAME(), HOST_NAME(), N'System', N'ConnectionTest', N'STAR', 1, N'M365Manager connectivity test');
-- SELECT TOP 5 * FROM dbo.LogEntries ORDER BY Id DESC;
GO

PRINT 'M365Manager database and logging schema are ready.';
GO
