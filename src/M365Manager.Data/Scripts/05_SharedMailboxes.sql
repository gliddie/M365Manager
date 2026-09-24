/**********************************************************************************************
    M365Manager - Shared Mailboxes Overview Cache
    ------------------------------------------------------------------------------------------
    Shared mailboxes are Exchange mailboxes, not Microsoft 365 groups, so - unlike dbo.Teams -
    this table is NOT foreign-keyed to dbo.Groups; it is keyed by the mailbox's own ExchangeGuid.
    The .ED/.AU/.RE access-group addresses are recorded for reference, but their membership and
    ManagedBy detail lives in Exchange, not here (see SharedMailboxService/ISharedMailboxService).

    Populated by src/ImportScripts/Import-SharedMailboxesToSql.ps1 (scheduled) and refreshed
    immediately by SharedMailboxService.CreateSharedMailboxAsync / ChangeOwnerAsync for mailboxes
    created or changed through this app.

    Safe to re-run: every step is guarded with IF NOT EXISTS.
**********************************************************************************************/

USE [M365Manager];
GO

IF OBJECT_ID(N'dbo.SharedMailboxes', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SharedMailboxes
    (
        ExchangeGuid                        UNIQUEIDENTIFIER NOT NULL,
        DisplayName                         NVARCHAR(256)    NOT NULL,
        Alias                                NVARCHAR(128)    NOT NULL,
        PrimarySmtpAddress                   NVARCHAR(256)    NOT NULL,
        EDGroupAddress                       NVARCHAR(256)    NULL,   -- Editor: FullAccess + SendAs
        AUGroupAddress                       NVARCHAR(256)    NULL,   -- Author: SendOnBehalf only
        REGroupAddress                       NVARCHAR(256)    NULL,   -- Reader: folder-level Reviewer only
        Owners                               NVARCHAR(MAX)    NULL,   -- .ED group's ManagedBy, semicolon-joined UPNs
        RequireSenderAuthenticationEnabled   BIT              NULL,
        CreatedDateTime                      DATETIME2(3)     NULL,
        IsDeletedInM365                      BIT              NOT NULL CONSTRAINT DF_SharedMailboxes_IsDeleted DEFAULT (0),
        LastImportedAtUtc                    DATETIME2(3)     NOT NULL CONSTRAINT DF_SharedMailboxes_LastImportedAtUtc DEFAULT (SYSUTCDATETIME()),
        LastSeenAtUtc                        DATETIME2(3)     NULL,

        CONSTRAINT PK_SharedMailboxes PRIMARY KEY CLUSTERED (ExchangeGuid)
    );

    CREATE UNIQUE INDEX IX_SharedMailboxes_PrimarySmtpAddress ON dbo.SharedMailboxes(PrimarySmtpAddress);
END
GO

PRINT 'M365Manager Shared Mailboxes overview schema is ready.';
GO
