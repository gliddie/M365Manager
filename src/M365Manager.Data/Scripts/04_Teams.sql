/**********************************************************************************************
    M365Manager - Teams Overview Cache
    ------------------------------------------------------------------------------------------
    Team-specific attributes for groups already tracked in dbo.Groups (populated by the
    existing group import - see src/ImportScripts/Import-M365GroupsToSql.ps1). Populated by
    src/ImportScripts/Import-TeamsToSql.ps1 (scheduled) and refreshed immediately by
    TeamsService.CreateTeamAsync for teams created through this app.

    Safe to re-run: every step is guarded with IF NOT EXISTS.

    PREREQUISITE: run src/ImportScripts/01_CreateGroupImportTables.sql first - dbo.Teams has a
    foreign key on dbo.Groups(Id).
**********************************************************************************************/

USE [M365Manager];
GO

IF OBJECT_ID(N'dbo.Teams', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Teams
    (
        GroupId                         UNIQUEIDENTIFIER NOT NULL,
        Visibility                      NVARCHAR(32)     NULL,   -- Public / Private (Graph group.visibility)
        IsArchived                      BIT              NOT NULL CONSTRAINT DF_Teams_IsArchived DEFAULT (0),
        AllowToAddGuests                BIT              NULL,   -- Graph groupSettings (Group.Unified.Guest)
        HiddenFromAddressListsEnabled   BIT              NULL,   -- Exchange Get-UnifiedGroup
        WelcomeMessageEnabled           BIT              NULL,   -- Exchange Get-UnifiedGroup
        SharePointSiteUrl               NVARCHAR(512)    NULL,   -- Exchange Get-UnifiedGroup
        CreatedDateTime                 DATETIME2(3)     NULL,
        LastImportedAtUtc               DATETIME2(3)     NOT NULL CONSTRAINT DF_Teams_LastImportedAtUtc DEFAULT (SYSUTCDATETIME()),

        CONSTRAINT PK_Teams PRIMARY KEY CLUSTERED (GroupId),
        CONSTRAINT FK_Teams_Groups FOREIGN KEY (GroupId) REFERENCES dbo.Groups(Id)
    );
END
GO

PRINT 'M365Manager Teams overview schema is ready.';
GO
