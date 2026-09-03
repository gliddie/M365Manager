/**********************************************************************************************
    M365Manager - Room & Resource Mailboxes
    ------------------------------------------------------------------------------------------
    Two tables:

    dbo.RoomSites     - editable site/time-zone list, maintained in Settings. Replaces the legacy
                        E:\O365AdminShared\Data\RoomTimeZones.csv that RoomResourceNewForm.ps1
                        read to resolve a site code to its time zone and regional RRS admin group.

    dbo.RoomResources - overview cache for room and equipment mailboxes. Like dbo.SharedMailboxes
                        (and unlike dbo.Teams) this is NOT foreign-keyed to dbo.Groups: rooms are
                        mailboxes, not groups. Keyed by the mailbox's own ExchangeGuid. The room
                        list / delegate / users group names are recorded for reference; their
                        membership lives in Exchange and in dbo.Groups/dbo.GroupMembers via the
                        existing group import.

    Populated by src/ImportScripts/Import-RoomResourcesToSql.ps1 (scheduled) and refreshed
    immediately by RoomResourceService for rooms created or changed through this app.

    Safe to re-run: every step is guarded with IF NOT EXISTS.
**********************************************************************************************/

USE [M365Manager];
GO

IF OBJECT_ID(N'dbo.RoomSites', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.RoomSites
    (
        Id                  INT              NOT NULL IDENTITY(1,1),
        SiteCode            NVARCHAR(16)     NOT NULL,   -- e.g. "NBK" - first word of every room name
        TimeZone            NVARCHAR(128)    NOT NULL,   -- Windows time zone id, e.g. "W. Europe Standard Time"
        Region              NVARCHAR(16)     NULL,       -- legacy pick list: AP / CA / EU / LA / US
        RegionalAdminGroup  NVARCHAR(256)    NULL,       -- e.g. "MBX.EU.RRS.Admins" - FullAccess on this site's rooms

        CONSTRAINT PK_RoomSites PRIMARY KEY CLUSTERED (Id)
    );

    CREATE UNIQUE INDEX IX_RoomSites_SiteCode ON dbo.RoomSites(SiteCode);
END
GO

IF OBJECT_ID(N'dbo.RoomResources', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.RoomResources
    (
        ExchangeGuid        UNIQUEIDENTIFIER NOT NULL,
        DisplayName         NVARCHAR(256)    NOT NULL,
        Alias               NVARCHAR(128)    NOT NULL,
        PrimarySmtpAddress  NVARCHAR(256)    NOT NULL,
        ResourceKind        NVARCHAR(16)     NOT NULL,   -- Room / Equipment
        Capacity            INT              NULL,       -- rooms only
        Building            NVARCHAR(128)    NULL,
        Floor               NVARCHAR(32)     NULL,
        Office              NVARCHAR(256)    NULL,       -- "Building 3, Floor 2" - the Set-User -Office value
        SiteCode            NVARCHAR(16)     NULL,
        TimeZone            NVARCHAR(128)    NULL,
        RoomListName        NVARCHAR(256)    NULL,       -- "<SITE> Conference Rooms" / "<SITE> Restricted Rooms"
        DelegateGroup       NVARCHAR(256)    NULL,       -- MBX.<SITE>.RRS.OutOfPolicy[.<Room>].DE
        UsersGroup          NVARCHAR(256)    NULL,       -- MBX.<SITE>.RRS.OutOfPolicy.<Room>.US (restricted only)
        AccessModel         NVARCHAR(32)     NULL,       -- GeneralUseSiteDelegates / GeneralUseCustomDelegates / Restricted
        BookingPolicy       NVARCHAR(32)     NULL,       -- Standard / Hoteling / Custom
        BookingWindowInDays INT              NULL,
        MaximumDurationInMinutes INT         NULL,
        AllowRecurringMeetings BIT           NULL,
        CreatedDateTime     DATETIME2(3)     NULL,
        IsDeletedInM365     BIT              NOT NULL CONSTRAINT DF_RoomResources_IsDeleted DEFAULT (0),
        LastImportedAtUtc   DATETIME2(3)     NOT NULL CONSTRAINT DF_RoomResources_LastImportedAtUtc DEFAULT (SYSUTCDATETIME()),
        LastSeenAtUtc       DATETIME2(3)     NULL,

        CONSTRAINT PK_RoomResources PRIMARY KEY CLUSTERED (ExchangeGuid)
    );

    CREATE UNIQUE INDEX IX_RoomResources_PrimarySmtpAddress ON dbo.RoomResources(PrimarySmtpAddress);
    CREATE INDEX IX_RoomResources_SiteCode ON dbo.RoomResources(SiteCode);
END
GO

PRINT 'M365Manager Room & Resource schema is ready.';
GO
