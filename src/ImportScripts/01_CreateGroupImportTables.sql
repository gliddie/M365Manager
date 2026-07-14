/**********************************************************************************************
    M365Manager - Group Import Schema
    ------------------------------------------------------------------------------------------
    Run this script in SQL Server Management Studio (SSMS) against your SQL Server.
    It creates the tables used by the PowerShell import script for Microsoft 365 groups.

    Safe to re-run: every step is guarded with IF NOT EXISTS.
**********************************************************************************************/

IF DB_ID(N'M365Manager') IS NULL
BEGIN
    CREATE DATABASE [M365Manager];
END
GO

USE [M365Manager];
GO

IF OBJECT_ID(N'dbo.Groups', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Groups (
        Id UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,
        ExternalId NVARCHAR(256) NULL,
        DisplayName NVARCHAR(256) NULL,
        MailNickname NVARCHAR(128) NULL,
        Alias NVARCHAR(128) NULL,
        PrimarySmtpAddress NVARCHAR(256) NULL,
        GroupTypes NVARCHAR(512) NULL,
        RecipientTypeDetails NVARCHAR(128) NULL,
        MailEnabled BIT NOT NULL DEFAULT 0,
        SecurityEnabled BIT NOT NULL DEFAULT 0,
        IsAssignableToRole BIT NOT NULL DEFAULT 0,
        IsHiddenInOutlookClients BIT NOT NULL DEFAULT 0,
        IsMembershipDynamic BIT NOT NULL DEFAULT 0,
        MembershipRule NVARCHAR(4000) NULL,
        Notes NVARCHAR(4000) NULL,
        ManagedBy NVARCHAR(MAX) NULL,
        CreatedDateTime DATETIME2(3) NULL,
        UpdatedDateTime DATETIME2(3) NULL,
        LastImportedAtUtc DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        IsDeletedInM365 BIT NOT NULL DEFAULT 0,
        LastSeenAtUtc DATETIME2(3) NULL
    );
END
GO

IF OBJECT_ID(N'dbo.GroupMembers', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GroupMembers (
        Id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        GroupId UNIQUEIDENTIFIER NOT NULL,
        MemberType NVARCHAR(64) NOT NULL,
        MemberId NVARCHAR(256) NULL,
        DisplayName NVARCHAR(256) NULL,
        PrimarySmtpAddress NVARCHAR(256) NULL,
        UserPrincipalName NVARCHAR(256) NULL,
        UserType NVARCHAR(64) NULL,
        IsNested BIT NOT NULL DEFAULT 0,
        Path NVARCHAR(2048) NULL,
        LastImportedAtUtc DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_GroupMembers_Groups FOREIGN KEY (GroupId) REFERENCES dbo.Groups(Id)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Groups_PrimarySmtpAddress' AND object_id = OBJECT_ID(N'dbo.Groups'))
    CREATE INDEX IX_Groups_PrimarySmtpAddress ON dbo.Groups (PrimarySmtpAddress);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_GroupMembers_GroupId' AND object_id = OBJECT_ID(N'dbo.GroupMembers'))
    CREATE INDEX IX_GroupMembers_GroupId ON dbo.GroupMembers (GroupId);
GO

PRINT 'Group import tables are ready.';
GO
