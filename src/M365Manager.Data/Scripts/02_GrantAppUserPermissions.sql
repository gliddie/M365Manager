/**********************************************************************************************
    M365Manager - Grant an existing login access to the database
    ------------------------------------------------------------------------------------------
    Gives an EXISTING login the least-privilege rights the app needs:
        db_datareader  -> SELECT (view logs)
        db_datawriter  -> INSERT/UPDATE/DELETE (write logs)
    No schema-change (DDL) rights are granted - the schema is created by 01_CreateDatabase.sql.

    Who signs in how (since 2026-10, when the shared SQL login sfbhelper was dropped):
      - the app: Windows authentication ("Use integrated security" in Settings > SQL) through the
        AD group GLOBAL\ACL.NBK.M365Manager - that group is the login to grant here.
      - the scheduled import scripts: SQL login M365Manager, read from a credential file. It needs
        more than this script grants (the imports create/alter their tables), so it is db_owner.
      - New Hire also reads the uchelper database; grant the same login read/write there separately.

    HOW TO USE:
      1. Make sure the login already exists at server level (SSMS -> Security -> Logins). For an AD
         group: CREATE LOGIN [GLOBAL\<group>] FROM WINDOWS; This script does NOT create it.
      2. Set @LoginName below to that login's name.
      3. Run this script (it targets the M365Manager database). Safe to re-run.
**********************************************************************************************/

USE [M365Manager];
GO

DECLARE @LoginName sysname = N'GLOBAL\ACL.NBK.M365Manager';   -- <<< your existing login (AD group or SQL login)
DECLARE @sql nvarchar(max);
DECLARE @UserName sysname;
DECLARE @LoginSid varbinary(85);

/*------------------------------------------------------------------------------------------
  1. Verify the server login exists
------------------------------------------------------------------------------------------*/
SELECT @LoginSid = sid FROM sys.server_principals WHERE name = @LoginName;

IF @LoginSid IS NULL
BEGIN
    RAISERROR('Server login "%s" does not exist. Create it first (Security -> Logins), then re-run.', 16, 1, @LoginName);
    RETURN;
END

/*------------------------------------------------------------------------------------------
  2. Find the database user already mapped to this login (matched by SID, not by name).
     Handles the case where the login owns the DB and is mapped to "dbo".
------------------------------------------------------------------------------------------*/
SELECT @UserName = name
FROM sys.database_principals
WHERE type IN ('S', 'U', 'G') AND sid = @LoginSid;

IF @UserName IS NULL
BEGIN
    -- No mapping yet: create a database user for the login.
    SET @sql = N'CREATE USER ' + QUOTENAME(@LoginName) + N' FOR LOGIN ' + QUOTENAME(@LoginName) + N';';
    EXEC sp_executesql @sql;
    SET @UserName = @LoginName;
    PRINT 'Database user created: ' + @UserName;
END
ELSE
    PRINT 'Login is already mapped to database user: ' + @UserName;

/*------------------------------------------------------------------------------------------
  3. Grant read + write - unless the login is the database owner (dbo already has full access).
------------------------------------------------------------------------------------------*/
IF @UserName = N'dbo'
BEGIN
    PRINT 'This login OWNS the database (mapped to dbo) and already has full access. Nothing to grant.';
END
ELSE
BEGIN
    SET @sql =
          N'ALTER ROLE db_datareader ADD MEMBER ' + QUOTENAME(@UserName) + N';'
        + N'ALTER ROLE db_datawriter ADD MEMBER ' + QUOTENAME(@UserName) + N';';
    EXEC sp_executesql @sql;
    PRINT 'Granted db_datareader + db_datawriter to ' + @UserName + ' on M365Manager.';
END
GO

/*------------------------------------------------------------------------------------------
  4. Verify (optional): list the app user's role memberships
------------------------------------------------------------------------------------------*/
SELECT dp.name AS [User], r.name AS [Role]
FROM sys.database_role_members drm
JOIN sys.database_principals r  ON r.principal_id  = drm.role_principal_id
JOIN sys.database_principals dp ON dp.principal_id = drm.member_principal_id
ORDER BY dp.name, r.name;
GO
