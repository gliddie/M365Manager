/**********************************************************************************************
    M365Manager - Team Naming Convention
    ------------------------------------------------------------------------------------------
    Supports the "Create Team" feature: enforces the legacy GRP.<Location>.<Name> naming
    convention (see docs/legacy-powershell/Automation/Scripts/NewUnifiedGrp.ps1, Check-GroupName)
    by translating known acronyms/terms in the proposed name before the group/team is created.

    Safe to re-run: every step is guarded with IF NOT EXISTS.
**********************************************************************************************/

USE [M365Manager];
GO

/*------------------------------------------------------------------------------------------
  1. Acronym translation table
------------------------------------------------------------------------------------------*/
IF OBJECT_ID(N'dbo.TeamNameAcronyms', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TeamNameAcronyms
    (
        Id           INT IDENTITY(1,1) NOT NULL,
        Acronym      NVARCHAR(64)      NOT NULL,
        Translation  NVARCHAR(64)      NOT NULL,
        CONSTRAINT PK_TeamNameAcronyms PRIMARY KEY CLUSTERED (Id)
    );
END
GO

/*------------------------------------------------------------------------------------------
  2. Seed from the legacy Input-KnownAcronyms.csv (118 entries), only on first run.
     Leading/trailing spaces in Acronym/Translation are intentional - they come from the
     original list and control word-boundary replacement (see TeamNamingService).
------------------------------------------------------------------------------------------*/
IF NOT EXISTS (SELECT 1 FROM dbo.TeamNameAcronyms)
BEGIN
    INSERT INTO dbo.TeamNameAcronyms (Acronym, Translation)
    VALUES
        (N'UL.IMS', N'UL.IMS'),
        (N'CRS', N'CRS'),
        (N'eCommerce', N'eCommerce'),
        (N'LHS', N'LHS'),
        (N'HVAC', N'HVAC'),
        (N'pureLearning', N'pureLearning'),
        (N'AWS ', N'AWS '),
        (N' AWS', N' AWS'),
        (N'CEC ', N'CEC '),
        (N'MyHome', N'MyHome'),
        (N'BI Group', N'BI Group'),
        (N' BI ', N' BI '),
        (N'AES ', N'AES '),
        (N'RS CVS', N'RS CVS'),
        (N'CVS', N'CVS'),
        (N'UL TS', N'UL TS'),
        (N'NIAP', N'NIAP'),
        (N' RS ', N' RS '),
        (N' CRM ', N' CRM '),
        (N' CRM', N' CRM'),
        (N'CRM ', N'CRM '),
        (N'ECM ', N'ECM '),
        (N'Labops', N'LabOps'),
        (N'FUS ', N'FUS '),
        (N' FUS', N' FUS'),
        (N'AHL ', N'AHL '),
        (N' AHL ', N' AHL '),
        (N' IT', N' IT'),
        (N' ISE', N' ISE'),
        (N'HR ', N'HR '),
        (N'GCAR', N'GCAR'),
        (N' And ', N' and '),
        (N'OCD v', N'OCD v'),
        (N'Power BI', N'PowerBI'),
        (N' PM', N' PM'),
        (N'PM ', N' PM'),
        (N'UL', N'UL'),
        (N'ASEAN', N'ASEAN'),
        (N'EMEA', N'EMEA'),
        (N'EMEALA', N'EMEALA'),
        (N'EMEA+LA', N'EMEALA'),
        (N'CTECH', N'CTECH'),
        (N'R&I', N'RI'),
        (N'R & I', N'RI'),
        (N'RI ', N'RI '),
        (N' RI', N' RI'),
        (N'RnI', N'RI'),
        (N'R and I', N'RI'),
        (N'Retail and Insustrial', N'RI'),
        (N'C&I', N'CI'),
        (N'C & I', N'CI'),
        (N'CI ', N'CI '),
        (N' CI', N' CI'),
        (N'Commercial and Industrial', N'CI'),
        (N'C and I', N'CI'),
        (N'CnI', N'CI'),
        (N'COE', N'COE'),
        (N'CTF', N'CTF'),
        (N'CTL', N'CTL'),
        (N'ECD', N'ECD'),
        (N'DTE', N'DTE'),
        (N'BTT', N'BTT'),
        (N'CNAS', N'CNAS'),
        (N'CLS', N'CLS'),
        (N'EMC', N'EMC'),
        (N'KPI', N'KPI'),
        (N'GMA', N'GMA'),
        (N'GC', N'GC'),
        (N'PQA', N'PQA'),
        (N'IoT', N'IoT'),
        (N'DEWI-OCC', N'DEWI-OCC'),
        (N'DEWI', N'DEWI'),
        (N'McD', N'McD'),
        (N'PMO', N'PMO'),
        (N'ECG', N'ECG'),
        (N'R&D', N'R&D'),
        (N'RnD', N'R&D'),
        (N'ULU', N'ULU'),
        (N'SMA ', N'SMA '),
        (N'SOP', N'SOP'),
        (N'QA ', N'QA '),
        (N' QA', N' QA'),
        (N'QM', N'QM'),
        (N'OSHA', N'OSHA'),
        (N'ASTM', N'ASTM'),
        (N'NFPA', N'NFPA'),
        (N'IEEE', N'IEEE'),
        (N'EUCS', N'EUCS'),
        (N'FSS', N'FSS'),
        (N'GC ', N'GC '),
        (N' GC ', N' GC'),
        (N'GLP', N'GLP'),
        (N'IEC', N'IEC'),
        (N'HiTech', N'HiTech'),
        (N'EHSS', N'EHSS'),
        (N'eLearning', N'eLearning'),
        (N'WERCS', N'WERCS'),
        (N'WERCSmart', N'WERCSmart'),
        (N'USO', N'USO'),
        (N'EPT', N'EPT'),
        (N'TS STP', N'TS STP'),
        (N'CR360', N'CR360'),
        (N'FSS TMO', N'FSS TMO'),
        (N'FSS SLT', N'FSS SLT'),
        (N'GoodGuide', N'GoodGuide'),
        (N'SCS ', N'SCS '),
        (N' SCS', N' SCS'),
        (N'ULE ', N'ULE '),
        (N' ULE', N' ULE'),
        (N'IFS', N'IFS'),
        (N' ULI', N' ULI'),
        (N'ULI ', N'ULI '),
        (N'CWC', N'CWC'),
        (N'TC ', N'TC '),
        (N' TC', N' TC'),
        (N'BLST', N'BLST'),
        (N'BMS', N'BMS'),
        (N'DevOps', N'DevOps');
END
GO

PRINT 'M365Manager team naming schema is ready.';
GO
