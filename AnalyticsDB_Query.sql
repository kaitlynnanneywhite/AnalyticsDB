USE [DataAnalysis]
GO
/****** Object:  StoredProcedure [dbo].[MSSO_AnalyticsDB]    Script Date: 11/8/2016 12:57:17 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:          <Whitt, Finis>
-- Create date: <January, 2015>
-- Description:     <This is the script to update the Member Services Strat&Ops Analytics Tables>
-- =============================================
ALTER PROCEDURE [dbo].[MSSO_AnalyticsDB]

AS
BEGIN
/***** BEGIN MSSO ANALYTICS DATABASE CALCULATIONS *****/

/***********************************************************************************************************************************************
DROP ALL TABLES
***********************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MS_Accounts', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MS_Accounts
IF OBJECT_ID('DataAnalysis.dbo.TblData_Inst_SF', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_Inst_SF
IF OBJECT_ID('DataAnalysis.dbo.MSMK_LeadMtMs', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_LeadMtMs
IF OBJECT_ID('DataAnalysis.dbo.TblData_WarmLeadMtM', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_WarmLeadMtM
IF OBJECT_ID('DataAnalysis.dbo.MSMK_Activities_Raw', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Activities_Raw
IF OBJECT_ID('DataAnalysis.dbo.MSMK_Activities', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Activities
IF OBJECT_ID('DataAnalysis.dbo.TblData_VisitsthatCount_SF', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_VisitsthatCount_SF
IF OBJECT_ID('DataAnalysis.dbo.MSMK_NBB_Binder', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_NBB_Binder
IF OBJECT_ID('DataAnalysis.dbo.Tbl_NBBBinders_SF', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.Tbl_NBBBinders_SF
IF OBJECT_ID('DataAnalysis.dbo.MSMK_Opportunities', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Opportunities
IF OBJECT_ID('DataAnalysis.dbo.TblData_OppHistory_MktgFastTrack', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_OppHistory_MktgFastTrack
IF OBJECT_ID('DataAnalysis.dbo.MSMK_OppHistory', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_OppHistory
IF OBJECT_ID('DataAnalysis.dbo.MS_MemberHistory', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MS_MemberHistory
IF OBJECT_ID('DataAnalysis.dbo.MSMK_Leads', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Leads

/***********************************************************************************************************************************************
TopParent_Calculation
     Use: Create Top_Parent_Counter_Id field based on Account Hierarchy (because the field in SF may not be updated yet)
***********************************************************************************************************************************************/

/**
IF OBJECT_ID('DataAnalysis.dbo.MS_TopParent_Calc', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MS_TopParent_Calc

CREATE TABLE DataAnalysis.dbo.MS_TopParent_Calc (
AccountId    NVARCHAR(18), 
TopParentId  NVARCHAR(18)     )
**/

TRUNCATE TABLE DataAnalysis.dbo.MS_TopParent_Calc

--CREATE TABLE OF ALL ACCOUNTS

INSERT INTO DataAnalysis.dbo.MS_TopParent_Calc

SELECT    Account_SF.Id, 
          ISNULL( Account_SF.ParentId, Account_SF.Id )
FROM      DBAmp_SF.dbo.Account AS Account_SF

--UPDATE ACCOUNT PARENT TO PARENT'S PARENT

DECLARE @LoopCounter INT
SET @LoopCounter = 0     

WHILE     
  ( SELECT    COUNT( CASE WHEN Account_SF.ParentId IS NOT NULL 
                          THEN 1 
                          ELSE NULL END )
    FROM      DataAnalysis.dbo.MS_TopParent_Calc AS Tops
              LEFT JOIN  DBAmp_SF.dbo.Account AS Account_SF ON Tops.TopParentId = Account_SF.Id ) > 0
  AND @LoopCounter < 20
  
BEGIN
  SET @LoopCounter = @LoopCounter + 1
  
  UPDATE DataAnalysis.dbo.MS_TopParent_Calc
  SET    DataAnalysis.dbo.MS_TopParent_Calc.TopParentId = 
           ISNULL( Account_SF.ParentId, DataAnalysis.dbo.MS_TopParent_Calc.TopParentId )
  FROM   DataAnalysis.dbo.MS_TopParent_Calc
         LEFT JOIN DBAmp_SF.dbo.Account AS Account_SF ON DataAnalysis.dbo.MS_TopParent_Calc.TopParentId = Account_SF.Id
END

/***********************************************************************************************************************************************
MS_Accounts
     Use: Create account table with all commonly used fields to avoid excess joins in later queries
***********************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MS_Accounts', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MS_Accounts

SELECT    Account_SF.Id AS AccountSFID, 
          Account_SF.Counter_Id__c AS AccountCounterID,
          Account_SF.Name AS AccountName, 
          Account_SF.Account_Formal_Name__c AS AccountFormalName, 
          Account_SF.System_Status__c AS SystemStatus, 
          Account_SF.Status__c  AS [Status], 
          RT_Account.Name AS RecordType, 
          Account_Parent.Counter_Id__c AS ParentCounterID, 
          Account_TP.Counter_Id__c AS TopParentCounterID, 
          Account_SF.Primary_City__c AS PrimaryCity, 
          Account_SF.Primary_State_Province__c AS PrimaryState, 
          Account_SF.Primary_Postal_Code__c AS PrimaryPostalCode,
          Account_SF.Primary_Country__c AS PrimaryCountry, 
          Account_SF.BI_Market_Segment__c AS BIMarketSegment, 
          Account_SF.Core_Market_Segment__c AS RIMarketSegment,
          Account_SF.Blue_Book_Institution_Id__c AS BlueBookID, 
          Account_SF.Medicare_Provider_Number__c AS MedicareProviderNumber, 
          Account_SF.Carnegie_Code__c AS CarnegieCode, 
          Account_SF.Provider_Type__c AS ProviderType, 
          Account_SF.Contribution_Category__c AS ContributionCategory, 
          Account_SF.Carnegie_Classification__c AS CarnegieClassification, 
          Account_SF.Operating_Expense_In_Thousands__c AS OperatingExpense, 
          Account_SF.Bed_Size__c AS BedSize, 
          User_EA.Name AS EA, 
          User_EA.Employee_Contact_RecordID__c AS EACounterID, 
          Account_SF.Historical_Linker__c AS HistoricalLinker, 
          User_Principal.Name AS NASAPrincipal, 
          CASE WHEN RT_Account.Name LIKE 'US Provider%'
               THEN 'US Healthcare'
               ELSE RT_Account.Name END AS RecordTypeBucket  
INTO      DataAnalysis.dbo.MS_Accounts
          
FROM      DBAmp_SF.dbo.Account AS Account_SF
          INNER JOIN DBAmp_SF.dbo.RecordType AS RT_Account ON Account_SF.RecordTypeId = RT_Account.Id
          INNER JOIN DataAnalysis.dbo.MS_TopParent_Calc AS TopParent_Calc ON Account_SF.Id = TopParent_Calc.AccountId
          INNER JOIN DBAmp_SF.dbo.Account AS Account_TP ON TopParent_Calc.TopParentId = Account_TP.Id
          LEFT JOIN  DBAmp_SF.dbo.Account AS Account_Parent ON Account_SF.ParentId = Account_Parent.Id
          LEFT JOIN  DBAmp_SF.dbo.[User] AS User_EA ON Account_SF.EA_AE__C = User_EA.ID
          LEFT JOIN  DBAmp_SF.dbo.[User] AS User_Principal ON Account_SF.Principal__c = User_Principal.Id

---------------------------------------------------------------------------------------------------------------------
--Keep the old TblData_Inst_SF Table running

IF OBJECT_ID('DataAnalysis.dbo.TblData_Inst_SF', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_Inst_SF

SELECT    Accounts.AccountCounterID AS InstCounterID,
          Accounts.AccountName AS InstName, 
          Accounts.AccountFormalName AS AccountFormalName, 
          Accounts.SystemStatus AS SystemStatus, 
          Accounts.[Status] AS [Status], 
          Accounts.RecordType AS RecordType, 
          Accounts.ParentCounterID AS ParentCounterID, 
          Accounts.TopParentCounterID AS TopParentCounterID, 
          Accounts.PrimaryCity AS PrimaryCity, 
          Accounts.PrimaryState AS PrimaryState, 
          Accounts.PrimaryPostalCode AS PrimaryPostalCode,
          Accounts.PrimaryCountry AS PrimaryCountry, 
          Accounts.BIMarketSegment AS BIMarketSegment, 
          Accounts.RIMarketSegment AS RIMarketSegment,
          Accounts.BlueBookID AS BlueBookID, 
          Accounts.MedicareProviderNumber AS MedicareProviderNumber, 
          Accounts.CarnegieCode AS CarnegieCode, 
          Accounts.ProviderType AS ProviderType, 
          Accounts.ContributionCategory AS ContributionCategory, 
          Accounts.CarnegieClassification AS CarnegieClassification, 
          Accounts.OperatingExpense AS OperatingExpense, 
          Accounts.BedSize AS BedSize, 
          Accounts.EA AS EA, 
          Accounts.EACounterID AS EACounterID, 
          Accounts.HistoricalLinker AS HistoricalLinker, 
          Accounts.AccountSFID AS InstSFID, 
          Accounts.NASAPrincipal AS NASA_Principal
         
INTO      DataAnalysis.dbo.TblData_Inst_SF
FROM      DataAnalysis.dbo.MS_Accounts AS Accounts

/***********************************************************************************************************************************************
MSMK_LeadMtMs
     Use: Create a lead MtM table with all commonly used fields to avoid excess joins in later queries
***********************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MSMK_LeadMtMs', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_LeadMtMs

SELECT    AccountCalc.AccountCounterID AS LeadAccountCounterID, 
          AccountCalc.AccountSFID AS LeadAccountSFID, 
          Lead_MtM_SF.Opportunity__c AS OppSFID, 
          Opp_SF.Counter_Id__c AS OppCounterID, 
          Lead_MtM_SF.Warm_lead__c AS LeadSFID, 
          Lead_SF.Counter_ID__c AS LeadCounterID, 
          Lead_MtM_SF.Id AS LeadMtMSFID, 
          AccountCalc.AccountName AS LeadAccountName, 
          Lead_SF.Lead_Channel__c AS LeadEvent, 
          Lead_MtM_SF.Warm_Lead_Date__c AS LeadDate, 
          Lead_SF.OAB_NBB__c AS OABNBB, 
          Program_WarmLead.Counter_Id__c AS LeadProgramCounterID, 
          Program_WarmLead.Program_Acronym__c AS LeadProgram, 
          Program_Opp.Counter_Id__c AS OppProgramCounterID, 
          Program_Opp.Program_Acronym__c AS OppProgram, 
          Lead_SF.Ever_Visited__c AS EverVisited, 
          Lead_SF.Lead_Origin__c AS LeadOrigin, 
          Lead_SF.Status__c AS [Status], 
          Lead_SF.Name AS LeadName
          
INTO      DataAnalysis.dbo.MSMK_LeadMtMs
          
FROM      DBAmp_SF.dbo.Warm_Lead_MtM__c AS Lead_MtM_SF
          LEFT JOIN DBAmp_SF.dbo.Opportunity AS Opp_SF ON Lead_MtM_SF.Opportunity__c = Opp_SF.Id 
          LEFT JOIN DBAmp_SF.dbo.Warm_Lead__c AS Lead_SF ON Lead_MtM_SF.Warm_Lead__c = Lead_SF.Id 
          LEFT JOIN DBAmp_SF.dbo.Program__c AS Program_WarmLead ON Lead_SF.Primary_Program__c = Program_WarmLead.Id 
          LEFT JOIN DBAmp_SF.dbo.Program__c AS Program_Opp ON Opp_SF.Program__c = Program_Opp.Id 
          LEFT JOIN DataAnalysis.dbo.MS_Accounts AS AccountCalc ON  Lead_SF.Account__c = AccountCalc.AccountSFID
WHERE     Lead_MtM_SF.Opportunity__c IS NOT NULL 
          AND Lead_SF.RecordTypeId <> '012C0000000BkFkIAK'

---------------------------------------------------------------------------------------------------------------------
--Keep the old TblData_WarmLeadMtM Table running

IF OBJECT_ID('DataAnalysis.dbo.TblData_WarmLeadMtM', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_WarmLeadMtM

SELECT    LeadMtMs.LeadMtMSFID AS ID, 
          LeadMtMs.LeadSFID AS WarmLeadID, 
          LeadMtMs.LeadName AS WarmLeadName, 
          LeadMtMs.LeadAccountCounterID AS InstCounterID, 
          LeadMtMs.LeadEvent AS LeadChannel, 
          LeadMtMs.LeadDate AS LeadDate, 
          LeadMtMs.OABNBB AS OABNBB, 
          LeadMtMs.OppSFID AS OppID, 
          LeadMtMs.OppCounterID AS OppCounterID, 
          LeadMtMs.LeadProgramCounterID AS WLProgramCounterID, 
          LeadMtMs.LeadProgram AS WLProgram, 
          LeadMtMs.OppProgramCounterID AS OppProgramCounterID, 
          LeadMtMs.OppProgram AS OppProgram, 
          LeadMtMs.EverVisited AS EverVisited, 
          LeadMtMs.LeadOrigin AS LeadChannel_New, 
          LeadMtMs.[Status] AS [Status]
          
INTO      DataAnalysis.dbo.TblData_WarmLeadMtM
FROM      DataAnalysis.dbo.MSMK_LeadMtMs AS LeadMtMs

/***********************************************************************************************************************************************
MSMK_Activities_Raw (Base of MSMK_Activities_Raw)

     Calculated Columns Added: 
          AssignedRoleNASA
          AssignedRoleRCP
***********************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MSMK_Activities_Raw', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Activities_Raw

SELECT    AccountCalc.AccountSFID, 
          AccountCalc.AccountCounterID, 
          Opp_SF.Id AS OppSFID, 
          Opp_SF.Counter_Id__c AS OppCounterID, 
          Event_SF.Id AS EventSFID, 
          Event_SF.Counter_Id__c AS EventCounterID,
          AccountCalc.AccountName, 
          CAST ( Event_SF.ActivityDate AS DATETIME ) AS ActivityDate,
          MONTH(Event_SF.ActivityDate) AS ActivityMonth,
          YEAR(Event_SF.ActivityDate) AS ActivityYear,
          User_EventOwner.Name AS AssignedName,
          User_EventOwner.Id AS AssignedSFID, 
          User_EventOwner.Employee_Contact_RecordId__c AS AssignedCounterID,
          Event_SF.Event_Purpose__c AS EventPurpose,
          Event_SF.Event_Type__c AS EventType,
          Program_Opp.Program_Acronym__c AS Program,
          Program_Opp.Counter_Id__c AS ProgramCounterID,
          RepGrp.NewReportingGroup AS ReportingGroup,
          RepGrp.[Type] AS Business,
          CAST ( Event_SF.CreatedDate AS DATETIME ) AS EventCreatedDate, 
          AccountCalc.RecordType AS AccountRecordType, 
          AccountCalc.TopParentCounterID,
          AccountCalc.EA, 
          AccountCalc.NASAPrincipal,
          --AssignedRoleNASA
          CASE WHEN UserRole_EventOwner.Name LIKE 'NASA%' 
                    AND UserRole_EventOwner.Name NOT LIKE '%PTE%'
                    AND YEAR(Event_SF.ActivityDate) >= 2015
               THEN 1 
               ELSE 0 END 
               AS AssignedRoleNASA, 
          --AssignedRoleRCP
          CASE WHEN UserRole_EventOwner.Name LIKE 'Rev Cycle Sol%' 
               THEN 1 ELSE 0 END 
               AS AssignedRoleRCP, 
          --ActivityCategory
          CASE WHEN ( Event_SF.Event_Purpose__c IN ('Relationship Visit','Initial','Follow Up','Closing') 
                      AND Event_SF.Event_Type__c = 'In Person Visit'     )
                    OR Event_SF.Event_Type__c = 'Opportunity Visit'
               THEN 'Visit'
               WHEN Event_SF.Event_Type__c IN ('Phone Appointment','Web Visit')
                    AND Event_SF.Event_Purpose__c IN ('Initial','Follow Up PA','Closing PA','Closing')
               THEN 'Activity'
               WHEN Event_SF.Event_Type__c = 'Prospect Meeting Attendee - Count as Visit'
               THEN 'Prospect Meeting'
               ELSE NULL END 
               AS ActivityCategory,
          --CycleTime
          CASE WHEN AccountCalc.NASAPrincipal IS NULL
               THEN RepGrp.CycleTime
               ELSE RepGrp.NASACycleTime END AS CycleTime
          
INTO      DataAnalysis.dbo.MSMK_Activities_Raw
          
FROM      DBAmp_SF.dbo.[Event] AS Event_SF 
          INNER JOIN DataAnalysis.dbo.MS_Accounts AS AccountCalc ON Event_SF.AccountId = AccountCalc.AccountSFID
          LEFT JOIN  DBAmp_SF.dbo.Opportunity AS Opp_SF ON Event_SF.WhatId = Opp_SF.Id  
          LEFT JOIN  DBAmp_SF.dbo.[User] AS User_EventOwner ON Event_SF.OwnerId = User_EventOwner.Id
          LEFT JOIN  DBAmp_SF.dbo.[UserRole] AS UserRole_EventOwner ON User_EventOwner.UserRoleId = UserRole_EventOwner.Id
          LEFT JOIN  DBAmp_SF.dbo.RecordType AS RT_Event ON Event_SF.RecordTypeId = RT_Event.Id 
          LEFT JOIN  DBAmp_SF.dbo.Program__c AS Program_Opp ON Opp_SF.Program__c = Program_Opp.Id
          LEFT JOIN  DataAnalysis.dbo.ReportingHierarchyTable AS RepGrp ON Program_Opp.Program_Acronym__c = RepGrp.Program
WHERE     Event_SF.AccountId IS NOT NULL
          AND AccountCalc.TopParentCounterID <>'N00001424' --Removes all records tagged to ABC & children (goal records)
          AND Program_Opp.Program_Acronym__c IS NOT NULL
          AND Event_SF.IsChild = 0
          AND Event_SF.IsDeleted = 0 
          AND Event_SF.Cancelled_Did_Not_Occur__c = 0 
          --Visit types that count
          AND ( ( Event_SF.Event_Purpose__c IN ('Relationship Visit') 
                  AND Event_SF.Event_Type__c = 'In Person Visit')
                --Activity types that count
                OR (  Event_SF.Event_Type__c IN ('Phone Appointment','Web Visit')
                      AND Event_SF.Event_Purpose__c IN ('Initial','Follow Up','Intro PA','Follow Up PA','Closing PA','Closing') )
                --Other Events that count
                OR Event_SF.Event_Type__c IN ('Opportunity Visit - For Opportunity Goal')     )

/***********************************************************************************************************************************************
MSMK_Activities

     Calculated Columns Added:
          WarmLead_OAB
          WarmLead_NonOAB
          ColdOutreach
          CRVisitCount_RepGrp
          ProgTripCount
***********************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MSMK_Activities', 'U') IS NOT NULL
     DROP TABLE DataAnalysis.dbo.MSMK_Activities

SELECT    Activities_Raw.*,
          DATEADD(DD, Activities_Raw.CycleTime, Activities_Raw.ActivityDate) AS CTActivityDate,
          --WarmLead_OAB
          CASE WHEN Tbl_VisitsthatCount_WarmLeads.OAB_Leads > 0 THEN 1 ELSE 0 END AS WarmLead_OAB,
          --WarmLead_NonOAB
          CASE WHEN Tbl_VisitsthatCount_WarmLeads.NonOAB_Leads > 0 THEN 1 ELSE 0 END AS WarmLead_NonOAB,
          --ColdOutreach
          CASE WHEN ISNULL( Tbl_VisitsthatCount_WarmLeads.OAB_Leads, 0) = 0 
                    AND ISNULL( Tbl_VisitsthatCount_WarmLeads.NonOAB_Leads, 0) = 0 
               THEN 1 
               ELSE 0 END AS ColdOutreach, 
          --CRVisitCount_RepGrp
          CASE WHEN Qry_CRCount_Extract.EventSFID IS NULL THEN 0 ELSE 1 END AS CRVisitCount_RepGrp, 
          --ProgTripCount
          CASE WHEN PTC_Calc.AccountCounterID IS NULL THEN 0 ELSE 1 END AS ProgTripCount

INTO      DataAnalysis.dbo.MSMK_Activities

FROM      DataAnalysis.dbo.MSMK_Activities_Raw AS Activities_Raw
          LEFT JOIN  DataAnalysis.dbo.ReportingHierarchyTable AS RepGrp ON Activities_Raw.Program = RepGrp.Program
          
          --JOIN CRVisitCount_RepGrp (mostly used to ignore duplicates within a given Reporting Group)
          LEFT JOIN  (  SELECT    Activities_Raw.EventSFID
                        FROM      DataAnalysis.dbo.MSMK_Activities_Raw AS Activities_Raw
                                  LEFT JOIN  (  SELECT    Activities_Raw.AccountSFID,
                                                          Activities_Raw.ActivityMonth,
                                                          Activities_Raw.ActivityYear,
                                                          Activities_Raw.AssignedCounterID,
                                                          Activities_Raw.ReportingGroup,
                                                          MIN(Activities_Raw.EventCreatedDate) AS EventCreatedDate
                                                FROM      DataAnalysis.dbo.MSMK_Activities_Raw AS Activities_Raw
                                                WHERE     Activities_Raw.ActivityCategory = 'Visit'
                                                          AND ISNULL( Activities_Raw.EventPurpose, '') NOT IN ('Follow Up')
                                                GROUP BY  Activities_Raw.AccountSFID,
                                                          Activities_Raw.ActivityMonth,
                                                          Activities_Raw.ActivityYear,
                                                          Activities_Raw.AssignedCounterID,
                                                          Activities_Raw.ReportingGroup     ) 
                                             AS Qry_CRCount
                                                  --In the event of duplicates of Account+RepGrp+Assigned+Month&Year, only first visit created counts
                                                   --Exclude Follow Up & Closing visits from count 
                                                ON Activities_Raw.EventCreatedDate = Qry_CRCount.EventCreatedDate
                                                   AND Activities_Raw.AccountSFID = Qry_CRCount.AccountSFID 
                                                   AND Activities_Raw.ReportingGroup = Qry_CRCount.ReportingGroup 
                                                   AND Activities_Raw.AssignedCounterID = Qry_CRCount.AssignedCounterID 
                                                   AND Activities_Raw.ActivityYear = Qry_CRCount.ActivityYear 
                                                   AND Activities_Raw.ActivityMonth = Qry_CRCount.ActivityMonth
                                                   AND Activities_Raw.ActivityCategory = 'Visit'
                                                   AND ISNULL( Activities_Raw.EventPurpose, '') NOT IN ('Follow Up')
                        WHERE     Qry_CRCount.AccountSFID IS NOT NULL     ) 
                     AS Qry_CRCount_Extract 
                        ON Activities_Raw.EventSFID = Qry_CRCount_Extract.EventSFID
                     
          --JOIN Leads Passed prior to Opp First Visited Date
          LEFT JOIN  (  SELECT    VisitsRaw_x_Opp.OppCounterID AS OppCounterID, 
                                  COUNT( CASE WHEN LeadMtM_x_OppChannel.LeadChannel = 'OAB' THEN 1 ELSE NULL END ) AS OAB_Leads, 
                                  COUNT( CASE WHEN LeadMtM_x_OppChannel.LeadChannel = 'Non-OAB' THEN 0 ELSE NULL END ) AS NonOAB_Leads
                        FROM      --Group VisitsRaw by Opportunity to find first visit date
                                  (  SELECT    Activities_Raw.OppCounterID AS OppCounterID,
                                               MIN (Activities_Raw.ActivityDate) AS MinActivityDate
                                     FROM      DataAnalysis.dbo.MSMK_Activities_Raw AS Activities_Raw
                                     WHERE     Activities_Raw.ActivityCategory = 'Visit'
                                     GROUP BY  Activities_Raw.OppCounterID ) 
                                  AS VisitsRaw_x_Opp
                                  LEFT JOIN  DBAmp_SF.dbo.Opportunity AS Opp_SF 
                                             ON VisitsRaw_x_Opp.OppCounterID = Opp_SF.Counter_Id__c 
                                  --Lead MtMs grouped by OAB vs. Non-OAB
                                  LEFT JOIN  (  SELECT    Lead_MtM_SF.Opportunity__c AS OppID, 
                                                          Lead_MtM_SF.Warm_Lead_Date__c AS LeadDate, 
                                                          CASE WHEN Lead_MtM_SF.Warm_Lead_Lead_Channel__c = 'OAB' THEN 'OAB' ELSE 'Non-OAB' END AS LeadChannel, 
                                                          Lead_MtM_SF.Warm_Lead__c AS WarmLeadID
                                                FROM      DBAmp_SF.dbo.Warm_Lead_MtM__c AS Lead_MtM_SF
                                                WHERE     Lead_MtM_SF.Opportunity__c IS NOT NULL
                                                          AND Lead_MtM_SF.Warm_Lead_Lead_Channel__c IS NOT NULL
                                                GROUP BY  Lead_MtM_SF.Opportunity__C, 
                                                          Lead_MtM_SF.Warm_Lead_Date__c, 
                                                          CASE WHEN Lead_MtM_SF.Warm_Lead_Lead_Channel__c = 'OAB' THEN 'OAB' ELSE 'Non-OAB' END, 
                                                          Lead_MtM_SF.Warm_Lead__c     ) 
                                             AS LeadMtM_x_OppChannel 
                                                ON Opp_SF.ID = LeadMtM_x_OppChannel.OppID
                        WHERE     LeadMtM_x_OppChannel.LeadDate <= VisitsRaw_x_Opp.MinActivityDate
                        GROUP BY  VisitsRaw_x_Opp.OppCounterID ) 
                     AS Tbl_VisitsthatCount_WarmLeads 
                        ON Activities_Raw.OppCounterID = Tbl_VisitsthatCount_WarmLeads.OppCounterID 
                     
          --JOIN ProgTripCount (Ignores duplicates for a given Program and adds a few additional criteria for the count used in Exec Sales)
          LEFT JOIN  (  SELECT     Activities_Raw.AccountCounterID, 
                                   Activities_Raw.ProgramCounterID, 
                                   Activities_Raw.ActivityDate, 
                                   MIN(Activities_Raw.EventCreatedDate) AS MinCreated
                        FROM       DataAnalysis.dbo.MSMK_Activities_Raw AS Activities_Raw
                        WHERE      --Exclude NASA visits
                                   ( Activities_Raw.AssignedRoleNASA = 0
                                     OR ( Activities_Raw.AssignedName IN ('Sean Tivnan','Samantha Goldman') 
                                          AND Activities_Raw.ActivityDate >= '7/1/2015'     
                                          AND Activities_Raw.Program != 'BIXX' ) )
                                   --Only count Relationship Visits for Rev Cycle Principals 
                                   AND ( Activities_Raw.AssignedRoleRCP = 1 
                                         OR Activities_Raw.EventPurpose = 'Relationship Visit'     )
                                   AND Activities_Raw.ActivityCategory = 'Visit'
                        GROUP BY   Activities_Raw.AccountCounterID, 
                                   Activities_Raw.ProgramCounterID, 
                                   Activities_Raw.ActivityDate     ) 
                     AS PTC_Calc 
                        ON Activities_Raw.AccountCounterID = PTC_Calc.AccountCounterID 
                           AND Activities_Raw.ProgramCounterID = PTC_Calc.ProgramCounterID 
                           AND Activities_Raw.ActivityDate = PTC_Calc.ActivityDate 
                           AND Activities_Raw.EventCreatedDate = PTC_Calc.MinCreated
                           AND Activities_Raw.ActivityCategory = 'Visit' 
ORDER BY  Activities_Raw.ActivityDate DESC

--DROP VISITS_RAW TABLE
IF OBJECT_ID('DataAnalysis.dbo.MSMK_Activities_Raw', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Activities_Raw

---------------------------------------------------------------------------------------------------------------------
--Keep the old TblData_VisitsthatCount_SF Table running

IF OBJECT_ID('DataAnalysis.dbo.TblData_VisitsthatCount_SF', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_VisitsthatCount_SF

SELECT    Visits.EventSFID AS Id, 
          Visits.AccountSFID AS InstSFID,
          Visits.AccountCounterID AS InstCounterID,
          Visits.AccountName AS InstName,
          Visits.AccountRecordType AS InstType,
          Visits.OppSFID,
          Visits.OppCounterID,
          Visits.ActivityDate,
          Visits.ActivityMonth,
          Visits.ActivityYear,
          Visits.AssignedName,
          Visits.AssignedCounterID,
          Visits.EventPurpose,
          Visits.EventType,
          Visits.Program,
          Visits.ProgramCounterID,
          Visits.EventCounterID AS VisitCounterID,
          AccountCalc.HistoricalLinker AS InstHistLinker,
          Visits.Program AS CleanedProgram,
          Visits.ReportingGroup,
          Visits.Business,
          Visits.CRVisitCount_RepGrp, 
          Visits.EventCreatedDate AS VisitCreatedDate, 
          Visits.CycleTime, 
          Visits.CTActivityDate, 
          Visits.WarmLead_OAB, 
          Visits.WarmLead_NonOAB, 
          Visits.ColdOutreach, 
          Visits.AssignedRoleNASA, 
          Visits.AssignedRoleRCP, 
          Visits.ProgTripCount, 
          Visits.EA, 
          Visits.NASAPrincipal
          
INTO      DataAnalysis.dbo.TblData_VisitsthatCount_SF
FROM      DataAnalysis.dbo.MSMK_Activities AS Visits
          LEFT JOIN     DataAnalysis.dbo.MS_Accounts AS AccountCalc ON Visits.AccountSFID = AccountCalc.AccountSFID
WHERE     Visits.ActivityCategory = 'Visits'

/***********************************************************************************************************************************************
MSMK_NBB_Binder
     Use: NBB records filtered down to only those that are posted
***********************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MSMK_NBB_Binder', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_NBB_Binder

SELECT    AccountCalc.AccountCounterID, 
          AccountCalc.AccountSFID, 
          Opp_SF.Counter_ID__c AS OppCounterID, 
          Opp_SF.Id AS OppSFID, 
          NBB_SF.Id AS NBBSFID, 
          AccountCalc.AccountName, 
          Program_SF.Program_Acronym__c AS Program, 
          CAST( NBB_SF.SA_EA_Date__c AS DATETIME ) AS EADate, 
          CAST( NBB_SF.ATL_Date__c AS DATETIME ) AS ATLDate, 
          User_Marketer_NBB.Name AS Marketer, 
          NBB_SF.NBB__c AS NBB, 
          NBB_SF.Unit__c AS Units, 
          NBB_SF.Status__c AS [Status], 
          CAST( NBB_SF.NA_Date__c AS DATETIME ) AS NADate, 
          CAST( NBB_SF.Binder_Year__c AS INT ) AS Binder, 
          Contract_ABC_SF.Counter_ID__c AS ContractCounterID, 
          Program_SF.Counter_ID__c AS ProgramCounterID, 
          User_Marketer_NBB.Employee_Contact_RecordID__c AS MarketerCounterID, 
          NBB_SF.NBB_Type__c AS NBBType, 
          CAST( NBB_SF.Upsell__c AS INT ) AS Upsell, 
          RepGrp.NewReportingGroup AS ReportingGroup, 
          RepGrp.[Type] AS Business,
          NBB_x_Opp.TotalOppNBB AS TotalOppNBB, 
          NBB_x_Opp.LrgContractThresh AS LgContractThreshold , 
          CASE WHEN NBB_x_Opp.TotalOppNBB >= NBB_x_Opp.LrgContractThresh THEN 1 ELSE 0 END AS LgContract,
          CAST( NBB_SF.Binder_Date__c AS DATETIME ) AS BinderDate, 
          AccountCalc.RecordType AS AccountRecordType, 
          AccountCalc.TopParentCounterID, 
          AccountCalc.EA, 
          AccountCalc.NASAPrincipal
          
INTO      DataAnalysis.dbo.MSMK_NBB_Binder
          
FROM      DBAmp_SF.dbo.NBB__c AS NBB_SF
          INNER JOIN  DBAmp_SF.dbo.RecordType AS RT_NBB ON NBB_SF.RecordTypeId = RT_NBB.Id
          LEFT JOIN   DBAmp_SF.dbo.Opportunity AS Opp_SF ON NBB_SF.Opportunity__c = Opp_SF.Id
          LEFT JOIN   DataAnalysis.dbo.MS_Accounts AS AccountCalc ON Opp_SF.AccountId = AccountCalc.AccountSFID
          LEFT JOIN   DBAmp_SF.dbo.[User] AS User_Marketer_NBB ON NBB_SF.Marketer__c = User_Marketer_NBB.Id
          LEFT JOIN   DBAmp_SF.dbo.Program__c AS Program_SF ON NBB_SF.Program__c = Program_SF.Id
          LEFT JOIN   DBAmp_SF.dbo.Contract__c AS Contract_ABC_SF ON NBB_SF.Contract__c = Contract_ABC_SF.Id
          LEFT JOIN   DataAnalysis.dbo.ReportingHierarchyTable AS RepGrp ON Program_SF.Program_Acronym__c = RepGrp.Program 
          LEFT JOIN   (  SELECT    NBB_SF.Opportunity__c AS OppSFId, 
                                   SUM(NBB_SF.NBB__c) AS TotalOppNBB, 
                                   MAX(Program_SF.Large_Contract_Threshold__c) AS LrgContractThresh
                         FROM      DBAmp_SF.dbo.NBB__c AS NBB_SF
                                   INNER JOIN DBAmp_SF.dbo.RecordType AS RT_NBB ON NBB_SF.RecordTypeId = RT_NBB.Id
                                   LEFT JOIN  DBAmp_SF.dbo.Program__c AS Program_SF ON NBB_SF.Program__c = Program_SF.Id
                         WHERE     NBB_SF.Opportunity__c IS NOT NULL
                                   AND NBB_SF.NBB_Type__c IN ('Base Fee', 'Shadow Credit')
                                   AND RT_NBB.Name = 'Standard'
                         GROUP BY  NBB_SF.Opportunity__c     ) 
                      AS NBB_x_Opp ON Opp_SF.Id = NBB_x_Opp.OppSFId
WHERE     NBB_SF.NBB_Type__c IN ('Base Fee', 'Shadow Credit', 'Posted Risk')
          AND RT_NBB.Name = 'Standard'

---------------------------------------------------------------------------------------------------------------------
--Keep the old Tbl_NBBBinders_SF Table running

IF OBJECT_ID('DataAnalysis.dbo.Tbl_NBBBinders_SF', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.Tbl_NBBBinders_SF

SELECT    Contract_SF.Historical_Linker__c AS Opportunity,
          AccountCalc.HistoricalLinker AS OrgID, 
          Binder.EADate, 
          Binder.ATLDate, 
          Binder.Marketer, 
          Binder.[Status], 
          Binder.Program, 
          Binder.AccountName AS Institution, 
          Binder.NBB AS Revenue, 
          Binder.Units, 
          CAST( NULL AS NVARCHAR(9) ) AS Adjustment, 
          Binder.NADate, 
          Binder.Binder, 
          NULL AS Travel, 
          Binder.NBBSFID, 
          Binder.OppCounterID, 
          Binder.ContractCounterID, 
          Binder.AccountCounterID AS InstCounterID, 
          Binder.AccountSFID AS InstSFID, 
          Binder.ProgramCounterID, 
          Binder.MarketerCounterID, 
          Binder.NBBType, 
          Binder.Upsell, 
          Binder.ReportingGroup, 
          Binder.Business,
          Binder.TotalOppNBB AS TotalOppRevenue, 
          Binder.LgContractThreshold, 
          Binder.LgContract,
          Binder.BinderDate, 
          Binder.AccountRecordType AS InstRecordType, 
          Binder.EA, 
          Binder.NASAPrincipal
INTO      DataAnalysis.dbo.Tbl_NBBBinders_SF
FROM      DataAnalysis.dbo.MSMK_NBB_Binder AS Binder
          LEFT JOIN  DBAmp_SF.dbo.Contract__c AS Contract_SF ON Binder.ContractCounterID = Contract_SF.Counter_ID__c
          LEFT JOIN  DataAnalysis.dbo.MS_Accounts AS AccountCalc ON Binder.AccountSFID = AccountCalc.AccountSFID

/*************************************************************************************************************************************************************
MSMK_Opportunities

     1. Visit to Eval
     2. Large Contract
     3. Cycle Time
     4. Misc Opp Metrics
**************************************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MSMK_Opportunities', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Opportunities

SELECT    AccountCalc.AccountSFID,
          AccountCalc.AccountCounterID, 
          Opp_SF.Id AS OppSFID,
          Opp_SF.Counter_ID__c AS OppCounterID,
          AccountCalc.AccountName,
          Opp_Marketer.Name AS Marketer,
          Opp_SF.StageName AS Stage, 
          Opp_SF.Amount AS ProposalValue, 
          Opp_SF.Probability, 
          CAST( Opp_SF.CloseDate AS DATETIME ) AS CloseDate, 
          Opp_Program.Program_Acronym__c AS Program, 
          RepGrp.NewReportingGroup AS ReportingGroup, 
          RepGrp.[Type] AS BusinessLine, 
          Opp_CreatedBy.Name AS CreatedBy,
          AccountCalc.TopParentCounterID,
          AccountCalc.NASAPrincipal, 
          AccountCalc.EA, 
          Opp_SF.CreatedDate AS OppCreatedDate, 
          
          -- Visit_x_Opp Sub-Query
          CAST ( Activities_x_Opp.MinVisitDate AS DATETIME ) AS FirstVisitDate,
          CAST ( Activities_x_Opp.MaxVisitDate AS DATETIME ) AS LastVisitDate,
          CASE WHEN Activities_x_Opp.MinVisitDate IS NULL THEN 0 ELSE 1 END AS Visited, 
          CAST ( Activities_x_Opp.MinEventDate AS DATETIME ) AS FirstEventDate,
          CAST ( Activities_x_Opp.MaxEventDate AS DATETIME ) AS LastEventDate,
          CAST ( Activities_x_Opp.MinPADate AS DATETIME ) AS FirstPADate,
          CAST ( Activities_x_Opp.MaxPADate AS DATETIME ) AS LastPADate,
          CASE WHEN Activities_x_Opp.TotalPGTVisits > 0 THEN 1 ELSE 0 END AS ProgTripCount_Visited, 
          CASE WHEN Activities_x_Opp.TotalCRVCVisits > 0 THEN 1 ELSE 0 END AS CRVisitCount_Visited, 
          
          -- NBB_x_Opp Sub-Query
          NBB_x_Opp.NBB_Total AS NBBTotal, 
          NBB_x_Opp.Units_Total AS UnitsTotal, 
          
          -- History_x_Opp Sub-Query
          CASE WHEN     History_x_Opp.FastTrackDate     IS NULL     THEN 0 ELSE     1 END AS Evaluated, 
          CAST ( History_x_Opp.FastTrackDate AS DATETIME ) AS FastTrackDate, 
          CAST ( History_x_Opp.SlowTrackDate AS DATETIME ) AS SlowTrackDate, 
          CAST ( History_x_Opp.SendToFinanceDate AS DATETIME ) AS SendToFinanceDate, 
          CAST ( History_x_Opp.[Pipeline<20%Date] AS DATETIME ) AS [Pipeline<20%Date], 
          CAST ( History_x_Opp.[Pipeline20%Date] AS DATETIME ) AS [Pipeline20%Date], 
          CAST ( History_x_Opp.[Pipeline40%Date] AS DATETIME ) AS [Pipeline40%Date], 
          CAST ( History_x_Opp.[Pipeline60%Date] AS DATETIME ) AS [Pipeline60%Date], 
          CAST ( History_x_Opp.[Pipeline80%Date] AS DATETIME ) AS [Pipeline80%Date],
          CASE WHEN     History_x_Opp.Max_ProposalValue     > 400000 THEN 1 ELSE 0 END AS EverLC,
          History_x_Opp.Max_ProposalValue AS MaxProposalValue, 
          History_x_Opp.Max_Probability AS MaxProbability, 
          
          -- Cycle Time Fields
          DATEDIFF(DD, Activities_x_Opp.MinVisitDate, Opp_FirstATLDate.First_ATL_Date) AS CTDays_FirstVisit,
          DATEDIFF(DD, Activities_x_Opp.MaxVisitDate, Opp_FirstATLDate.First_ATL_Date) AS CTDays_LastVisit,
          NULL AS Visits_Within_CycleTime,
          CAST ( Opp_FirstATLDate.First_ATL_Date AS DATETIME ) AS FirstATLDate, 
          CASE WHEN AccountCalc.NASAPrincipal IS NULL
               THEN RepGrp.CycleTime
               ELSE RepGrp.NASACycleTime END AS CycleTime, 
          CAST ( NULL AS DATETIME ) AS CT_FirstVisitDate, 
          CAST ( NULL AS DATETIME ) AS CT_FastTrackDate, 

          --Outreach fields, filled out in update
          CAST( 0 AS FLOAT ) AS OutreachCount, 
          CAST( NULL AS DATETIME ) AS FirstOutreachDate, 
          CAST( NULL AS DATETIME ) AS LastOutreachDate, 
          CAST( NULL AS FLOAT ) AS SMG, 

          --MtM/Lead Fields
          MTM_x_Opp.MTMCount,
          MTM_x_Opp.FirstLeadCreatedDate,
          MTM_x_Opp.FirstLeadDate,
          MTM_x_Opp.MtMOABCount AS OAB

INTO      DataAnalysis.dbo.MSMK_Opportunities
FROM      DBAmp_SF.dbo.Opportunity AS Opp_SF
          LEFT JOIN  Dbamp_SF.dbo.RecordType AS Opp_RT ON Opp_SF.RecordTypeID = Opp_RT.ID
          LEFT JOIN  DBAmp_SF.dbo.Program__c AS Opp_Program ON Opp_SF.Program__c = Opp_Program.Id
          LEFT JOIN  DataAnalysis.dbo.ReportingHierarchyTable AS RepGrp ON Opp_Program.Program_Acronym__c = RepGrp.Program
          LEFT JOIN  Dbamp_SF.dbo.[User] AS Opp_Marketer ON Opp_SF.Marketer__c = Opp_Marketer.Id
          LEFT JOIN  DBAmp_SF.dbo.[User] AS Opp_CreatedBy ON Opp_SF.CreatedById = Opp_CreatedBy.Id
          LEFT JOIN  DataAnalysis.dbo.MS_Accounts AS AccountCalc ON Opp_SF.AccountId = AccountCalc.AccountSFID

          ---------------------------------------------------------------------------------------------------------------------
          --Activities_x_Opp: Group Activities by Opportunity

          LEFT JOIN  (  SELECT    Activities.OppCounterID AS OppCounterID,
                                  MIN ( CASE WHEN Activities.ActivityCategory = 'Visit' THEN Activities.ActivityDate ELSE NULL END ) AS MinVisitDate,
                                  MAX ( CASE WHEN Activities.ActivityCategory = 'Visit' THEN Activities.ActivityDate ELSE NULL END ) AS MaxVisitDate,
                                  SUM ( CASE WHEN Activities.ActivityCategory = 'Visit' THEN Activities.ProgTripCount ELSE NULL END ) AS TotalPGTVisits,
                                  SUM ( CASE WHEN Activities.ActivityCategory = 'Visit' THEN Activities.CRVisitCount_RepGrp ELSE NULL END ) AS TotalCRVCVisits, 
                                  MIN ( Activities.ActivityDate ) AS MinEventDate, 
                                  MAX ( Activities.ActivityDate ) AS MaxEventDate,
                                  MIN ( CASE WHEN Activities.EventType = 'Phone Appointment' THEN Activities.ActivityDate ELSE NULL END ) AS MinPADate,
                                  MAX ( CASE WHEN Activities.EventType = 'Phone Appointment' THEN Activities.ActivityDate ELSE NULL END ) AS MaxPADate
                        FROM      DataAnalysis.dbo.MSMK_Activities AS Activities
                        GROUP BY  Activities.OppCounterID          ) 
                     AS Activities_x_Opp ON Opp_SF.Counter_Id__c = Activities_x_Opp.OppCounterID

          ---------------------------------------------------------------------------------------------------------------------
          --NBB_x_Opp: Group NBB by Opportunity

          LEFT JOIN  (  SELECT    Binder.OppCounterID AS OppCounterID,
                                  SUM (Binder.NBB) AS NBB_Total,
                                  SUM (Binder.Units) AS Units_Total
                        FROM      DataAnalysis.dbo.MSMK_NBB_Binder AS Binder
                        GROUP BY  Binder.OppCounterID          ) 
                     AS NBB_x_Opp ON Opp_SF.Counter_Id__c = NBB_x_Opp.OppCounterID

          ---------------------------------------------------------------------------------------------------------------------
          --History_x_Opp: Group OppHistory by Opportunity

          LEFT JOIN  (  SELECT    Opp_History.OpportunityID,
                                  
                                  --FastTrackDate
                                  MIN( CASE WHEN Opp_History.StageName = 'Active in FastTrack' 
                                            THEN Opp_History.CREATEDDATE ELSE NULL END ) 
                                            AS FastTrackDate,
                                  --SendToFinanceDate
                                  MIN( CASE WHEN Opp_History.StageName = 'Contract Received - Send to Finance' 
                                            THEN Opp_History.CREATEDDATE ELSE NULL END ) 
                                            AS SendToFinanceDate,
                                  --SlowTrackDate
                                  MIN( CASE WHEN Opp_History.StageName = 'Active in Slowtrack' 
                                            THEN Opp_History.CREATEDDATE ELSE NULL END ) 
                                            AS SlowTrackDate,
                                  --[Pipeline<20%Date]
                                  MIN( CASE WHEN Opp_History.Probability < 20     
                                                 AND Opp_History.StageName IN ('Active in FastTrack','Verbal Yes') 
                                            THEN Opp_History.CreatedDate ELSE NULL END ) 
                                            AS [Pipeline<20%Date],
                                  --[Pipeline20%Date]
                                  MIN( CASE WHEN Opp_History.Probability >= 20 
                                                 AND Opp_History.StageName IN ('Active in FastTrack','Verbal Yes') 
                                            THEN Opp_History.CreatedDate ELSE NULL END ) 
                                            AS [Pipeline20%Date],
                                  --[Pipeline40%Date]
                                  MIN( CASE WHEN Opp_History.Probability >= 40 
                                                 AND Opp_History.StageName IN ('Active in FastTrack','Verbal Yes') 
                                            THEN Opp_History.CreatedDate ELSE NULL END ) 
                                            AS [Pipeline40%Date],
                                  --[Pipeline60%Date]
                                  MIN( CASE WHEN Opp_History.Probability >= 60 
                                                 AND Opp_History.StageName IN ('Active in FastTrack','Verbal Yes') 
                                            THEN Opp_History.CreatedDate ELSE NULL END ) 
                                            AS [Pipeline60%Date],
                                  --[Pipeline80%Date]
                                  MIN( CASE WHEN Opp_History.Probability >= 80 
                                                 AND Opp_History.StageName IN ('Active in FastTrack','Verbal Yes') 
                                            THEN Opp_History.CreatedDate ELSE NULL END ) 
                                            AS [Pipeline80%Date],
                                  --Max_ProposalValue
                                  MAX( Opp_History.Amount ) AS Max_ProposalValue,
                                  --Max_Probability
                                  MAX( Opp_History.Probability ) AS Max_Probability
                        FROM      DBAmp_SF.dbo.OpportunityHISTORY AS Opp_History
                                  LEFT JOIN DBAmp_SF.dbo.Opportunity AS Opp_SF ON Opp_History.OpportunityID = Opp_SF.Id
                        GROUP BY  Opp_History.OpportunityID  ) 
                     AS History_x_Opp ON Opp_SF.ID = History_x_Opp.OpportunityID          

          ---------------------------------------------------------------------------------------------------------------------
          --Opp_FirstATLDate: First date on which an Opportunities total NBB > 0

          LEFT JOIN  (  SELECT    MIN(Opp_NBB.ATLDate) AS First_ATL_Date,
                                  Opp_NBB.OppCounterID
                        FROM      (  SELECT    NBB_Binder.ATLDate,
                                               NBB_Binder.OppCounterID,
                                               SUM(NBB_Binder.NBB) AS ATL_NBB
                                     FROM      DataAnalysis.dbo.MSMK_NBB_Binder AS NBB_Binder
                                     GROUP BY  NBB_Binder.ATLDate,
                                               NBB_Binder.OppCounterID )
                                  AS Opp_NBB
                        WHERE     Opp_Nbb.ATL_NBB > 0
                        GROUP BY  Opp_NBB.OppCounterID          )
                     AS Opp_FirstATLDate ON Opp_SF.Counter_ID__c = Opp_FirstATLDate.OppCounterID

          ---------------------------------------------------------------------------------------------------------------------
          --History_x_Opp: Group OppHistory by Opportunity
          
          LEFT JOIN (  SELECT    MTM.Opportunity__c AS OppSFID, 
                                 
                                 --MtMCount
                                 COUNT( MTM.Id ) AS MtMCount, 
                                 --MtMOABCount
                                 SUM( CASE WHEN Leads.Lead_Origin__c = 'OAB'
                                                  AND Leads.Lead_Channel__c = 'Direct to Visit'
                                             THEN 1 ELSE 0 END ) AS MtMOABCount,
                                 
                                 --FirstLeadCreatedDate
                                 MIN( Leads.CreatedDate ) AS FirstLeadCreatedDate, 
                                 --FirstOABLeadCreatedDate
                                 COUNT( CASE WHEN Leads.Lead_Origin__c = 'OAB'
                                                  AND Leads.Lead_Channel__c = 'Direct to Visit'
                                             THEN Leads.CreatedDate ELSE NULL END ) AS FirstOABLeadCreatedDate,
                                 
                                 --FirstLeadCreatedDate
                                 MIN( Leads.Lead_Date__c ) AS FirstLeadDate, 
                                 --FirstOABLeadCreatedDate
                                 MIN( CASE WHEN Leads.Lead_Origin__c = 'OAB'
                                                  AND Leads.Lead_Channel__c = 'Direct to Visit'
                                             THEN Leads.Lead_Date__c ELSE NULL END ) AS FirstOABLeadDate
                                 
                       FROM      DBAmp_SF.dbo.Warm_Lead_MtM__c AS MTM
                                 INNER JOIN DBAmp_SF.dbo.Warm_Lead__c AS Leads ON Leads.Id = MTM.Warm_Lead__c
                       GROUP BY  MTM.Opportunity__c  )
                    AS MTM_x_Opp ON MTM_x_Opp.OppSFID = Opp_SF.Id

WHERE     Opp_RT.Name IN ('RI Marketing','PT Marketing','Consulting & Management')

---------------------------------------------------------------------------------------------------------------------
--Visits_Within_CycleTime: Count number of visits between first visit and first ATL

UPDATE DataAnalysis.dbo.MSMK_Opportunities          
SET    Visits_Within_CycleTime = Visits_Within_CycleTime.Number_of_Visits, 
       CT_FirstVisitDate = DATEADD( DD, MSMK_Opportunities.CycleTime, MSMK_Opportunities.FirstVisitDate ), 
       CT_FastTrackDate = DATEADD( DD, MSMK_Opportunities.CycleTime, MSMK_Opportunities.FastTrackDate )
       
FROM   DataAnalysis.dbo.MSMK_Opportunities
       LEFT JOIN (  SELECT    Opps.OppSFID, 
                              COUNT(Activities.EventSFID) AS Number_of_Visits
                    FROM      DataAnalysis.dbo.MSMK_Opportunities AS Opps
                              LEFT JOIN DataAnalysis.dbo.MSMK_Activities AS Activities ON Opps.OppCounterID = Activities.OppCounterID
                    WHERE     Activities.ActivityDate >= Opps.FirstVisitDate
                              AND Activities.ActivityDate <= Opps.FirstATLDate
                              AND Activities.ActivityCategory = 'Visits'
                    GROUP BY  Opps.OppSFID  )
                 AS Visits_Within_CycleTime
                    ON MSMK_Opportunities.OppSFID = Visits_Within_CycleTime.OppSFID

UPDATE DataAnalysis.dbo.MSMK_Opportunities
SET    OutreachCount = Task_x_Opp.OutreachCount, 
       FirstOutreachDate = Task_x_Opp.FirstOutreachDate, 
       LastOutreachDate = Task_x_Opp.LastOutreachDate
FROM   DataAnalysis.dbo.MSMK_Opportunities AS Opps
       LEFT JOIN  (  SELECT    Task_SF.WhatId,
                               COUNT(Task_SF.WhatId) AS OutreachCount,
                               MIN(Task_SF.ActivityDate) AS FirstOutreachDate,
                               MAX(Task_SF.ActivityDate) AS LastOutreachDate
                     
                     FROM      DBAmp_SF.dbo.Task AS Task_SF
                               INNER JOIN DBAmp_SF.dbo.RecordType AS RT_Task ON Task_SF.RecordTypeId = RT_Task.Id
                     WHERE     Task_SF.[Status]= 'Completed'
                               AND Task_SF.ActivityDate IS NOT NULL
                               AND Task_SF.Event_Purpose__c IN ( 'Initial','Intro','Schedule','Follow Up' )
                               AND RT_Task.Name IN ( 'Task PT Marketing','Task RI Marketing' )
                     GROUP BY  Task_SF.WhatId )
                  AS Task_x_Opp ON Opps.OppSFID = Task_x_Opp.WhatId

UPDATE DataAnalysis.dbo.MSMK_Opportunities
SET    SMG = CASE WHEN ISNULL( FirstOutreachDate, OppCreatedDate) > FirstLeadCreatedDate
                       AND MTMCount IS NOT NULL 
                  THEN 1 ELSE 0 END

---------------------------------------------------------------------------------------------------------------------
--Keep the old TblData_OppHistory_MktgFastTrack Table running

IF OBJECT_ID('DataAnalysis.dbo.TblData_OppHistory_MktgFastTrack', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.TblData_OppHistory_MktgFastTrack

SELECT    Opps.OppSFID, 
          Opps.OppCounterId, 
          Opps.AccountCounterID AS InstCounterId, 
          Opps.AccountName AS InstName, 
          Opps.Program AS Program, 
          Program_SF.New_Business_Group__c AS NewBusinessGroup, 
          Program_SF.New_Business_Vertical__c AS NewBusinessVertical, 
          Program_SF.New_Business_Business__c AS NewBusinessBusiness, 
          Opps.Marketer, 
          Opps.Stage AS StageName, 
          Opps.Probability, 
          Opps.CloseDate, 
          Opps.ProposalValue, 
          Opps.FastTrackDate, 
          Opps.[Pipeline<20%Date], 
          Opps.[Pipeline20%Date], 
          Opps.[Pipeline40%Date], 
          Opps.[Pipeline60%Date], 
          Opps.[Pipeline80%Date], 
          Opps.SlowTrackDate, 
          Opps.SendToFinanceDate, 
          Opps.TopParentCounterID, 
          Opps.ReportingGroup, 
          Opps.BusinessLine, 
          Opps.NASAPrincipal, 
          Opps.EA, 
          Opps.FirstVisitDate, 
          Opps.LastVisitDate, 
          Opps.CT_FirstVisitDate, 
          Opps.CT_FastTrackDate, 
          Opps.CTDays_FirstVisit, 
          Opps.CTDays_LastVisit, 
          Opps.Visited, 
          Opps.Evaluated, 
          Opps.NBBTotal, 
          Opps.UnitsTotal, 
          Opps.FirstEventDate, 
          Opps.LastEventDate, 
          Opps.FirstPADate, 
          Opps.LastPADate, 
          Opps.ProgTripCount_Visited, 
          Opps.CRVisitCount_Visited
INTO      DataAnalysis.dbo.TblData_OppHistory_MktgFastTrack
FROM      DataAnalysis.dbo.MSMK_Opportunities AS Opps
          LEFT JOIN  DBAmp_SF.dbo.Program__c AS Program_SF ON Opps.Program = Program_SF.Program_Acronym__c

/*************************************************************************************************************************************************************
MSMK_OppHistory

     Fields Calculated: Effective End
**************************************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MSMK_OppHistory', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_OppHistory

SELECT    OppHistory_SF.OpportunityId AS OppSFID, 
          OppHistory_SF.Id AS OppHistorySFID, 
          OppHistory_SF.StageName AS Stage, 
          OppHistory_SF.Amount AS ProposalValue, 
          OppHistory_SF.Probability, 
          OppHistory_SF.CloseDate, 
          OppHistory_SF.CreatedDate AS EffectiveStart, 
          Next_History_Grouped.EffectiveEnd
INTO      DataAnalysis.dbo.MSMK_OppHistory
FROM      DBAmp_SF.dbo.OpportunityHistory AS OppHistory_SF
          INNER JOIN DBAmp_SF.dbo.Opportunity AS Opp_SF ON OppHistory_SF.OpportunityId = Opp_SF.Id
          INNER JOIN DBAmp_SF.dbo.RecordType AS RT_Opp ON Opp_SF.RecordTypeId = RT_Opp.Id
          LEFT JOIN  (  SELECT    OppHistory_SF.Id, 
                                  MIN( Next_OppHistory.CreatedDate ) AS EffectiveEnd
                        FROM      DBAmp_SF.dbo.OpportunityHistory AS OppHistory_SF
                                  LEFT JOIN  DBAmp_SF.dbo.OpportunityHistory AS Next_OppHistory
                                             ON OppHistory_SF.OpportunityId = Next_OppHistory.OpportunityId
                                                AND ( OppHistory_SF.CreatedDate < Next_OppHistory.CreatedDate
                                                      OR ( OppHistory_SF.CreatedDate = Next_OppHistory.CreatedDate
                                                           AND OppHistory_SF.Id < Next_OppHistory.Id ) )
                        GROUP BY  OppHistory_SF.Id )
                     AS Next_History_Grouped ON OppHistory_SF.Id = Next_History_Grouped.Id
WHERE     RT_Opp.Name IN ('PT Marketing','RI Marketing','Consulting & Management')

/*************************************************************************************************************************************************************
MS_MemberHistory
**************************************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MS_MemberHistory', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MS_MemberHistory

SELECT    Contracts_Union_Integrated.*
INTO      DataAnalysis.dbo.MS_MemberHistory
FROM      (  SELECT    AccountCalc.AccountSFID AS PayerAccountSFID, 
                       AccountCalc.AccountCounterID AS PayerAccountCounterID, 
                       AccountCalc.AccountName AS PayerAccountName, 
                       AccountCalc.PrimaryCity AS PayerAccountCity, 
                       AccountCalc.PrimaryState AS PayerAccountState, 
                       AccountCalc.AccountSFID AS MemberAccountSFID, 
                       AccountCalc.AccountCounterID AS MemberAccountCounterID, 
                       AccountCalc.AccountName AS MemberAccountName, 
                       AccountCalc.PrimaryCity AS MemberAccountCity, 
                       AccountCalc.PrimaryState AS MemberAccountState, 
                       Contract_SF.ID AS ContractSFID, 
                       Contract_SF.Counter_ID__c AS ContractCounterID, 
                       CAST ( Contract_SF.Start__c AS DATETIME) AS StartDate, 
                       CAST ( Contract_SF.End__c AS DATETIME) AS EndDate, 
                       Contract_SF.Status__c AS [Status], 
                       'Payer' AS PayerStatus, 
                       Program_SF.ID AS ProgramSFID, 
                       Program_SF.Program_Acronym__c AS Program, 
                       Program_SF.Counter_ID__c AS ProgramCounterID, 
                       Contract_SF.Annual_Contract_Value__c AS AnnualCV, 
                       Contract_SF.Negotiated_Amount__c AS TotalCV, 
                       Membership_SF.ID AS MembershipSFID, 
                       Membership_SF.Counter_ID__c AS MembershipCounterID, 
                       Contract_SF.ID AS UniqueID, 
                       RepGrp.NewReportingGroup AS ReportingGroup, 
                       RepGrp.[Type] AS Business
             FROM      DBAmp_SF.dbo.Contract__c AS Contract_SF
                       INNER JOIN DataAnalysis.dbo.MS_Accounts AS AccountCalc ON Contract_SF.Payer_Account__c = AccountCalc.AccountSFID
                       INNER JOIN DBAmp_SF.dbo.Program__c AS Program_SF ON Contract_SF.Program__c = Program_SF.Id
                       LEFT JOIN  DBAmp_SF.dbo.Membership__c AS Membership_SF
                                  ON Contract_SF.Payer_Account__c = Membership_SF.Account_Name__c
                                     AND Contract_SF.Program__c = Membership_SF.Program__c
                       LEFT JOIN  DataAnalysis.dbo.ReportingHierarchyTable AS RepGrp ON Program_SF.Program_Acronym__c = RepGrp.Program
             WHERE     Contract_SF.IsDeleted = 0
                       
             UNION ALL
                       
             SELECT    PayerAccountCalc.AccountSFID AS PayerAccountSFID, 
                       PayerAccountCalc.AccountCounterID AS PayerAccountCounterID, 
                       PayerAccountCalc.AccountName AS PayerAccountName, 
                       PayerAccountCalc.PrimaryCity AS PayerAccountCity, 
                       PayerAccountCalc.PrimaryState AS PayerAccountState, 
                       MemberAccountCalc.AccountSFID AS MemberAccountSFID, 
                       MemberAccountCalc.AccountCounterID AS MemberAccountCounterID, 
                       MemberAccountCalc.AccountName AS MemberAccountName, 
                       MemberAccountCalc.PrimaryCity AS MemberAccountCity, 
                       MemberAccountCalc.PrimaryState AS MemberAccountState, 
                       Contract_SF.ID AS ContractSFID, 
                       Contract_SF.Counter_ID__c AS ContractCounterID, 
                       CAST ( Contract_SF.Start__c AS DATETIME) AS StartDate, 
                       CAST ( Contract_SF.End__c AS DATETIME) AS EndDate, 
                       Contract_SF.Status__c AS [Status], 
                       'Integrated' AS PayerStatus, 
                       Program_SF.ID AS ProgramSFID, 
                       Program_SF.Program_Acronym__c AS Program, 
                       Program_SF.Counter_ID__c AS ProgramCounterID, 
                       0 AS AnnualCV, 
                       0 AS TotalCV, 
                       Membership_SF.ID AS MembershipSFID, 
                       Membership_SF.Counter_ID__c AS MembershipCounterID, 
                       Integrated_SF.Id AS UniqueID, 
                       RepGrp.NewReportingGroup AS ReportingGroup, 
                       RepGrp.[Type] AS Business
             FROM      DBAmp_SF.dbo.Contract_Integrated_Accounts__c AS Integrated_SF
                       INNER JOIN DBAmp_SF.dbo.Contract__c AS Contract_SF ON Integrated_SF.Contract__c = Contract_SF.Id
                       INNER JOIN DataAnalysis.dbo.MS_Accounts AS PayerAccountCalc ON Contract_SF.Payer_Account__c = PayerAccountCalc.AccountSFID
                       INNER JOIN DataAnalysis.dbo.MS_Accounts AS MemberAccountCalc ON Integrated_SF.Account__c = MemberAccountCalc.AccountSFID
                       INNER JOIN DBAmp_SF.dbo.Program__c AS Program_SF ON Contract_SF.Program__c = Program_SF.Id
                       LEFT JOIN  DBAmp_SF.dbo.Membership__c AS Membership_SF
                                  ON Integrated_SF.Account__c = Membership_SF.Account_Name__c
                                     AND Contract_SF.Program__c = Membership_SF.Program__c
                       LEFT JOIN  DataAnalysis.dbo.ReportingHierarchyTable AS RepGrp ON Program_SF.Program_Acronym__c = RepGrp.Program
             WHERE     Integrated_SF.IsDeleted = 0
                       AND Contract_SF.IsDeleted = 0
                       AND MemberAccountCalc.AccountSFID != PayerAccountCalc.AccountSFID )
          AS Contracts_Union_Integrated

/*************************************************************************************************************************************************************
MSMK_Leads

     Written by Alex Koeberle 8/26/2016
**************************************************************************************************************************************************************/

IF OBJECT_ID('DataAnalysis.dbo.MSMK_Leads', 'U') IS NOT NULL
   DROP TABLE DataAnalysis.dbo.MSMK_Leads

SELECT    Leads.Id AS LeadSFID,
          Leads.Counter_ID__c AS LeadCounterID,
          Leads.Account__c AS AccountSFID,
          Leads.Lead_Date__c AS LeadDate,
          Leads.CreatedDate AS LeadCreatedDate,
          Leads.Reporting_Lead_Channel__c AS LeadReportingChannel,
          Leads.Lead_Origin__c AS LeadOrigin,
          Leads.Lead_Channel__c AS LeadEvent,
          Leads.Status__c AS LeadStatus,
          Leads.Contact_Title__c AS LeadContactTitle,
          RepGrp.Program AS LeadProgram,
          RepGrp.NewReportingGroup AS ReportingGroup,
          CASE WHEN Leads.Visited__c = 'Visited' THEN 1 ELSE 0 END AS LeadVisited,
          CASE WHEN OAB.OAB_Credit IS NOT NULL THEN 1 ELSE 0 END AS OABCredit, 
          ROW_NUMBER() OVER( 
            PARTITION BY Leads.Account__c, 
                         Leads.Primary_Program__c
            ORDER BY     Leads.CreatedDate ASC ) AS LeadOrder
INTO      DataAnalysis.dbo.MSMK_Leads  
FROM      DBAmp_SF.dbo.Warm_Lead__c AS Leads
          LEFT JOIN  DBAmp_SF.dbo.Program__c AS Program ON Program.ID = Leads.Primary_Program__c
          LEFT JOIN  DataAnalysis.dbo.ReportingHierarchyTable AS RepGrp ON RepGrp.Program = Program.Program_Acronym__c
          LEFT JOIN  (  SELECT     COUNT( CASE WHEN OABx.OAB_Unit_Credit__c IS NOT NULL 
                                                    AND OABx.OAB_Unit_Credit__c > '0' 
                                               THEN 1 ELSE 0 END ) 
                                               AS OAB_Credit,
                                    OABx.Warm_Lead__c
                          FROM      DBAmp_SF.dbo.OAB_Lead_Passer__c AS OABx
                          GROUP BY  OABx.Warm_Lead__c )     
                      AS OAB ON OAB.Warm_Lead__c = Leads.ID

/***********************************************************************************************************************************************

REMAINDER OF THIS IS JUST-BARELY-ALIVE GARBAGE. WHAT IS DEAD MAY NEVER DIE.

***********************************************************************************************************************************************/

/** CHECKED: Updating Visit to Eval Table (DataAnalysis.dbo.TblData_VisittoEval) **/
IF OBJECT_ID('DataAnalysis.dbo.TblData_VisittoEval', 'U') IS NOT NULL
     DROP TABLE DataAnalysis.dbo.TblData_VisittoEval

SELECT 
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID AS OppSFID, 
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.COUNTER_ID__C AS OppCounterID, 
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.NAME AS Name, 
     SalesforceReplication.dbo.SFREP_ACCOUNT.COUNTER_ID__C AS InstCounterID, 
     SalesforceReplication.dbo.SFREP_ACCOUNT.NAME AS InstName, 
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.STAGENAME AS StageName, 
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.CLOSEDATE AS CloseDate, 
     CASE WHEN SalesforceReplication.dbo.SFREP_USER.Name IS NULL THEN User_Owner.Name ELSE SalesforceReplication.dbo.SFREP_USER.Name END AS Marketer, 
     CASE WHEN SalesforceReplication.dbo.SFREP_USER.EMPLOYEE_CONTACT_RECORDID__C IS NULL 
          THEN User_Owner.EMPLOYEE_CONTACT_RECORDID__C 
          ELSE SalesforceReplication.dbo.SFREP_USER.EMPLOYEE_CONTACT_RECORDID__C END AS MarketerCounterID, 
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROGRAM_ACRONYM__C AS Program, 
     DataAnalysis.dbo.TblLookup_ReportingGroup.ReportingGroup As ReportingGroup, 
     DataAnalysis.dbo.TblLookup_ReportingGroup.BusinessLine AS Business, 
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.INITIAL_VISIT_DATE__C AS IntialVisitDate, 
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.EVALUATED__C AS Evaluated, 
     1 AS OppCount
INTO  DataAnalysis.dbo.TblData_VisittoEval
FROM ((((((SalesforceReplication.dbo.SFREP_OPPORTUNITY 
     LEFT JOIN SalesforceReplication.dbo.SFREP_2_OPPORTUNITY ON SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID = SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.ID) 
     LEFT JOIN DataAnalysis.dbo.TblData_VisitsthatCount_SF ON SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID = DataAnalysis.dbo.TblData_VisitsthatCount_SF.OppSFID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_USER ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.MARKETER__C = SalesforceReplication.dbo.SFREP_USER.ID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_PROGRAM__C ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROGRAM__C = SalesforceReplication.dbo.SFREP_PROGRAM__C.ID) 
     LEFT JOIN DataAnalysis.dbo.TblLookup_ReportingGroup ON SalesforceReplication.dbo.SFREP_PROGRAM__C.PROGRAM_ACRONYM__C = DataAnalysis.dbo.TblLookup_ReportingGroup.Program) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_ACCOUNT ON SalesforceReplication.dbo.SFREP_OPPORTUNITY.ACCOUNTID = SalesforceReplication.dbo.SFREP_ACCOUNT.ID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_USER AS User_Owner ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.OWNERID = User_Owner.ID
WHERE (SalesforceReplication.dbo.SFREP_OPPORTUNITY.INITIAL_VISIT_DATE__C >='20120201' 
     AND (SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID = '012C0000000BkFWIA0' OR SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID = '012C0000000BkFZIA0') 
     AND DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate >= '20120101' 
     AND SalesforceReplication.dbo.SFREP_OPPORTUNITY.CREATEDDATE >= '20120211')
     
/** CHECKED: Update Opps With Cycle Time Table **/
IF OBJECT_ID('DataAnalysis.dbo.TblData_OppsWithCycleTime', 'U') IS NOT NULL
     DROP TABLE DataAnalysis.dbo.TblData_OppsWithCycleTime
SELECT 
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.COUNTER_ID__C AS OppCounterID,
     SalesforceReplication.dbo.SFREP_ACCOUNT.COUNTER_ID__C AS InstCounterID,
     SalesforceReplication.dbo.SFREP_ACCOUNT.NAME AS InstName,
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.STAGENAME AS StageName,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.CLOSEDATE AS CloseDate,
     CASE WHEN SalesforceReplication.dbo.SFREP_USER.NAME IS NULL THEN User_Owner.NAME      ELSE SalesforceReplication.dbo.SFREP_USER.NAME END AS Marketer, 
     CASE WHEN SalesforceReplication.dbo.SFREP_User.EMPLOYEE_CONTACT_RECORDID__C IS NULL 
          THEN User_Owner.EMPLOYEE_CONTACT_RECORDID__C 
          ELSE SalesforceReplication.dbo.SFREP_User.EMPLOYEE_CONTACT_RECORDID__C END AS MarketerCounterID,
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROGRAM_ACRONYM__C AS Program, 
     DataAnalysis.dbo.TblLookup_ReportingGroup.ReportingGroup AS ReportingGroup,
     DataAnalysis.dbo.TblLookup_ReportingGroup.BusinessLine AS Business,
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.NBB_UNITS__C AS Units, 
     NULL AS VisitsWithinWindowAll, 
     CAST(NULL AS DATETIME) AS FirstVisitDate,
     CAST(NULL AS DATETIME)  AS LastVisitDate,
     NULL AS CycleTime_FirstVisit, 
     NULL AS CycleTime_LastVisit
INTO DataAnalysis.dbo.TblData_OppsWithCycleTime
FROM (((((SalesforceReplication.dbo.SFREP_OPPORTUNITY 
     LEFT JOIN SalesforceReplication.dbo.SFREP_2_OPPORTUNITY ON SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID = SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.ID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_USER ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.MARKETER__C = SalesforceReplication.dbo.SFREP_USER.ID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_PROGRAM__C ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROGRAM__C = SalesforceReplication.dbo.SFREP_PROGRAM__C.ID) 
     LEFT JOIN DataAnalysis.dbo.TblLookup_ReportingGroup ON SalesforceReplication.dbo.SFREP_PROGRAM__C.PROGRAM_ACRONYM__C = DataAnalysis.dbo.TblLookup_ReportingGroup.Program) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_ACCOUNT ON SalesforceReplication.dbo.SFREP_OPPORTUNITY.ACCOUNTID = SalesforceReplication.dbo.SFREP_ACCOUNT.ID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_USER AS User_Owner ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.OWNERID = User_Owner.ID
WHERE SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.STAGENAME = 'Closed Won' 
     AND SalesforceReplication.dbo.SFREP_OPPORTUNITY.CLOSEDATE >= '20080101'
     AND SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.NBB_UNITS__C > 0
     AND (SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID = '012C0000000BkFWIA0' OR SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID = '012C0000000BkFZIA0')

SELECT /** TEMP - CHECKED: Determining the Number of Visits within the Window Given the Reporting Group (DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow) **/
     DataAnalysis.dbo.TblData_OppsWithCycleTime.ID, 
     SUM(DataAnalysis.dbo.TblData_VisitsthatCount_SF.CRVisitCount_RepGrp) AS VisitsWithinWindow_RG
INTO DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow
FROM (DataAnalysis.dbo.TblData_OppswithCycleTime 
     LEFT JOIN DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow ON DataAnalysis.dbo.TblData_OppswithCycleTime.Business = DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.BusinessLine) 
     LEFT JOIN DataAnalysis.dbo.TblData_VisitsthatCount_SF ON (DataAnalysis.dbo.TblData_OppswithCycleTime.ReportingGroup = DataAnalysis.dbo.TblData_VisitsthatCount_SF.ReportingGroup AND DataAnalysis.dbo.TblData_OppswithCycleTime.InstCounterID = DataAnalysis.dbo.TblData_VisitsthatCount_SF.InstCounterID)
WHERE (((DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate)<=DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate 
     AND DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate>=(DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate - (CASE WHEN DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.EvaluationWindow IS NULL THEN 0 ELSE DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.EvaluationWindow END))))
GROUP BY DataAnalysis.dbo.TblData_OppsWithCycleTime.ID

SELECT /** TEMP - CHECKED: Determining First Visit Date for Opps with Cycle Time (DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstVisitDate) **/
     DataAnalysis.dbo.TblData_OppswithCycleTime.ID, 
     MIN(DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate) AS FirstVisitDate
INTO DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstVisitDate
FROM ((DataAnalysis.dbo.TblData_OppswithCycleTime 
     LEFT JOIN DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow ON DataAnalysis.dbo.TblData_OppswithCycleTime.Business = DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.BusinessLine) 
     LEFT JOIN DataAnalysis.dbo.TblData_VisitsthatCount_SF ON (DataAnalysis.dbo.TblData_OppswithCycleTime.ReportingGroup = DataAnalysis.dbo.TblData_VisitsthatCount_SF.ReportingGroup) AND (DataAnalysis.dbo.TblData_OppswithCycleTime.InstCounterID = DataAnalysis.dbo.TblData_VisitsthatCount_SF.InstCounterID)) 
     LEFT JOIN DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow ON DataAnalysis.dbo.TblData_OppswithCycleTime.ID = DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.ID
WHERE (((DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate)<=DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate 
     AND (DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate)>=DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate - (CASE WHEN DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.EvaluationWindow IS NULL THEN 0 ELSE DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.EvaluationWindow END)) 
     AND ((DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.VisitsWithinWindow_RG)>0))
GROUP BY DataAnalysis.dbo.TblData_OppswithCycleTime.ID

SELECT /** TEMP - CHECKED: Determining Last Visit Date for Opps with Cycle Time (DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_LastVisitDate)**/
     DataAnalysis.dbo.TblData_OppswithCycleTime.ID, 
     MAX(DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate) AS LastVisitDate
INTO DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_LastVisitDate
FROM ((DataAnalysis.dbo.TblData_OppswithCycleTime 
     LEFT JOIN DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow ON DataAnalysis.dbo.TblData_OppswithCycleTime.Business = DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.BusinessLine) 
     LEFT JOIN DataAnalysis.dbo.TblData_VisitsthatCount_SF ON (DataAnalysis.dbo.TblData_OppswithCycleTime.ReportingGroup = DataAnalysis.dbo.TblData_VisitsthatCount_SF.ReportingGroup) AND (DataAnalysis.dbo.TblData_OppswithCycleTime.InstCounterID = DataAnalysis.dbo.TblData_VisitsthatCount_SF.InstCounterID)) 
     LEFT JOIN DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow ON DataAnalysis.dbo.TblData_OppswithCycleTime.ID = DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.ID
WHERE (((DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate)<=DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate 
     AND (DataAnalysis.dbo.TblData_VisitsthatCount_SF.ActivityDate)>=DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate - (CASE WHEN DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.EvaluationWindow IS NULL THEN 0 ELSE DataAnalysis.dbo.TblSettings_CycleTime_EvalWindow.EvaluationWindow END)) 
     AND ((DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.VisitsWithinWindow_RG)>0))
GROUP BY DataAnalysis.dbo.TblData_OppsWithCycleTime.ID

SELECT /** TEMP - CHECKED: Aligning the Visit Data **/
     DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.ID, 
     DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.VisitsWithinWindow_RG, 
     DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstVisitDate.FirstVisitDate, 
     DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_LastVisitDate.LastVisitDate
INTO  DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate
FROM (DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow 
     LEFT JOIN DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstVisitDate ON DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.ID = DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstVisitDate.ID) 
     LEFT JOIN DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_LastVisitDate ON DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow.ID = DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_LastVisitDate.ID

UPDATE /** COMPLETE: Updating the Opps with Cycle Time Table to Include Visit Data **/
     DataAnalysis.dbo.TblData_OppswithCycleTime
SET 
     DataAnalysis.dbo.TblData_OppswithCycleTime.VisitsWithinWindowAll = DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate.VisitsWithinWindow_RG, 
     DataAnalysis.dbo.TblData_OppswithCycleTime.FirstVisitDate = DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate.FirstVisitDate, 
     DataAnalysis.dbo.TblData_OppswithCycleTime.LastVisitDate = DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate.LastVisitDate, 
     DataAnalysis.dbo.TblData_OppswithCycleTime.CycleTime_FirstVisit = DATEDIFF(day, DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate.FirstVisitDate, DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate), 
     DataAnalysis.dbo.TblData_OppswithCycleTime.CycleTime_LastVisit = DATEDIFF(day, DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate.LastVisitDate, DataAnalysis.dbo.TblData_OppswithCycleTime.CloseDate)
FROM DataAnalysis.dbo.TblData_OppswithCycleTime 
     LEFT JOIN DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate ON DataAnalysis.dbo.TblData_OppswithCycleTime.ID = DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate.ID 
WHERE (((DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate.ID) IS NOT NULL))

IF OBJECT_ID('DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow', 'U') IS NOT NULL
     DROP TABLE DataAnalysis.dbo.Tmp_OppswithCT_RepGrp_VisitsinWindow
IF OBJECT_ID('DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstVisitDate', 'U') IS NOT NULL
     DROP TABLE DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstVisitDate
IF OBJECT_ID('DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_LastVisitDate', 'U') IS NOT NULL     
     DROP TABLE DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_LastVisitDate
IF OBJECT_ID('DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate', 'U') IS NOT NULL     
     DROP TABLE DataAnalysis.dbo.TblTmp_OppswithCT_RepGrp_FirstLastVisitDate

/** COMPLETE: Update Large Contracts Table (DataAnalysis.dbo.TblData_OppHistory_LgContracts) **/
IF OBJECT_ID('DataAnalysis.dbo.TblData_OppHistory_LgContracts', 'U') IS NOT NULL     
     DROP TABLE DataAnalysis.dbo.TblData_OppHistory_LgContracts
SELECT 
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.ID,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID AS OpportunityID,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CREATEDBYID AS CreatedByID,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CREATEDDATE AS CreatedByDate,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.STAGENAME AS StageName,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.AMOUNT AS Amount,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.EXPECTEDREVENUE AS ExpectedRevenue,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CLOSEDATE AS CloseDate,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.PROBABILITY AS Probability,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.FORECASTCATEGORY AS ForecastCategory,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.SYSTEMMODSTAMP AS SystemModStamp,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.ISDELETED AS IsDeleted,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CurrencyIsoCode,
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.COUNTER_ID__C AS OppCounterID, 
     SalesforceReplication.dbo.SFREP_ACCOUNT.COUNTER_ID__C AS InstCounterID, 
     SalesforceReplication.dbo.SFREP_ACCOUNT.NAME AS Inst, 
     SalesforceReplication.dbo.SFREP_Program__c.PROGRAM_ACRONYM__C AS Program, 
     DataAnalysis.dbo.TblLookup_ReportingGroup.ReportingGroup AS ReportingGroup,
     DataAnalysis.dbo.TblLookup_ReportingGroup.BusinessLine AS Business, 
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.STAGENAME AS CurrentStage, 
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROBABILITY AS CurrentProbability,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.AMOUNT AS CurrentAmount,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.CLOSEDATE AS CurrentCloseDate,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.INITIAL_VISIT_DATE__C AS InitialVisitDate,
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID      AS RecordTypeID
INTO DataAnalysis.dbo.TblData_OppHistory_LgContracts
FROM 
     ((((((SELECT SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID AS OpportunityID
               FROM (((SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY
                    LEFT JOIN SalesforceReplication.dbo.SFREP_2_OPPORTUNITY ON SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID = SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.ID)
                    LEFT JOIN SalesforceReplication.dbo.SFREP_OPPORTUNITY ON SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID = SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID)
                    LEFT JOIn SalesforceReplication.dbo.SFREP_Program__c ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROGRAM__C = SalesforceReplication.dbo.SFREP_Program__c.ID)
               WHERE (SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.AMOUNT >= 400000
                    AND (SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID = '012C0000000BkFWIA0' OR SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID = '012C0000000BkFZIA0')
                    AND SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CREATEDDATE >= '20110801'
                    AND SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CLOSEDATE >= '20110801')
               GROUP BY SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID) AS QryCalc_OppHistory_LgContracts_OppIDs
     LEFT JOIN SalesforceReplication.dbo.SFREP_OPPORTUNITY ON QryCalc_OppHistory_LgContracts_OppIDs.OpportunityID = SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID)
     LEFT JOIN SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY ON QryCalc_OppHistory_LgContracts_OppIDs.OpportunityID = SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID)
     LEFT JOIN SalesforceReplication.dbo.SFREP_2_OPPORTUNITY ON QryCalc_OppHistory_LgContracts_OppIDs.OpportunityID = SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.ID)
     LEFT JOIN SalesforceReplication.dbo.SFREP_ACCOUNT ON SalesforceReplication.dbo.SFREP_OPPORTUNITY.ACCOUNTID = SalesforceReplication.dbo.SFREP_ACCOUNT.ID)
     LEFT JOIN SalesforceReplication.dbo.SFREP_Program__c ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROGRAM__C = SalesforceReplication.dbo.SFREP_Program__c.ID)
     LEFT JOIN DataAnalysis.dbo.TblLookup_ReportingGroup ON SalesforceReplication.dbo.SFREP_Program__c.PROGRAM_ACRONYM__C = DataAnalysis.dbo.TblLookup_ReportingGroup.Program

/** COMPLETE: Update Open Marketing Table (DataAnalysis.dbo.TblData_OppHistory_OpenMktg) **/
IF OBJECT_ID('DataAnalysis.dbo.TblData_OppHistory_OpenMktg', 'U') IS NOT NULL     
     DROP TABLE DataAnalysis.dbo.TblData_OppHistory_OpenMktg
SELECT 
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.ID,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID AS OpportunityID,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CREATEDBYID AS CreatedByID,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CREATEDDATE AS CreatedDate,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.STAGENAME AS StageName,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.AMOUNT AS Amount, 
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.EXPECTEDREVENUE AS ExpectedRevenue,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CLOSEDATE AS CloseDate,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.PROBABILITY AS Probability,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.FORECASTCATEGORY AS ForecastCategory,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.SYSTEMMODSTAMP AS SystemModStamp,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.ISDELETED AS IsDeleted,
     SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.CurrencyIsoCode AS CurrencyIsoCode,
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.COUNTER_ID__C AS OppCounterID,  
     SalesforceReplication.dbo.SFREP_ACCOUNT.COUNTER_ID__C AS InstCounterID, 
     SalesforceReplication.dbo.SFREP_ACCOUNT.NAME AS Inst, 
     SalesforceReplication.dbo.SFREP_Program__c.PROGRAM_ACRONYM__C AS Program, 
     DataAnalysis.dbo.TblLookup_ReportingGroup.ReportingGroup AS ReportingGroup,
     DataAnalysis.dbo.TblLookup_ReportingGroup.BusinessLine AS BusinessLine, 
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.STAGENAME AS CurrentStage, 
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROBABILITY AS CurrentProbability,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.AMOUNT AS CurrentAmount,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.CLOSEDATE AS CurrentCloseDate,
     SalesforceReplication.dbo.SFREP_OPPORTUNITY.INITIAL_VISIT_DATE__C AS InitialVisitDate,
     SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID      AS RecordTypeID
INTO DataAnalysis.dbo.TblData_OppHistory_OpenMktg
FROM (((((SalesforceReplication.dbo.SFREP_2_OPPORTUNITY 
     LEFT JOIN SalesforceReplication.dbo.SFREP_PROGRAM__C ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.PROGRAM__C = SalesforceReplication.dbo.SFREP_PROGRAM__C.ID) 
     LEFT JOIN DataAnalysis.dbo.TblLookup_ReportingGroup ON SalesforceReplication.dbo.SFREP_PROGRAM__C.PROGRAM_ACRONYM__C = TblLookup_ReportingGroup.Program) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_OPPORTUNITY ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.ID = SalesforceReplication.dbo.SFREP_OPPORTUNITY.ID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_ACCOUNT ON SalesforceReplication.dbo.SFREP_OPPORTUNITY.ACCOUNTID = SalesforceReplication.dbo.SFREP_ACCOUNT.ID) 
     LEFT JOIN SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY ON SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.ID = SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.OPPORTUNITYID)
WHERE (SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.STAGENAME <>'Closed Lost' 
     AND SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.STAGENAME <>'Closed Won'
     AND (SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID ='012C0000000BkFWIA0' OR SalesforceReplication.dbo.SFREP_2_OPPORTUNITY.RECORDTYPEID = '012C0000000BkFZIA0')
     AND SalesforceReplication.dbo.SFREP_OPPORTUNITYHISTORY.ID IS NOT NULL)

END
