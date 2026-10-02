USE [FraudHub];


GO
/* =========================================================================
   1. SCHEMAS
   ========================================================================= */
IF NOT EXISTS (SELECT *
               FROM   sys.schemas
               WHERE  name = 'core')
    EXECUTE ('CREATE SCHEMA [core]');


GO
IF NOT EXISTS (SELECT *
               FROM   sys.schemas
               WHERE  name = 'stg')
    EXECUTE ('CREATE SCHEMA [stg]');


GO
/* =========================================================================
   2. CORE SCHEMA TABLES & CONSTRAINTS
   ========================================================================= */
-- Players reference table
CREATE TABLE [core].[players] (
    [player_id] VARCHAR (10) NOT NULL,
    [full_name] VARCHAR (100) NULL,
    [email] VARCHAR (150) NULL,
    [country] VARCHAR (5) NULL,
    [registered_at] DATETIME2 (7) NULL,
    CONSTRAINT [PK_core_players] PRIMARY KEY CLUSTERED ([player_id] ASC)
);


GO
-- Rule weights configuration table
CREATE TABLE [core].[rule_weights] (
    [rule_code] VARCHAR (30) NOT NULL,
    [points] INT NOT NULL,
    CONSTRAINT [PK_core_rule_weights] PRIMARY KEY CLUSTERED ([rule_code] ASC)
);


GO
-- Default rule configuration seed
INSERT  INTO [core].[rule_weights] ([rule_code], [points])
VALUES                            ('MULTI_ACCOUNT_DEVICE', 30),
('SHARED_IP', 20),
('SHARED_CARD', 35),
('PAYMENT_VELOCITY', 25),
('BONUS_ABUSE', 40),
('FAILED_PAYMENTS', 20);


GO
-- Risk flags staging table
CREATE TABLE [core].[risk_flags] (
    [player_id] VARCHAR (10) NULL,
    [rule_code] VARCHAR (30) NULL,
    [detail] VARCHAR (200) NULL
);


GO
-- Player bonus claims ledger
CREATE TABLE [core].[bonuses] (
    [bonus_id] VARCHAR (10) NOT NULL,
    [player_id] VARCHAR (10) NULL,
    [claimed_at] DATETIME2 (7) NULL,
    [bonus_amount] DECIMAL (10, 2) NULL,
    [wagered_amount] DECIMAL (12, 2) NULL,
    [hours_to_first_withdrawal] DECIMAL (9, 1) NULL,
    CONSTRAINT [PK_core_bonuses] PRIMARY KEY CLUSTERED ([bonus_id] ASC),
    CONSTRAINT [FK_bonuses_players] FOREIGN KEY ([player_id]) REFERENCES [core].[players] ([player_id])
);


GO
-- Player device/network tracking
CREATE TABLE [core].[devices] (
    [player_id] VARCHAR (10) NULL,
    [device_id] VARCHAR (10) NULL,
    [ip_address] VARCHAR (20) NULL,
    CONSTRAINT [FK_devices_players] FOREIGN KEY ([player_id]) REFERENCES [core].[players] ([player_id])
);


GO
-- Payment transactions ledger
CREATE TABLE [core].[payments] (
    [txn_id] VARCHAR (12) NOT NULL,
    [player_id] VARCHAR (10) NULL,
    [txn_time] DATETIME2 (7) NULL,
    [txn_type] VARCHAR (15) NULL,
    [method] VARCHAR (20) NULL,
    [amount] DECIMAL (12, 2) NULL,
    [status] VARCHAR (15) NULL,
    [card_hash] VARCHAR (20) NULL,
    CONSTRAINT [PK_core_payments] PRIMARY KEY CLUSTERED ([txn_id] ASC),
    CONSTRAINT [FK_payments_players] FOREIGN KEY ([player_id]) REFERENCES [core].[players] ([player_id])
);


GO
/* =========================================================================
   3. APPLICATION TABLES (dbo)
   ========================================================================= */
-- Analysts directory
CREATE TABLE [dbo].[Analysts] (
    [AnalystEmail] NVARCHAR (100) NOT NULL,
    [AnalystName] NVARCHAR (100) NOT NULL,
    [IsActive] BIT CONSTRAINT [DF_Analysts_IsActive] DEFAULT (1) NOT NULL,
    CONSTRAINT [PK_dbo_Analysts] PRIMARY KEY CLUSTERED ([AnalystEmail] ASC)
);


GO
-- Fraud case queue
CREATE TABLE [dbo].[FraudCases] (
    [CaseID] INT IDENTITY (1, 1) NOT NULL,
    [PlayerID] VARCHAR (50) NOT NULL,
    [RiskBand] VARCHAR (20) NULL,
    [RiskScore] INT NULL,
    [Rules] VARCHAR (MAX) NULL,
    [Evidence] VARCHAR (MAX) NULL,
    [Status] VARCHAR (50) CONSTRAINT [DF_FraudCases_Status] DEFAULT ('New') NULL,
    [Disposition] VARCHAR (50) NULL,
    [AssignedToEmail] VARCHAR (100) NULL,
    [AssignedToName] VARCHAR (100) NULL,
    [AnalystNotes] VARCHAR (MAX) NULL,
    [Source] VARCHAR (50) CONSTRAINT [DF_FraudCases_Source] DEFAULT ('Daily exceptions') NULL,
    [CreatedOn] DATETIME CONSTRAINT [DF_FraudCases_CreatedOn] DEFAULT (GETDATE()) NULL,
    [ClosedOn] DATETIME NULL,
    [ApprovalSentOn] DATETIME2 (7) NULL,
    [ApprovedBy] NVARCHAR (100) NULL,
    [ApprovedOn] DATETIME2 (7) NULL,
    [ApprovalComment] NVARCHAR (500) NULL,
    [SLADue] DATETIME2 (7) NULL,
    [TurnaroundHours] INT NULL,
    CONSTRAINT [PK_dbo_FraudCases] PRIMARY KEY CLUSTERED ([CaseID] ASC)
);


GO
/* =========================================================================
   4. STAGING TABLES (stg)
   ========================================================================= */
CREATE TABLE [stg].[players] (
    [player_id] VARCHAR (10) NULL,
    [full_name] VARCHAR (100) NULL,
    [email] VARCHAR (150) NULL,
    [country] VARCHAR (5) NULL,
    [registered_at] VARCHAR (30) NULL
);


GO
CREATE TABLE [stg].[bonuses] (
    [bonus_id] VARCHAR (10) NULL,
    [player_id] VARCHAR (10) NULL,
    [claimed_at] VARCHAR (30) NULL,
    [bonus_amount] VARCHAR (20) NULL,
    [wagered_amount] VARCHAR (20) NULL,
    [hours_to_first_withdrawal] VARCHAR (20) NULL
);


GO
CREATE TABLE [stg].[devices] (
    [player_id] VARCHAR (10) NULL,
    [device_id] VARCHAR (10) NULL,
    [ip_address] VARCHAR (20) NULL
);


GO
CREATE TABLE [stg].[payments] (
    [txn_id] VARCHAR (12) NULL,
    [player_id] VARCHAR (10) NULL,
    [txn_time] VARCHAR (30) NULL,
    [txn_type] VARCHAR (15) NULL,
    [method] VARCHAR (20) NULL,
    [amount] VARCHAR (20) NULL,
    [status] VARCHAR (15) NULL,
    [card_hash] VARCHAR (20) NULL
);


GO
/* =========================================================================
   5. VIEWS
   ========================================================================= */
CREATE OR ALTER VIEW [core].[vw_daily_exceptions]
AS
WITH     f
AS       (SELECT   player_id,
                   rule_code,
                   MIN(detail) AS detail
          FROM     core.risk_flags
          GROUP BY player_id, rule_code)
SELECT   f.player_id,
         p.full_name,
         p.country,
         SUM(w.points) AS risk_score,
         CASE WHEN SUM(w.points) >= 60 THEN 'High' WHEN SUM(w.points) >= 40 THEN 'Medium' ELSE 'Low' END AS risk_band,
         STRING_AGG(f.rule_code, '; ') WITHIN GROUP (ORDER BY f.rule_code) AS rules,
         STRING_AGG(f.detail, ' | ') AS details
FROM     f
         INNER JOIN
         core.rule_weights AS w
         ON w.rule_code = f.rule_code
         INNER JOIN
         core.players AS p
         ON p.player_id = f.player_id
GROUP BY f.player_id, p.full_name, p.country;


GO
CREATE OR ALTER VIEW [dbo].[vw_CasesJustCreated]
AS
SELECT CaseID,
       PlayerID,
       RiskBand,
       RiskScore,
       Rules,
       SLADue,
       AssignedToName
FROM   dbo.FraudCases
WHERE  Source = 'Auto'
       AND CreatedOn >= DATEADD(MINUTE, -30, SYSDATETIME());


GO
CREATE OR ALTER VIEW [dbo].[vw_SLA_Watch]
AS
SELECT   TOP (25) CaseID,
                  PlayerID,
                  RiskBand,
                  Status,
                  AssignedToName,
                  SLADue,
                  DATEDIFF(HOUR, SYSDATETIME(), SLADue) AS HoursLeft,
                  COUNT(*) OVER () AS TotalBreaching
FROM     dbo.FraudCases
WHERE    Status <> 'Closed'
         AND SLADue <= DATEADD(HOUR, 4, SYSDATETIME())
ORDER BY SLADue;


GO
CREATE OR ALTER VIEW [dbo].[vw_FraudHub_UserKPIs]
AS
SELECT   AssignedToEmail,
         COUNT(CASE WHEN Status <> 'Closed' THEN 1 END) AS MyOpenQueue
FROM     dbo.FraudCases
GROUP BY AssignedToEmail;


GO
CREATE OR ALTER VIEW [dbo].[vw_FraudHub_KPIs]
AS
SELECT COUNT(CASE WHEN Status <> 'Closed' THEN 1 END) AS TotalActiveCases,
       COUNT(CASE WHEN RiskBand = 'High'
                       AND Status <> 'Closed' THEN 1 END) AS HighRiskPriority,
       COUNT(CASE WHEN AssignedToEmail = USER_NAME()
                       AND Status <> 'Closed' THEN 1 END) AS MyActiveCases,
       COUNT(CASE WHEN Status <> 'Closed'
                       AND SLADue < GETDATE() THEN 1 END) AS SLABreachWarning
FROM   dbo.FraudCases;


GO
/* =========================================================================
   6. STORED PROCEDURES
   ========================================================================= */
CREATE OR ALTER PROCEDURE [core].[usp_refresh_risk_flags]
@device_min INT=3, @ip_min INT=3, @card_min INT=3, @vel_min INT=5, @wager_max DECIMAL (9, 2)=5, @wd_hours_max DECIMAL (9, 2)=24
AS
BEGIN
    SET NOCOUNT ON;
    TRUNCATE TABLE core.risk_flags;
    -- Multi-accounting: shared device
    WITH s
    AS   (SELECT   device_id,
                   COUNT(DISTINCT player_id) AS n
          FROM     core.devices
          GROUP BY device_id
          HAVING   COUNT(DISTINCT player_id) >= @device_min)
    INSERT INTO core.risk_flags (player_id, rule_code, detail)
    SELECT DISTINCT d.player_id,
                    'MULTI_ACCOUNT_DEVICE',
                    CONCAT('device_id=', d.device_id, ' shared by ', s.n, ' players')
    FROM   core.devices AS d
           INNER JOIN
           s
           ON s.device_id = d.device_id;
    -- Multi-accounting: shared IP
    WITH s
    AS   (SELECT   ip_address,
                   COUNT(DISTINCT player_id) AS n
          FROM     core.devices
          GROUP BY ip_address
          HAVING   COUNT(DISTINCT player_id) >= @ip_min)
    INSERT INTO core.risk_flags (player_id, rule_code, detail)
    SELECT DISTINCT d.player_id,
                    'SHARED_IP',
                    CONCAT('ip_address=', d.ip_address, ' shared by ', s.n, ' players')
    FROM   core.devices AS d
           INNER JOIN
           s
           ON s.ip_address = d.ip_address;
    -- Shared payment card
    WITH s
    AS   (SELECT   card_hash,
                   COUNT(DISTINCT player_id) AS n
          FROM     core.payments
          WHERE    card_hash IS NOT NULL
          GROUP BY card_hash
          HAVING   COUNT(DISTINCT player_id) >= @card_min)
    INSERT INTO core.risk_flags (player_id, rule_code, detail)
    SELECT DISTINCT p.player_id,
                    'SHARED_CARD',
                    CONCAT('card ', p.card_hash, ' used by ', s.n, ' players')
    FROM   core.payments AS p
           INNER JOIN
           s
           ON s.card_hash = p.card_hash;
    -- Deposit velocity: deposits within a 1-hour rolling window
    WITH v
    AS   (SELECT d.player_id,
                 (SELECT COUNT(*)
                  FROM   core.payments AS d2
                  WHERE  d2.player_id = d.player_id
                         AND d2.txn_type = 'deposit'
                         AND d2.txn_time >= d.txn_time
                         AND d2.txn_time < DATEADD(HOUR, 1, d.txn_time)) AS n
          FROM   core.payments AS d
          WHERE  d.txn_type = 'deposit')
    INSERT INTO core.risk_flags (player_id, rule_code, detail)
    SELECT   player_id,
             'PAYMENT_VELOCITY',
             CONCAT(MAX(n), ' deposits within 1 hour')
    FROM     v
    GROUP BY player_id
    HAVING   MAX(n) >= @vel_min;
    -- Bonus abuse: low wagering ratio with rapid withdrawal
    INSERT INTO core.risk_flags (player_id, rule_code, detail)
    SELECT DISTINCT player_id,
                    'BONUS_ABUSE',
                    CONCAT('wagered ', CAST (wagered_amount / bonus_amount AS DECIMAL (6, 1)), 'x, withdrew after ', hours_to_first_withdrawal, 'h')
    FROM   core.bonuses
    WHERE  wagered_amount / bonus_amount < @wager_max
           AND hours_to_first_withdrawal < @wd_hours_max;
END


GO
CREATE OR ALTER PROCEDURE [dbo].[usp_CreateFraudCases]
@cooldown_days INT=30, @sla_high INT=24, @sla_medium INT=72, @sla_low INT=120, @fail_min INT=4
AS
BEGIN
    SET NOCOUNT ON;
    EXECUTE core.usp_refresh_risk_flags ;
    IF NOT EXISTS (SELECT 1
                   FROM   core.rule_weights
                   WHERE  rule_code = 'FAILED_PAYMENTS')
        INSERT  INTO core.rule_weights (rule_code, points)
        VALUES                        ('FAILED_PAYMENTS', 20);
    INSERT INTO core.risk_flags (player_id, rule_code, detail)
    SELECT   p.player_id,
             'FAILED_PAYMENTS',
             CONCAT(COUNT(*), ' failed payments')
    FROM     core.payments AS p
    WHERE    p.status = 'failed'
    GROUP BY p.player_id
    HAVING   COUNT(*) >= @fail_min
             AND NOT EXISTS (SELECT 1
                             FROM   core.risk_flags AS f
                             WHERE  f.player_id = p.player_id);
    DECLARE @n AS INT = (SELECT COUNT(*)
                         FROM   dbo.Analysts
                         WHERE  IsActive = 1);
    WITH an
    AS   (SELECT AnalystEmail,
                 AnalystName,
                 ROW_NUMBER() OVER (ORDER BY AnalystEmail) - 1 AS slot
          FROM   dbo.Analysts
          WHERE  IsActive = 1),
         newc
    AS   (SELECT v.player_id,
                 v.risk_band,
                 v.risk_score,
                 v.rules,
                 v.details,
                 ROW_NUMBER() OVER (ORDER BY v.risk_score DESC, v.player_id) - 1 AS rn
          FROM   core.vw_daily_exceptions AS v
          WHERE  NOT EXISTS (SELECT 1
                             FROM   dbo.FraudCases AS c
                             WHERE  c.PlayerID = v.player_id
                                    AND (c.Status <> 'Closed'
                                         OR c.ClosedOn IS NULL
                                         OR c.ClosedOn >= DATEADD(DAY, -@cooldown_days, SYSDATETIME()))))
    INSERT INTO dbo.FraudCases (PlayerID, RiskBand, RiskScore, Rules, Evidence, Status, Source, CreatedOn, SLADue, AssignedToEmail, AssignedToName)
    SELECT n.player_id,
           n.risk_band,
           n.risk_score,
           n.rules,
           n.details,
           CASE WHEN an.AnalystEmail IS NULL THEN 'New' ELSE 'Assigned' END,
           'Auto',
           SYSDATETIME(),
           DATEADD(HOUR, CASE n.risk_band WHEN 'High' THEN @sla_high WHEN 'Medium' THEN @sla_medium ELSE @sla_low END, SYSDATETIME()),
           an.AnalystEmail,
           an.AnalystName
    FROM   newc AS n
           LEFT OUTER JOIN
           an
           ON an.slot = n.rn % NULLIF (@n, 0);
END


GO
CREATE OR ALTER PROCEDURE [dbo].[usp_MarkApprovalSent]
@CaseID INT
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.FraudCases
    SET    ApprovalSentOn = SYSDATETIME()
    WHERE  CaseID = @CaseID
           AND Status = 'Pending Approval';
END


GO
CREATE OR ALTER PROCEDURE [dbo].[usp_CompleteApproval]
@CaseID INT, @Approved BIT, @ApprovedBy NVARCHAR (100), @Comment NVARCHAR (500)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.FraudCases
    SET    Status          = CASE WHEN @Approved = 1 THEN 'Closed' ELSE 'In Review' END,
           ClosedOn        = CASE WHEN @Approved = 1 THEN SYSDATETIME() ELSE NULL END,
           TurnaroundHours = CASE WHEN @Approved = 1 THEN DATEDIFF(HOUR, CreatedOn, SYSDATETIME()) ELSE NULL END,
           ApprovedBy      = @ApprovedBy,
           ApprovedOn      = SYSDATETIME(),
           ApprovalComment = @Comment,
           ApprovalSentOn  = CASE WHEN @Approved = 1 THEN ApprovalSentOn ELSE NULL END
    WHERE  CaseID = @CaseID
           AND Status = 'Pending Approval';
END