/* ============================================================
   Phase 2: SQL Server staging, core tables, risk rules, views
   Run each numbered section in SSMS, in order.
   Section 3: change @DataPath to the folder holding your CSVs.
   ============================================================ */
-- 1) DATABASE AND SCHEMAS --------------------------------------
IF DB_ID('FraudHub') IS NULL
    CREATE DATABASE FraudHub;


GO
USE FraudHub;


GO
IF SCHEMA_ID('stg') IS NULL
    EXECUTE ('CREATE SCHEMA stg');

IF SCHEMA_ID('core') IS NULL
    EXECUTE ('CREATE SCHEMA core');


GO
-- 2) STAGING TABLES (everything as text so the load never fails) --
DROP TABLE IF EXISTS stg.players, stg.devices, stg.payments, stg.bonuses;

CREATE TABLE stg.players (
    player_id VARCHAR (10),
    full_name VARCHAR (100),
    email VARCHAR (150),
    country VARCHAR (5),
    registered_at VARCHAR (30)
);

CREATE TABLE stg.devices (
    player_id VARCHAR (10),
    device_id VARCHAR (10),
    ip_address VARCHAR (20)
);

CREATE TABLE stg.payments (
    txn_id VARCHAR (12),
    player_id VARCHAR (10),
    txn_time VARCHAR (30),
    txn_type VARCHAR (15),
    method VARCHAR (20),
    amount VARCHAR (20),
    status VARCHAR (15),
    card_hash VARCHAR (20)
);

CREATE TABLE stg.bonuses (
    bonus_id VARCHAR (10),
    player_id VARCHAR (10),
    claimed_at VARCHAR (30),
    bonus_amount VARCHAR (20),
    wagered_amount VARCHAR (20),
    hours_to_first_withdrawal VARCHAR (20)
);


GO
-- 3) LOAD CSVs INTO STAGING ------------------------------------
DECLARE @DataPath AS NVARCHAR (300) = N'C:\Users\REDTECH\Desktop\fraud_hub\data\';

DECLARE @sql AS NVARCHAR (MAX);

DECLARE @t TABLE (
    tbl SYSNAME);

INSERT  @t
VALUES ('players'),
('devices'),
('payments'),
('bonuses');

DECLARE @name AS SYSNAME;

DECLARE c CURSOR
    FOR SELECT tbl
        FROM   @t;

OPEN c;

FETCH NEXT FROM c INTO @name;

WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @sql = N'TRUNCATE TABLE stg.' + @name + N'; BULK INSERT stg.' + @name + N' FROM ''' + @DataPath + @name + N'.csv'' WITH (FIRSTROW = 2, FIELDTERMINATOR = '','', ROWTERMINATOR = ''0x0a'', CODEPAGE = ''65001'', TABLOCK);';
        EXECUTE sp_executesql @sql;
        FETCH NEXT FROM c INTO @name;
    END

CLOSE c;

DEALLOCATE c;


GO
-- Sanity check. Expect 2000 / 2956 / 54817 / 1036
SELECT 'players' AS t,
       COUNT(*) AS n
FROM   stg.players
UNION ALL
SELECT 'devices',
       COUNT(*)
FROM   stg.devices
UNION ALL
SELECT 'payments',
       COUNT(*)
FROM   stg.payments
UNION ALL
SELECT 'bonuses',
       COUNT(*)
FROM   stg.bonuses;


GO
-- 4) CORE TABLES (typed, keyed) ---------------------------------
DROP TABLE IF EXISTS core.risk_flags, core.bonuses, core.payments, core.devices, core.players, core.rule_weights;

CREATE TABLE core.players (
    player_id VARCHAR (10) PRIMARY KEY,
    full_name VARCHAR (100),
    email VARCHAR (150),
    country VARCHAR (5),
    registered_at DATETIME2
);

CREATE TABLE core.devices (
    player_id VARCHAR (10) FOREIGN KEY REFERENCES core.players (player_id),
    device_id VARCHAR (10),
    ip_address VARCHAR (20)
);

CREATE TABLE core.payments (
    txn_id VARCHAR (12) PRIMARY KEY,
    player_id VARCHAR (10) FOREIGN KEY REFERENCES core.players (player_id),
    txn_time DATETIME2,
    txn_type VARCHAR (15),
    method VARCHAR (20),
    amount DECIMAL (12, 2),
    status VARCHAR (15),
    card_hash VARCHAR (20)
);

CREATE TABLE core.bonuses (
    bonus_id VARCHAR (10) PRIMARY KEY,
    player_id VARCHAR (10) FOREIGN KEY REFERENCES core.players (player_id),
    claimed_at DATETIME2,
    bonus_amount DECIMAL (10, 2),
    wagered_amount DECIMAL (12, 2),
    hours_to_first_withdrawal DECIMAL (9, 1) NULL
);

CREATE TABLE core.rule_weights (
    rule_code VARCHAR (30) PRIMARY KEY,
    points INT NOT NULL
);

CREATE TABLE core.risk_flags (
    player_id VARCHAR (10),
    rule_code VARCHAR (30),
    detail VARCHAR (200)
);


GO
INSERT  core.rule_weights
VALUES ('MULTI_ACCOUNT_DEVICE', 50),
('SHARED_IP', 15),
('SHARED_CARD', 50),
('PAYMENT_VELOCITY', 40),
('BONUS_ABUSE', 45);

INSERT core.players
SELECT player_id,
       full_name,
       email,
       country,
       TRY_CAST (registered_at AS DATETIME2)
FROM   stg.players;

INSERT core.devices
SELECT player_id,
       device_id,
       ip_address
FROM   stg.devices;

INSERT core.payments
SELECT txn_id,
       player_id,
       TRY_CAST (txn_time AS DATETIME2),
       txn_type,
       method,
       TRY_CAST (amount AS DECIMAL (12, 2)),
       status,
       NULLIF (TRIM(card_hash), '')
FROM   stg.payments;

INSERT core.bonuses
SELECT bonus_id,
       player_id,
       TRY_CAST (claimed_at AS DATETIME2),
       TRY_CAST (bonus_amount AS DECIMAL (10, 2)),
       TRY_CAST (wagered_amount AS DECIMAL (12, 2)),
       TRY_CAST (NULLIF (TRIM(hours_to_first_withdrawal), '') AS DECIMAL (9, 1))
FROM   stg.bonuses;


GO
-- 5) STORED PROCEDURE: refresh risk flags (thresholds are parameters) --
CREATE OR ALTER PROCEDURE core.usp_refresh_risk_flags
@device_min INT=3, @ip_min INT=3, @card_min INT=3, @vel_min INT=5, @wager_max DECIMAL (9, 2)=5, @wd_hours_max DECIMAL (9, 2)=24
AS
BEGIN
    SET NOCOUNT ON;
    TRUNCATE TABLE core.risk_flags; -- Multi-accounting: shared device
    WITH s
    AS   (SELECT   device_id,
                   COUNT(DISTINCT player_id) AS n
          FROM     core.devices
          GROUP BY device_id
          HAVING   COUNT(DISTINCT player_id) >= @device_min)
    INSERT core.risk_flags
    SELECT DISTINCT d.player_id,
                    'MULTI_ACCOUNT_DEVICE',
                    CONCAT('device_id=', d.device_id, ' shared by ', s.n, ' players')
    FROM   core.devices AS d
           INNER JOIN
           s
           ON s.device_id = d.device_id; -- Multi-accounting: shared IP
    WITH s
    AS   (SELECT   ip_address,
                   COUNT(DISTINCT player_id) AS n
          FROM     core.devices
          GROUP BY ip_address
          HAVING   COUNT(DISTINCT player_id) >= @ip_min)
    INSERT core.risk_flags
    SELECT DISTINCT d.player_id,
                    'SHARED_IP',
                    CONCAT('ip_address=', d.ip_address, ' shared by ', s.n, ' players')
    FROM   core.devices AS d
           INNER JOIN
           s
           ON s.ip_address = d.ip_address; -- Shared payment card
    WITH s
    AS   (SELECT   card_hash,
                   COUNT(DISTINCT player_id) AS n
          FROM     core.payments
          WHERE    card_hash IS NOT NULL
          GROUP BY card_hash
          HAVING   COUNT(DISTINCT player_id) >= @card_min)
    INSERT core.risk_flags
    SELECT DISTINCT p.player_id,
                    'SHARED_CARD',
                    CONCAT('card ', p.card_hash, ' used by ', s.n, ' players')
    FROM   core.payments AS p
           INNER JOIN
           s
           ON s.card_hash = p.card_hash; -- Deposit velocity: deposits inside any 1-hour window starting at each deposit
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
    INSERT core.risk_flags
    SELECT   player_id,
             'PAYMENT_VELOCITY',
             CONCAT(MAX(n), ' deposits within 1 hour')
    FROM     v
    GROUP BY player_id
    HAVING   MAX(n) >= @vel_min;
    -- Bonus abuse: low wagering plus fast withdrawal
    INSERT core.risk_flags
    SELECT DISTINCT player_id,
                    'BONUS_ABUSE',
                    CONCAT('wagered ', CAST (wagered_amount / bonus_amount AS DECIMAL (6, 1)), 'x, withdrew after ', hours_to_first_withdrawal, 'h')
    FROM   core.bonuses
    WHERE  wagered_amount / bonus_amount < @wager_max
           AND hours_to_first_withdrawal < @wd_hours_max;
END


GO
-- 6) VIEW: one row per flagged player, scored and banded ----------
CREATE OR ALTER VIEW core.vw_daily_exceptions
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
-- 7) RUN IT AND CHECK --------------------------------------------
EXECUTE core.usp_refresh_risk_flags ;

SELECT   risk_band,
         COUNT(*) AS players
FROM     core.vw_daily_exceptions
GROUP BY risk_band; -- expect High 100, Medium 135

SELECT   TOP 20 *
FROM     core.vw_daily_exceptions
ORDER BY risk_score DESC;