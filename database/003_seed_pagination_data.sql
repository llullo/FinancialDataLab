
USE FinancialDataLab;
GO

-- Generate 300 synthetic financial observations.
-- Three observations share the same date.
-- Rerunning the script does not duplicate existing records.

;WITH Numbers AS
(
    SELECT 1 AS Number

    UNION ALL

    SELECT Number + 1
    FROM Numbers
    WHERE Number < 300
)
INSERT INTO dbo.FinancialObservations
(
    IndicatorCode,
    ObservationDate,
    Value
)
SELECT
    'EURUSD_DEMO',
    DATEADD(DAY, (Number - 1) / 3, CONVERT(DATE, '20261101', 112)),
    CAST(1.100000 + Number * 0.000100 AS DECIMAL(19,6))
FROM Numbers AS n
WHERE NOT EXISTS
(
    SELECT 1
    FROM dbo.FinancialObservations AS existing
    WHERE existing.IndicatorCode = 'EURUSD_DEMO'
      AND existing.ObservationDate =
          DATEADD(DAY, (n.Number - 1) / 3, CONVERT(DATE, '20261101', 112))
      AND existing.Value =
          CAST(1.100000 + n.Number * 0.000100 AS DECIMAL(19,6))
)
OPTION (MAXRECURSION 300);
GO
