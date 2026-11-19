USE FinancialDataLab;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.FinancialObservations
    WHERE IndicatorCode = 'EURUSD'
)
BEGIN
    INSERT INTO dbo.FinancialObservations
        (IndicatorCode, ObservationDate, Value)
    VALUES
        ('EURUSD', '2026-11-16', 1.158200),
        ('EURUSD', '2026-11-17', 1.160100),
        ('EURUSD', '2026-11-18', 1.157500);
END;
GO