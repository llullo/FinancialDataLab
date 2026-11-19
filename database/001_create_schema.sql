IF DB_ID(N'FinancialDataLab') IS NULL
BEGIN
    CREATE DATABASE FinancialDataLab;
END;
GO

USE FinancialDataLab;
GO

IF OBJECT_ID(N'dbo.FinancialObservations', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FinancialObservations
    (
        Id INT IDENTITY(1,1) NOT NULL,
        IndicatorCode VARCHAR(50) NOT NULL,
        ObservationDate DATETIME2(0) NOT NULL,
        Value DECIMAL(19,6) NOT NULL,

        CONSTRAINT PK_FinancialObservations
            PRIMARY KEY CLUSTERED (Id)
    );
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = N'IX_FinancialObservations_IndicatorDate'
      AND object_id = OBJECT_ID(N'dbo.FinancialObservations')
)
BEGIN
    CREATE NONCLUSTERED INDEX
        IX_FinancialObservations_IndicatorDate
    ON dbo.FinancialObservations
    (
        IndicatorCode,
        ObservationDate,
        Id
    )
    INCLUDE (Value);
END;
GO