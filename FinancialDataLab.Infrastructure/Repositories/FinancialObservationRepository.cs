using FinancialDataLab.Application.Interfaces;
using FinancialDataLab.Application.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace FinancialDataLab.Infrastructure.Repositories
{
    public class FinancialObservationRepository
        : IFinancialObservationRepository
    {
        private readonly string _connectionString;

        public FinancialObservationRepository(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException(
                    "Connection string cannot be empty.",
                    nameof(connectionString));
            }

            _connectionString = connectionString;
        }

        public IReadOnlyList<FinancialObservation> GetByIndicator(
            string indicatorCode,
            DateTime fromDate,
            DateTime toDate)
        {
            if (string.IsNullOrWhiteSpace(indicatorCode))
            {
                throw new ArgumentException(
                    "Indicator code cannot be empty.",
                    nameof(indicatorCode));
            }

            if (fromDate > toDate)
            {
                throw new ArgumentException(
                    "From date cannot be greater than to date.");
            }

            var observations = new List<FinancialObservation>();

            using (var connection = new SqlConnection(_connectionString))
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
                    SELECT
                        Id,
                        IndicatorCode,
                        ObservationDate,
                        Value
                    FROM dbo.FinancialObservations
                    WHERE IndicatorCode = @IndicatorCode
                      AND ObservationDate >= @FromDate
                      AND ObservationDate <= @ToDate
                    ORDER BY ObservationDate, Id;";

                command.CommandType = CommandType.Text;

                command.Parameters.Add(
                    "@IndicatorCode", SqlDbType.VarChar, 50)
                    .Value = indicatorCode;

                command.Parameters.Add(
                    "@FromDate", SqlDbType.DateTime2)
                    .Value = fromDate;

                command.Parameters.Add(
                    "@ToDate", SqlDbType.DateTime2)
                    .Value = toDate;

                connection.Open();

                using (var reader = command.ExecuteReader())
                {
                    int idOrdinal = reader.GetOrdinal("Id");
                    int codeOrdinal = reader.GetOrdinal("IndicatorCode");
                    int dateOrdinal = reader.GetOrdinal("ObservationDate");
                    int valueOrdinal = reader.GetOrdinal("Value");

                    while (reader.Read())
                    {
                        observations.Add(new FinancialObservation
                        {
                            Id = reader.GetInt32(idOrdinal),
                            IndicatorCode = reader.GetString(codeOrdinal),
                            ObservationDate = reader.GetDateTime(dateOrdinal),
                            Value = reader.GetDecimal(valueOrdinal)
                        });
                    }
                }
            }

            return observations;
        }

        public IReadOnlyList<FinancialObservation> GetNextPage(
            string indicatorCode,
            DateTime fromDate,
            DateTime toDate,
            int pageSize,
            DateTime? lastObservationDate,
            int? lastId)
        {
            if (string.IsNullOrWhiteSpace(indicatorCode))
                throw new ArgumentException(
                    "Indicator code is required.", nameof(indicatorCode));

            if (fromDate > toDate)
                throw new ArgumentException("Invalid date range.");

            if (pageSize < 1 || pageSize > 500)
                throw new ArgumentOutOfRangeException(nameof(pageSize));

            if (lastObservationDate.HasValue != lastId.HasValue)
                throw new ArgumentException(
                    "Both cursor values must be provided together.");

            var results = new List<FinancialObservation>();

            using (var connection = new SqlConnection(_connectionString))
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
                    SELECT TOP (@PageSize)
                        Id, IndicatorCode, ObservationDate, Value
                    FROM dbo.FinancialObservations
                    WHERE IndicatorCode = @IndicatorCode
                      AND ObservationDate >= @FromDate
                      AND ObservationDate <= @ToDate";

                if (lastObservationDate.HasValue)
                {
                    command.CommandText += @"
                    AND (
                    ObservationDate > @LastDate
                    OR (
                        ObservationDate = @LastDate
                        AND Id > @LastId
                    )
                    )";

                    command.Parameters.Add(
                        "@LastDate", SqlDbType.DateTime2)
                        .Value = lastObservationDate.Value;

                    command.Parameters.Add(
                        "@LastId", SqlDbType.Int)
                        .Value = lastId.Value;
                }

                command.CommandText += @"
                    ORDER BY ObservationDate, Id;";

                command.Parameters.Add(
                    "@PageSize", SqlDbType.Int).Value = pageSize;

                command.Parameters.Add(
                    "@IndicatorCode", SqlDbType.VarChar, 50)
                    .Value = indicatorCode;

                command.Parameters.Add(
                    "@FromDate", SqlDbType.DateTime2)
                    .Value = fromDate;

                command.Parameters.Add(
                    "@ToDate", SqlDbType.DateTime2)
                    .Value = toDate;

                connection.Open();

                using (var reader = command.ExecuteReader())
                {
                    int id = reader.GetOrdinal("Id");
                    int code = reader.GetOrdinal("IndicatorCode");
                    int date = reader.GetOrdinal("ObservationDate");
                    int value = reader.GetOrdinal("Value");

                    while (reader.Read())
                    {
                        results.Add(new FinancialObservation
                        {
                            Id = reader.GetInt32(id),
                            IndicatorCode = reader.GetString(code),
                            ObservationDate = reader.GetDateTime(date),
                            Value = reader.GetDecimal(value)
                        });
                    }
                }
            }

            return results;
        }

    }
}