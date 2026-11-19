
using System;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using FinancialDataLab.Infrastructure.Repositories;
using Xunit;

namespace FinancialDataLab.IntegrationTests.Repositories
{
    public class FinancialObservationRepositoryTests
    {
        [Fact]
        public void GetByIndicator_WhenDataExists_ReturnsObservations()
        {
            // Arrange
            string connectionString = BuildConnectionString();

            var repository =
                new FinancialObservationRepository(connectionString);

            // Act
            var observations = repository.GetByIndicator(
                "EURUSD",
                new DateTime(2026, 11, 16),
                new DateTime(2026, 11, 18));

            // Assert
            Assert.Equal(3, observations.Count);
            Assert.Equal("EURUSD", observations[0].IndicatorCode);
            Assert.Equal(1.158200m, observations[0].Value);
        }

        [Fact]
        public void GetNextPage_ReturnsRemainingRowsWithoutDuplicates()
        {
            var repository = new FinancialObservationRepository(
                BuildConnectionString());

            DateTime from = new DateTime(2026, 11, 16);
            DateTime to = new DateTime(2026, 11, 18);

            var firstPage = repository.GetNextPage(
                "EURUSD", from, to, 2, null, null);

            Assert.Equal(2, firstPage.Count);

            var last = firstPage[firstPage.Count - 1];

            var secondPage = repository.GetNextPage(
                "EURUSD",
                from,
                to,
                2,
                last.ObservationDate,
                last.Id);

            Assert.Single(secondPage);
            Assert.DoesNotContain(
                secondPage,
                row => firstPage.Any(first => first.Id == row.Id));
        }

        private static string BuildConnectionString()
        {
            // Search upwards for the solution root.
            var directory = new DirectoryInfo(
                AppDomain.CurrentDomain.BaseDirectory);

            while (directory != null &&
                   !File.Exists(Path.Combine(
                       directory.FullName,
                       "FinancialDataLab.slnx")))
            {
                directory = directory.Parent;
            }

            if (directory == null)
            {
                throw new FileNotFoundException(
                    "FinancialDataLab solution root not found.");
            }

            string envPath = Path.Combine(
                directory.FullName, "docker", ".env");

            string passwordLine = File.ReadLines(envPath)
                .First(line => line.StartsWith(
                    "MSSQL_SA_PASSWORD=",
                    StringComparison.Ordinal));

            string password = passwordLine.Substring(
                "MSSQL_SA_PASSWORD=".Length).Trim().Trim('"');

            var builder = new SqlConnectionStringBuilder
            {
                DataSource = "127.0.0.1,1433",
                InitialCatalog = "FinancialDataLab",
                UserID = "sa",
                Password = password,
                Encrypt = true,
                TrustServerCertificate = true,
                ConnectTimeout = 5
            };

            return builder.ConnectionString;
        }
    }
}
