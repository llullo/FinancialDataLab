using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using FinancialDataLab.Excel;
using FinancialDataLab.Infrastructure.Repositories;
namespace FinancialDataLab
{
    public partial class FinancialDataRibbon
    {
        private void FinancialDataRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void btnLoadObservations_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            try
            {
                // Get the active Excel workbook.
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                if (workbook == null)
                {
                    MessageBox.Show(
                        "Please open an Excel workbook first.",
                        "Financial Data Lab");
                    return;
                }

                // Read the SQL Server password from the local environment.
                string password = Environment.GetEnvironmentVariable(
                    "FINANCIALDATALAB_SQL_PASSWORD",
                    EnvironmentVariableTarget.User);

                if (string.IsNullOrWhiteSpace(password))
                {
                    MessageBox.Show(
                        "SQL Server password is not configured.",
                        "Financial Data Lab");
                    return;
                }

                // Configure the local SQL Server connection.
                var connectionBuilder = new SqlConnectionStringBuilder
                {
                    DataSource = "127.0.0.1,1433",
                    InitialCatalog = "FinancialDataLab",
                    UserID = "sa",
                    Password = password,
                    Encrypt = true,
                    TrustServerCertificate = true,
                    ConnectTimeout = 5
                };

                // Create the ADO.NET repository.
                var repository = new FinancialObservationRepository(
                    connectionBuilder.ConnectionString);

                // Retrieve the first 100 observations using OFFSET/FETCH.
                var observations = repository.GetPage(
                    "EURUSD_DEMO",
                    new DateTime(2026, 11, 1),
                    new DateTime(2027, 3, 1),
                    pageNumber: 1,
                    pageSize: 100);

                // Export the results to Excel.
                var exporter = new ExcelObservationExporter();

                exporter.Export(workbook, observations);

                MessageBox.Show(
                    $"{observations.Count} observations loaded successfully.",
                    "Financial Data Lab",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(
                    $"SQL Server error:\n{ex.Message}",
                    "Financial Data Lab",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Unable to load observations:\n{ex.Message}",
                    "Financial Data Lab",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

    }
}
