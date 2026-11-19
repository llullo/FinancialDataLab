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



        private const int PageSize = 100;

        private void btnLoadObservations_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            LoadPage(1);
        }

        private void btnNextPage_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("Open a workbook first.");
                return;
            }

            var exporter = new ExcelObservationExporter();
            int currentPage = exporter.GetCurrentPage(workbook);

            if (currentPage == 0)
            {
                MessageBox.Show("Load observations first.");
                return;
            }

            LoadPage(currentPage + 1);
        }

        private void btnPreviousPage_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("Open a workbook first.");
                return;
            }

            var exporter = new ExcelObservationExporter();
            int currentPage = exporter.GetCurrentPage(workbook);

            if (currentPage <= 1)
            {
                MessageBox.Show("You are already on the first page.");
                return;
            }

            LoadPage(currentPage - 1);
        }

        private void LoadPage(int pageNumber)
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                if (workbook == null)
                {
                    MessageBox.Show("Open an Excel workbook first.");
                    return;
                }

                string password = Environment.GetEnvironmentVariable(
                    "FINANCIALDATALAB_SQL_PASSWORD",
                    EnvironmentVariableTarget.User);

                if (string.IsNullOrWhiteSpace(password))
                {
                    MessageBox.Show("SQL Server password is not configured.");
                    return;
                }

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

                var repository = new FinancialObservationRepository(
                    connectionBuilder.ConnectionString);

                var observations = repository.GetPage(
                    "EURUSD_DEMO",
                    new DateTime(2026, 11, 1),
                    new DateTime(2027, 3, 1),
                    pageNumber,
                    PageSize);

                if (observations.Count == 0)
                {
                    MessageBox.Show(
                        "No more observations available.",
                        "Financial Data Lab");
                    return;
                }

                var exporter = new ExcelObservationExporter();

                exporter.Export(workbook, observations, pageNumber);

                MessageBox.Show(
                    $"Page {pageNumber}: {observations.Count} observations loaded.",
                    "Financial Data Lab");
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
                    ex.ToString(),
                    "Financial Data Lab - Debug",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

    }
}
