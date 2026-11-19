
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FinancialDataLab.Application.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace FinancialDataLab.Excel
{
    public class ExcelObservationExporter
    {
        private const string SheetName = "FDL Data";

        public ExcelInterop.Worksheet FindSheet(
            ExcelInterop.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            try
            {
                return (ExcelInterop.Worksheet)
                    workbook.Worksheets[SheetName];
            }
            catch (COMException)
            {
                return null;
            }
        }

        public int GetCurrentPage(ExcelInterop.Workbook workbook)
        {
            var sheet = FindSheet(workbook);

            if (sheet == null)
                return 0;

            object value = sheet.Range["F2"].Value2;

            if (value == null)
                return 0;

            return Convert.ToInt32(value);
        }

        public void Export(
            ExcelInterop.Workbook workbook,
            IReadOnlyList<FinancialObservation> observations,
            int pageNumber)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            if (observations == null)
                throw new ArgumentNullException(nameof(observations));

            if (pageNumber < 1)
                throw new ArgumentOutOfRangeException(nameof(pageNumber));

            var sheet = FindSheet(workbook);

            if (sheet == null)
            {
                sheet = (ExcelInterop.Worksheet)
                    workbook.Worksheets.Add();

                sheet.Name = SheetName;
            }

            // Clear only the data area managed by this add-in.
            sheet.Range["A1", "D501"].ClearContents();

            int rowCount = observations.Count + 1;
            const int columnCount = 4;

            var data = new object[rowCount, columnCount];

            data[0, 0] = "ID";
            data[0, 1] = "Indicator";
            data[0, 2] = "Observation Date";
            data[0, 3] = "Value";

            for (int i = 0; i < observations.Count; i++)
            {
                var observation = observations[i];

                data[i + 1, 0] = observation.Id;
                data[i + 1, 1] = observation.IndicatorCode;
                data[i + 1, 2] =
                    observation.ObservationDate.ToOADate();
                data[i + 1, 3] = (double)observation.Value;
            }

            ExcelInterop.Range targetRange = sheet.Range[
                sheet.Cells[1, 1],
                sheet.Cells[rowCount, columnCount]];

            // Bulk write: one assignment instead of one per cell.
            targetRange.Value2 = data;

            sheet.Range["A1", "D1"].Font.Bold = true;

            if (observations.Count > 0)
            {
                sheet.Range["C2", "C" + rowCount]
                    .NumberFormat = "dd/mm/yyyy";

                sheet.Range["D2", "D" + rowCount]
                    .NumberFormat = "0.000000";
            }

            // Page metadata belongs to this workbook.
            sheet.Range["F1"].Value2 = "Current Page";
            sheet.Range["F2"].Value2 = pageNumber;

            sheet.Range["G1"].Value2 = "Rows Loaded";
            sheet.Range["G2"].Value2 = observations.Count;

            sheet.Range["A:G"].Columns.AutoFit();

            sheet.Activate();
        }
    }
}
