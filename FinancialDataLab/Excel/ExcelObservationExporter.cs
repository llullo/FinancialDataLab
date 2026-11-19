
using System;
using System.Collections.Generic;
using FinancialDataLab.Application.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace FinancialDataLab.Excel
{
    public class ExcelObservationExporter
    {
        public void Export(
            ExcelInterop.Workbook workbook,
            IReadOnlyList<FinancialObservation> observations)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            if (observations == null)
                throw new ArgumentNullException(nameof(observations));

            // Create a new worksheet without overwriting existing data.
            var sheet = (ExcelInterop.Worksheet)workbook.Worksheets.Add();
            sheet.Name = "FDL " + DateTime.Now.ToString("HHmmssfff");

            int rowCount = observations.Count + 1;
            const int columnCount = 4;

            // Prepare all values in managed memory.
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

            // Transfer the complete matrix in a single COM write.
            ExcelInterop.Range firstCell = (ExcelInterop.Range)sheet.Cells[1, 1];
            ExcelInterop.Range lastCell =
                (ExcelInterop.Range)sheet.Cells[rowCount, columnCount];

            ExcelInterop.Range targetRange =
                sheet.Range[firstCell, lastCell];

            targetRange.Value2 = data;

            // Format entire ranges rather than individual cells.
            ExcelInterop.Range header = sheet.Range["A1", "D1"];
            header.Font.Bold = true;

            if (observations.Count > 0)
            {
                ExcelInterop.Range dates =
                    sheet.Range["C2", "C" + rowCount];

                dates.NumberFormat = "dd/mm/yyyy";

                ExcelInterop.Range values =
                    sheet.Range["D2", "D" + rowCount];

                values.NumberFormat = "0.000000";
            }

            targetRange.Columns.AutoFit();
        }
    }
}
