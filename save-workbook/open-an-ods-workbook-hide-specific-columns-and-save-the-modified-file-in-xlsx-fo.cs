using System;
using Aspose.Cells;

class HideColumnsInOds
{
    static void Main()
    {
        // Load the ODS workbook from disk
        Workbook workbook = new Workbook("input.ods");

        // Get the first worksheet (you can choose any worksheet as needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns B, C, and D (zero‑based index: start at 1, hide 3 columns)
        int startColumn = 1;   // Column B
        int columnCount = 3;   // B, C, D
        cells.HideColumns(startColumn, columnCount);

        // Save the modified workbook as XLSX
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}