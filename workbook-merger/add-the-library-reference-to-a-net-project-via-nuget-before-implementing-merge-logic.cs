// Install the Aspose.Cells library via NuGet before running this code:
//   PM> Install-Package Aspose.Cells

using System;
using Aspose.Cells;

namespace MergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put some sample values to visualize the merge
            cells[0, 0].PutValue("Merged Header");
            cells[1, 0].PutValue("Row 2");
            cells[2, 0].PutValue("Row 3");

            // Merge cells starting at row 0, column 0 spanning 3 rows and 2 columns
            // Using the Cells.Merge method (firstRow, firstColumn, totalRows, totalColumns)
            cells.Merge(0, 0, 3, 2);

            // Save the workbook to a file (macro‑enabled format is not required for merge)
            workbook.Save("MergedCellsOutput.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with merged cells.");
        }
    }
}