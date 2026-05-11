using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook (can be .xlsx, .xls, etc.)
        string sourceFile = "input.xlsx";

        // Desired output CSV file that will contain all worksheets
        string outputCsv = "all_sheets.csv";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(sourceFile);

        // Create CSV save options and enable exporting all sheets
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.ExportAllSheets = true; // Export every worksheet into the same CSV file

        // Save the workbook as a single CSV file containing all worksheets
        workbook.Save(outputCsv, csvOptions);
    }
}