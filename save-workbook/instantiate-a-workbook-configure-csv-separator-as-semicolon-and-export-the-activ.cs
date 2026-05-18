using System;
using Aspose.Cells;

class ExportActiveSheetToCsv
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Access the active worksheet (index 0 by default) and add sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Ensure the first worksheet is the active one (optional, shown for clarity)
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Configure CSV save options:
        // - Use semicolon as the field separator
        // - Export only the active sheet (ExportAllSheets = false is the default)
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ';';
        csvOptions.ExportAllSheets = false;

        // Save the active worksheet to a CSV file using the configured options
        workbook.Save("ActiveSheet.csv", csvOptions);
    }
}