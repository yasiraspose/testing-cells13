using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert placeholder text
        sheet.Cells["A1"].PutValue("{Name}");

        // Replace placeholder with actual value
        workbook.Replace("{Name}", "John Doe");

        // Save as ODS (LibreOffice generator is default)
        OdsSaveOptions saveOptions = new OdsSaveOptions();

        workbook.Save("output.ods", saveOptions);
    }
}