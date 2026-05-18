using System;
using Aspose.Cells;

class ExportTabDelimited
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Laptop");
        worksheet.Cells["B2"].PutValue(999.99);
        worksheet.Cells["A3"].PutValue("Phone");
        worksheet.Cells["B3"].PutValue(699.99);
        worksheet.Cells["A4"].PutValue("Tablet");
        worksheet.Cells["B4"].PutValue(399.99);

        // Configure TxtSaveOptions for tab-delimited output
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
        txtSaveOptions.Separator = '\t'; // Use tab as the delimiter

        // Save the workbook as a tab-delimited TXT file
        workbook.Save("ExportedData.txt", txtSaveOptions);

        Console.WriteLine("Workbook exported as tab-delimited TXT successfully.");
    }
}