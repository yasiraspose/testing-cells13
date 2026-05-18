using System;
using Aspose.Cells;

class ConvertExcelToPdf
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path where the high‑resolution PDF will be saved
        string destPath = "output.pdf";

        // Load the workbook from the XLSX file (create + load)
        Workbook workbook = new Workbook(sourcePath);

        // Ensure all formulas are calculated before saving
        workbook.CalculateFormula();

        // Save the workbook as a PDF file (save)
        workbook.Save(destPath, SaveFormat.Pdf);

        Console.WriteLine("Excel file has been successfully converted to PDF.");
    }
}