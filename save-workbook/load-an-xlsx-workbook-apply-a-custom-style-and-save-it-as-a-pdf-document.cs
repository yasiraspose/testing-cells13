using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsPdfConversion
{
    class Program
    {
        static void Main()
        {
            // Paths for source Excel file and destination PDF file
            string sourcePath = "input.xlsx";
            string destPath = "output.pdf";

            // Load the existing workbook using the constructor rule
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a custom style (rule: Workbook.CreateStyle)
            Style customStyle = workbook.CreateStyle();

            // Set background color to light yellow
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Configure font: Arial, 12pt, bold
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.Font.IsBold = true;

            // Apply the style to a specific range (e.g., A1:B2)
            var range = worksheet.Cells.CreateRange("A1:B2");
            range.ApplyStyle(customStyle, new StyleFlag { All = true });

            // Save the workbook as PDF using the Save(string, SaveFormat) rule
            workbook.Save(destPath, SaveFormat.Pdf);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook loaded, style applied, and saved as PDF successfully.");
        }
    }
}