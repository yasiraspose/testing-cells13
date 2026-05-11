using System;
using System.IO;
using Aspose.Cells;

namespace MergeWorkbooksExample
{
    class Program
    {
        static void Main()
        {
            // Define source file names
            string[] sourceFiles = new string[]
            {
                "Source1.xls",
                "Source2.xls",
                "Source3.xls"
            };

            // Ensure source workbooks exist (create simple ones if they don't)
            foreach (string fileName in sourceFiles)
            {
                string fullPath = Path.Combine(Directory.GetCurrentDirectory(), fileName);
                if (!File.Exists(fullPath))
                {
                    Workbook wb = new Workbook();
                    Worksheet ws = wb.Worksheets[0];
                    ws.Name = Path.GetFileNameWithoutExtension(fileName);
                    ws.Cells["A1"].PutValue($"Data from {fileName}");
                    wb.Save(fullPath);
                }
            }

            // Create destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Merge each source workbook into the destination workbook
            foreach (string fileName in sourceFiles)
            {
                string fullPath = Path.Combine(Directory.GetCurrentDirectory(), fileName);
                Workbook sourceWorkbook = new Workbook(fullPath);
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save merged workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.xls");
            destinationWorkbook.Save(outputPath);
        }
    }
}