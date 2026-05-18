using System;
using Aspose.Cells;

namespace MergeWorkbooksExample
{
    class Program
    {
        static void Main()
        {
            // Paths of the source XLS workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xls",
                "Source2.xls",
                "Source3.xls"
            };

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Load each source workbook using the Workbook(string) constructor
            // and combine it into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                Workbook sourceWorkbook = new Workbook(filePath);
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the merged workbook
            destinationWorkbook.Save("MergedOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}