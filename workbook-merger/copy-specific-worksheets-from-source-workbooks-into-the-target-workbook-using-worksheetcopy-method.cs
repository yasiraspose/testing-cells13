using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths of source workbooks
            string[] sourceFiles = { "source1.xlsx", "source2.xlsx" };

            // Names of worksheets to copy from each source workbook
            string[] sheetsToCopy = { "Report", "Summary" };

            // Create an empty target workbook
            Workbook targetWorkbook = new Workbook();
            // Remove the default sheet created with a new workbook
            targetWorkbook.Worksheets.Clear();

            foreach (string sourcePath in sourceFiles)
            {
                // Skip if the source file does not exist
                if (!File.Exists(sourcePath))
                    continue;

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                foreach (string sheetName in sheetsToCopy)
                {
                    // Verify the worksheet exists in the source workbook
                    Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];
                    if (sourceSheet == null)
                        continue;

                    // Determine a unique name for the destination sheet
                    string destSheetName = sheetName;
                    int duplicateIndex = 1;
                    while (targetWorkbook.Worksheets[destSheetName] != null)
                    {
                        destSheetName = $"{sheetName}_{duplicateIndex}";
                        duplicateIndex++;
                    }

                    // Add a new worksheet to the target workbook with the determined name
                    Worksheet destSheet = targetWorkbook.Worksheets.Add(destSheetName);

                    // Copy the contents and formats from the source worksheet to the destination worksheet
                    sourceSheet.Copy(destSheet);
                }
            }

            // Save the combined workbook
            targetWorkbook.Save("CombinedWorkbook.xlsx");
        }
    }
}