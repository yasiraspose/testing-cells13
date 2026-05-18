using System;
using Aspose.Cells;
using System.IO;

namespace AsposeCellsTxtToXml
{
    class Program
    {
        static void Main()
        {
            // Paths for the source TXT file and the target XML file
            string txtFilePath = "input.txt";
            string xmlOutputPath = "output.xml";

            // -------------------------------------------------
            // 1. Load the TXT file into a Workbook with a custom delimiter
            // -------------------------------------------------
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            // Example custom delimiter – pipe character
            loadOptions.Separator = '|';
            // Load the text file using the options
            Workbook workbook = new Workbook(txtFilePath, loadOptions);

            // -------------------------------------------------
            // 2. Split the content of the first column into multiple columns
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            // Determine the number of rows to process (all rows that contain data)
            int totalRows = sheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based
            // Perform TextToColumns on column A (row index 0, column index 0)
            sheet.Cells.TextToColumns(0, 0, totalRows, loadOptions);

            // -------------------------------------------------
            // 3. Export the worksheet data to an XML file (requires an XML map)
            // -------------------------------------------------
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Use the first XmlMap found in the workbook
                string mapName = workbook.Worksheets.XmlMaps[0].Name;
                workbook.ExportXml(mapName, xmlOutputPath);
                Console.WriteLine($"Data exported to XML file: {xmlOutputPath}");
            }
            else
            {
                Console.WriteLine("No XmlMap defined in the workbook. XML export skipped.");
            }
        }
    }
}