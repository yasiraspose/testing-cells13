using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source SXC workbook
        string sourcePath = "input.sxc";

        // Load the workbook from the SXC file (load rule)
        using (Workbook workbook = new Workbook(sourcePath))
        {
            // Access the active worksheet (first worksheet) and rename it
            Worksheet activeSheet = workbook.Worksheets[0];
            activeSheet.Name = "RenamedSheet";

            // Export the renamed worksheet to a CSV file (save rule)
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);
        }

        Console.WriteLine("Workbook loaded, active worksheet renamed, and exported to CSV.");
    }
}