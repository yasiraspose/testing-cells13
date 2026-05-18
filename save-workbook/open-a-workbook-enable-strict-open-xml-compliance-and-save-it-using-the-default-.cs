using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the workbook to open
        string inputPath = "input.xlsx";

        // Open the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Enable strict OOXML compliance (ISO/IEC 29500:2008 Strict)
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Save the workbook using a default file name
        workbook.Save("StrictComplianceWorkbook.xlsx");
    }
}