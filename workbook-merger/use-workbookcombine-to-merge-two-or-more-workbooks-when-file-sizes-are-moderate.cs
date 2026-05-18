using System;
using Aspose.Cells;

namespace AsposeCellsCombineDemo
{
    class Program
    {
        static void Main()
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook();
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Workbook Data");

            // Create the destination workbook (empty or existing) and add data
            Workbook destWorkbook = new Workbook(FileFormatType.Xlsx);
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Workbook Data");

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook to disk in XLSX format
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}