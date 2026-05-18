using System;
using Aspose.Cells;

namespace WorkbookMergeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define file paths for source, destination, and output workbooks
            string sourcePath = "Source.xlsx";
            string destinationPath = "Destination.xlsx";
            string outputPath = "CombinedWorkbook.xlsx";

            // -------------------------------------------------
            // Create a source workbook and add sample data
            // -------------------------------------------------
            Workbook sourceWorkbook = new Workbook(); // using Workbook() constructor
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");
            sourceWorkbook.Save(sourcePath, SaveFormat.Xlsx); // save the source workbook

            // -------------------------------------------------
            // Create a destination workbook and add sample data
            // -------------------------------------------------
            Workbook destinationWorkbook = new Workbook(FileFormatType.Xlsx); // using Workbook(FileFormatType) constructor
            destinationWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");
            destinationWorkbook.Save(destinationPath, SaveFormat.Xlsx); // save the destination workbook

            // -------------------------------------------------
            // Load the workbooks from disk (demonstrates loading)
            // -------------------------------------------------
            Workbook src = new Workbook(sourcePath); // using Workbook(string) constructor
            Workbook dest = new Workbook(destinationPath); // using Workbook(string) constructor

            // -------------------------------------------------
            // Combine the source workbook into the destination workbook
            // -------------------------------------------------
            dest.Combine(src); // using Workbook.Combine method

            // -------------------------------------------------
            // Save the combined workbook to the output file
            // -------------------------------------------------
            dest.Save(outputPath, SaveFormat.Xlsx); // using Workbook.Save method

            Console.WriteLine($"Workbooks have been combined and saved to '{outputPath}'.");
        }
    }
}