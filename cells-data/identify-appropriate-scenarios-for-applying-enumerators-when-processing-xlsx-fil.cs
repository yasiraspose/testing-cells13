using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsEnumeratorScenarios
{
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // 1. Load an existing workbook (or create a new one if file not found)
            // ------------------------------------------------------------
            string inputPath = "InputData.xlsx";
            Workbook workbook;

            if (System.IO.File.Exists(inputPath))
            {
                // Load the workbook from file
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with sample data
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Name");
                ws.Cells["B1"].PutValue("Score");
                ws.Cells["A2"].PutValue("Alice");
                ws.Cells["B2"].PutValue(85);
                ws.Cells["A3"].PutValue("Bob");
                ws.Cells["B3"].PutValue(92);
                ws.Cells["A4"].PutValue("Charlie");
                ws.Cells["B4"].PutValue(78);
                // Save the sample file for future runs
                workbook.Save(inputPath);
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 2. Scenario: Enumerate all cells in the worksheet using Cells.GetEnumerator()
            // ------------------------------------------------------------
            Console.WriteLine("=== Enumerating all cells (non‑null values) ===");
            IEnumerator cellEnum = worksheet.Cells.GetEnumerator();
            while (cellEnum.MoveNext())
            {
                Cell cell = (Cell)cellEnum.Current;
                if (cell.Value != null)
                {
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }

            // ------------------------------------------------------------
            // 3. Scenario: Enumerate rows using RowCollection.GetEnumerator()
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Enumerating rows (index & height) ===");
            IEnumerator rowEnum = worksheet.Cells.Rows.GetEnumerator();
            while (rowEnum.MoveNext())
            {
                Row row = (Row)rowEnum.Current;
                Console.WriteLine($"Row {row.Index} - Height: {row.Height}");
            }

            // ------------------------------------------------------------
            // 4. Scenario: Enumerate rows in reverse order with synchronization
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Enumerating rows in reverse order (synchronized) ===");
            for (int i = worksheet.Cells.Rows.Count - 1; i >= 0; i--)
            {
                Row row = worksheet.Cells.Rows[i];
                Console.WriteLine($"Row {row.Index}");
            }

            // ------------------------------------------------------------
            // 5. Scenario: Enumerate cells of a specific row using Row.GetEnumerator()
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Enumerating cells of the first data row (row index 1) ===");
            Row dataRow = worksheet.Cells.Rows[1]; // second row (index starts at 0)
            IEnumerator rowCellEnum = dataRow.GetEnumerator();
            while (rowCellEnum.MoveNext())
            {
                Cell cell = (Cell)rowCellEnum.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // ------------------------------------------------------------
            // 6. Scenario: Enumerate a defined range using Range.GetEnumerator()
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Enumerating cells in range B2:C4 ===");
            AsposeRange range = worksheet.Cells.CreateRange("B2:C4");
            IEnumerator rangeEnum = range.GetEnumerator();
            while (rangeEnum.MoveNext())
            {
                Cell cell = (Cell)rangeEnum.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // ------------------------------------------------------------
            // 7. Scenario: Enumerate pivot fields if a pivot table exists
            // ------------------------------------------------------------
            // Create a simple pivot table if none exists
            if (worksheet.PivotTables.Count == 0)
            {
                // Ensure source data is contiguous
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(45);
                worksheet.Cells["A4"].PutValue("Apple");
                worksheet.Cells["B4"].PutValue(20);

                int pivotIdx = worksheet.PivotTables.Add("A1:B4", "E1", "SalesPivot");
                PivotTable pivot = worksheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product
                pivot.AddFieldToArea(PivotFieldType.Data, 1); // Quantity
                pivot.RefreshData();
                pivot.CalculateData();
            }

            Console.WriteLine("\n=== Enumerating pivot fields in Row area ===");
            PivotTable pt = worksheet.PivotTables[0];
            PivotFieldCollection rowFields = pt.RowFields;
            IEnumerator pivotFieldEnum = rowFields.GetEnumerator();
            while (pivotFieldEnum.MoveNext())
            {
                PivotField field = (PivotField)pivotFieldEnum.Current;
                Console.WriteLine($"Pivot Field: {field.Name}");
                // Enumerate items of this field
                IEnumerator itemEnum = field.PivotItems.GetEnumerator();
                while (itemEnum.MoveNext())
                {
                    PivotItem item = (PivotItem)itemEnum.Current;
                    Console.WriteLine($"  Item: {item.Value}");
                }
            }

            // ------------------------------------------------------------
            // 8. Save the workbook after enumeration (no modifications made)
            // ------------------------------------------------------------
            string outputPath = "EnumeratedOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}