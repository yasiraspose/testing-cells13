using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

class StrongTypedColumnsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // --------------------------------------------------------------------
        // 1. Define column headers (strongly typed)
        // --------------------------------------------------------------------
        cells[0, 0].PutValue("Id");        // Integer column
        cells[0, 1].PutValue("Name");      // String column
        cells[0, 2].PutValue("Price");     // Decimal column
        cells[0, 3].PutValue("Created");   // DateTime column

        // --------------------------------------------------------------------
        // 2. Apply number formats to enforce visual type safety
        // --------------------------------------------------------------------
        // Integer format (no decimals)
        Style intStyle = workbook.CreateStyle();
        intStyle.Number = 0; // General integer format
        cells.Columns[0].ApplyStyle(intStyle, new StyleFlag { NumberFormat = true });

        // Decimal format (two decimal places)
        Style decimalStyle = workbook.CreateStyle();
        decimalStyle.Number = 2; // Fixed-point with 2 decimals
        cells.Columns[2].ApplyStyle(decimalStyle, new StyleFlag { NumberFormat = true });

        // Date format (mm-dd-yy)
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Number = 14; // Built‑in date format
        cells.Columns[3].ApplyStyle(dateStyle, new StyleFlag { NumberFormat = true });

        // --------------------------------------------------------------------
        // 3. Prepare strongly typed data objects
        // --------------------------------------------------------------------
        var products = new List<Product>
        {
            new Product { Id = 1, Name = "Apple",  Price = 0.99m, Created = DateTime.Today },
            new Product { Id = 2, Name = "Banana", Price = 1.25m, Created = DateTime.Today.AddDays(-1) }
        };

        // --------------------------------------------------------------------
        // 4. Import objects with ImportCustomObjects (type‑aware)
        // --------------------------------------------------------------------
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = false,   // Header already written
            InsertRows = true,          // Add rows as needed
            ConvertNumericData = true,  // Convert numeric strings to numbers
            DateFormat = "yyyy-MM-dd"
        };
        worksheet.Cells.ImportCustomObjects(products, 1, 0, importOptions);

        // --------------------------------------------------------------------
        // 5. Export to DataTable with type safety
        // --------------------------------------------------------------------
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            ExportColumnName = true    // Use first row as column names
        };
        DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, products.Count + 1, 4, exportOptions);

        // Display resulting column data types
        foreach (DataColumn col in dataTable.Columns)
        {
            Console.WriteLine($"{col.ColumnName} => {col.DataType}");
        }

        // --------------------------------------------------------------------
        // 6. Save the workbook (XLSX)
        // --------------------------------------------------------------------
        workbook.Save("StrongTypedColumns.xlsx");
    }

    // Simple POCO representing a row of data
    public class Product
    {
        public int Id { get; set; }               // Integer
        public string Name { get; set; } = "";    // String
        public decimal Price { get; set; }        // Decimal
        public DateTime Created { get; set; }     // DateTime
    }
}