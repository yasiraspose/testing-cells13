using System;
using System.Collections;
using Aspose.Cells;

class RowEnumeratorDemo
{
    static void Main()
    {
        // Load a large workbook in streaming (lightweight) mode.
        // The LightCellsDataHandler processes rows one‑by‑one without loading the whole sheet into memory.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new StreamingRowHandler();

        // Replace "LargeFile.xlsx" with the path to your source file.
        Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions);
        Worksheet worksheet = workbook.Worksheets[0];

        // Use a synchronized enumerator to safely traverse rows.
        // reversed = false (normal order), sync = true (keeps enumerator in sync with any modifications).
        IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator(false, true);

        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;
            Console.WriteLine($"Row index: {row.Index}, Height: {row.Height}");
        }

        // Optionally save the workbook after processing.
        workbook.Save("Processed.xlsx");
    }

    // LightCellsDataHandler implementation that processes rows lazily.
    private class StreamingRowHandler : LightCellsDataHandler
    {
        // Called when a worksheet starts to be processed.
        public bool StartSheet(Worksheet sheet)
        {
            // Process all worksheets.
            return true;
        }

        // Called before a row is processed; return true to process the row.
        public bool StartRow(int rowIndex)
        {
            // Process every row.
            return true;
        }

        // Called after a row has been instantiated; you can read its properties here.
        public bool ProcessRow(Row row)
        {
            // Example: output the row index during streaming.
            Console.WriteLine($"[Handler] Processing row {row.Index}");
            return true; // Continue processing.
        }

        // Called before a cell is processed; return true to process the cell.
        public bool StartCell(int columnIndex) => true;

        // Called after a cell has been instantiated; you can read its value here.
        public bool ProcessCell(Cell cell) => true;
    }
}