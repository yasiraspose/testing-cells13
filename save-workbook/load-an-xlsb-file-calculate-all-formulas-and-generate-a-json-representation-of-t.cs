using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Path to the source XLSB file
        string inputPath = "input.xlsb";

        // Path where the JSON representation will be saved
        string outputJsonPath = "output.json";

        // Load the XLSB workbook (using the provided LoadOptions)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb)
        {
            ParsingFormulaOnOpen = true
        };

        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Convert the workbook data to a serializable structure
        var workbookData = new List<object>();

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var sheetData = new Dictionary<string, object>
            {
                ["Name"] = sheet.Name
            };

            var rows = new List<Dictionary<string, object>>();
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int r = 0; r <= maxRow; r++)
            {
                var rowDict = new Dictionary<string, object>();

                for (int c = 0; c <= maxCol; c++)
                {
                    Cell cell = cells[r, c];
                    if (cell != null && cell.Type != CellValueType.IsNull)
                    {
                        // Use column letters (A, B, ...) as keys
                        string columnName = CellsHelper.ColumnIndexToName(c);
                        rowDict[columnName] = cell.Value;
                    }
                }

                // Add only rows that contain at least one cell
                if (rowDict.Count > 0)
                {
                    rows.Add(rowDict);
                }
            }

            sheetData["Rows"] = rows;
            workbookData.Add(sheetData);
        }

        // Serialize the structure to JSON and write to file
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(workbookData, jsonOptions);
        File.WriteAllText(outputJsonPath, json);

        Console.WriteLine($"Workbook data has been exported to JSON file: {outputJsonPath}");
    }
}