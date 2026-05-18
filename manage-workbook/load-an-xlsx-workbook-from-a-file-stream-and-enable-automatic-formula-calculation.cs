using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadAndCalculate
{
    class Program
    {
        static void Main()
        {
            try
            {
                        // Path to the source XLSX file
                        string filePath = "input.xlsx";

                        // Open the file as a read‑only stream
                        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        {
                            // Load the workbook from the stream using the provided constructor
                            Workbook workbook = new Workbook(stream);

                            // Enable automatic formula calculation by setting the calculation mode to Automatic
                            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                            // (Optional) Force calculation of all formulas immediately
                            workbook.CalculateFormula();

                            // Example: read a calculated cell value
                            Console.WriteLine("Value of A1 after calculation: " + workbook.Worksheets[0].Cells["A1"].Value);
                        }
        
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
}
    }
}