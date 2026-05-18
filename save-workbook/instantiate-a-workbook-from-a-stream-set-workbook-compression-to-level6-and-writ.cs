using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Prepare a source Excel stream (could be any existing stream)
            // ------------------------------------------------------------
            MemoryStream sourceStream = new MemoryStream();
            // Create a simple workbook and save it to the source stream
            Workbook tempWorkbook = new Workbook();
            tempWorkbook.Worksheets[0].Cells["A1"].PutValue("Sample");
            tempWorkbook.Save(sourceStream, SaveFormat.Xlsx);
            // Reset position so it can be read
            sourceStream.Position = 0;

            // ------------------------------------------------------------
            // 2. Instantiate a workbook from the prepared stream
            // ------------------------------------------------------------
            Workbook workbook = new Workbook(sourceStream);

            // ------------------------------------------------------------
            // 3. Configure OoxmlSaveOptions with Level6 compression
            // ------------------------------------------------------------
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.CompressionType = OoxmlCompressionType.Level6;

            // ------------------------------------------------------------
            // 4. Save the workbook into a new MemoryStream using the options
            // ------------------------------------------------------------
            using (MemoryStream outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, saveOptions);

                // The stream now contains the compressed XLSX file.
                // Reset position if further processing (e.g., reading) is needed.
                outputStream.Position = 0;

                // Example: write the stream to a file for verification (optional)
                // File.WriteAllBytes("CompressedOutput.xlsx", outputStream.ToArray());
            }

            // Clean up
            sourceStream.Dispose();
            tempWorkbook.Dispose();
            workbook.Dispose();
        }
    }
}