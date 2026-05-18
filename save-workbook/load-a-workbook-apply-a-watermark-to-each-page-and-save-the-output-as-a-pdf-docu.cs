using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create a font for the watermark
        RenderingFont font = new RenderingFont("Calibri", 68)
        {
            Italic = true,
            Bold = true,
            Color = Color.Blue
        };

        // Create a text watermark with the specified font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,
            Opacity = 0.3f,
            ScaleToPagePercent = 75,
            IsBackground = true
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied
        string outputPath = "output_watermark.pdf";
        workbook.Save(outputPath, pdfOptions);
    }
}