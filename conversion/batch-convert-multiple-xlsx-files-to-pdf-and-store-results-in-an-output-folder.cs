using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class BatchXlsxToPdfConverter
    {
        /// <summary>
        /// Converts all .xlsx files in the specified input directory to PDF files
        /// and saves them in the specified output directory.
        /// </summary>
        /// <param name="inputFolder">Folder containing source .xlsx files.</param>
        /// <param name="outputFolder">Folder where converted PDF files will be stored.</param>
        public static void Run(string inputFolder, string outputFolder)
        {
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            Directory.CreateDirectory(outputFolder);

            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            if (xlsxFiles.Length == 0)
            {
                Console.WriteLine("No .xlsx files found to convert.");
                return;
            }

            foreach (string sourcePath in xlsxFiles)
            {
                try
                {
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input folder and output folder.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: AsposeCellsBatchConversion <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            BatchXlsxToPdfConverter.Run(inputFolder, outputFolder);
        }
    }
}