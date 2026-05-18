using System;
using System.IO;
using Aspose.Cells;

class MergeLargeXlsFiles
{
    static void Main()
    {
        // Paths to the source XLS files that need to be merged.
        string[] sourceFiles = new string[]
        {
            "LargeFile1.xls",
            "LargeFile2.xls",
            "LargeFile3.xls"
        };

        // Temporary cache file required by the MergeFiles method.
        string cacheFile = "mergeCache.tmp";

        // Destination file where the merged result will be saved.
        string outputFile = "MergedResult.xls";

        try
        {
            // Merge the specified files into a single workbook.
            CellsHelper.MergeFiles(sourceFiles, cacheFile, outputFile);
            Console.WriteLine($"Files merged successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during merging: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary cache file.
            if (File.Exists(cacheFile))
                File.Delete(cacheFile);
        }

        // Optional verification: load the merged workbook and display worksheet count.
        try
        {
            Workbook mergedWorkbook = new Workbook(outputFile);
            Console.WriteLine($"Merged workbook contains {mergedWorkbook.Worksheets.Count} worksheet(s).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading merged workbook: {ex.Message}");
        }
    }
}