using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Add Aspose.Cells to the project via NuGet before running this code:
        //   Install-Package Aspose.Cells

        // -----------------------------------------------------------------
        // Prepare sample source workbooks (this would normally be existing files)
        // -----------------------------------------------------------------
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

        Workbook wb1 = new Workbook();                                   // create workbook
        wb1.Worksheets[0].Cells["A1"].PutValue("Data from file 1");
        wb1.Save(sourceFiles[0]);                                        // save workbook

        Workbook wb2 = new Workbook();                                   // create workbook
        wb2.Worksheets[0].Cells["A1"].PutValue("Data from file 2");
        wb2.Save(sourceFiles[1]);                                        // save workbook

        // -----------------------------------------------------------------
        // Merge the source files into a single workbook using CellsHelper
        // -----------------------------------------------------------------
        string cacheFile = "mergeCache.tmp";                             // temporary cache file required by MergeFiles
        string destFile = "MergedResult.xlsx";

        CellsHelper.MergeFiles(sourceFiles, cacheFile, destFile);        // merge logic (provided rule)

        // -----------------------------------------------------------------
        // Load the merged workbook to verify the result
        // -----------------------------------------------------------------
        Workbook merged = new Workbook(destFile);                        // load workbook
        Console.WriteLine("Merged workbook contains:");
        for (int i = 0; i < merged.Worksheets.Count; i++)
        {
            Console.WriteLine($"Sheet {i + 1} A1: {merged.Worksheets[i].Cells["A1"].StringValue}");
        }

        // -----------------------------------------------------------------
        // Clean up temporary files (optional)
        // -----------------------------------------------------------------
        foreach (string f in sourceFiles)
        {
            if (File.Exists(f)) File.Delete(f);
        }
        if (File.Exists(cacheFile)) File.Delete(cacheFile);
    }
}