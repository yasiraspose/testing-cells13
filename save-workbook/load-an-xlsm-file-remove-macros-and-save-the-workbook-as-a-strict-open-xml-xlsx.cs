using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveMacroAndSaveAsXlsx
    {
        public static void Run()
        {
            string sourcePath = "input.xlsm";
            Workbook workbook = new Workbook(sourcePath);

            // Remove all VBA modules if a VBA project exists
            if (workbook.VbaProject != null)
            {
                var modules = workbook.VbaProject.Modules;
                for (int i = modules.Count - 1; i >= 0; i--)
                {
                    modules.RemoveAt(i);
                }
            }

            string destPath = "output.xlsx";
            workbook.Save(destPath, SaveFormat.Xlsx);
            Console.WriteLine($"Macros removed successfully. Saved to: {destPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveMacroAndSaveAsXlsx.Run();
        }
    }
}