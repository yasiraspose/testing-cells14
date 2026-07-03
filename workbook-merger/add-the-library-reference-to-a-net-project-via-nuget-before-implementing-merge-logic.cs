// Install Aspose.Cells via NuGet before compiling this code:
//   PM> Install-Package Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Prepare temporary files to merge
            string[] sourceFiles = new string[2];
            sourceFiles[0] = Path.GetTempFileName().Replace(".tmp", ".xlsx");
            sourceFiles[1] = Path.GetTempFileName().Replace(".tmp", ".xlsx");

            // Create first workbook and add some data
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from Workbook 1");
            wb1.Save(sourceFiles[0]);

            // Create second workbook and add some data
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Data from Workbook 2");
            wb2.Save(sourceFiles[1]);

            // Define temporary cache file and destination merged file
            string cacheFile = Path.GetTempFileName();
            string mergedFile = Path.Combine(Path.GetTempPath(), "MergedResult.xlsx");

            try
            {
                // Merge the source workbooks into a single workbook
                CellsHelper.MergeFiles(sourceFiles, cacheFile, mergedFile);

                // Load the merged workbook to verify the result
                Workbook mergedWorkbook = new Workbook(mergedFile);
                Console.WriteLine("Merged workbook contains {0} worksheets.", mergedWorkbook.Worksheets.Count);
                Console.WriteLine("Sheet1 A1: " + mergedWorkbook.Worksheets[0].Cells["A1"].StringValue);
                Console.WriteLine("Sheet2 A1: " + mergedWorkbook.Worksheets[1].Cells["A1"].StringValue);
            }
            finally
            {
                // Clean up temporary files
                foreach (var file in sourceFiles)
                {
                    if (File.Exists(file)) File.Delete(file);
                }
                if (File.Exists(cacheFile)) File.Delete(cacheFile);
                // Optionally delete mergedFile if not needed
                // if (File.Exists(mergedFile)) File.Delete(mergedFile);
            }
        }
    }
}