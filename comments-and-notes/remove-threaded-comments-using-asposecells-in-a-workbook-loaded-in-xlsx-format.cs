using System;
using Aspose.Cells;

class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "InputWorkbook.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and clear all comments (including threaded comments)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clears both regular and threaded comments from the worksheet
            sheet.ClearComments();
        }

        // Save the workbook after removing the threaded comments
        string outputPath = "OutputWorkbook_NoThreadedComments.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Threaded comments removed and workbook saved to '{outputPath}'.");
    }
}