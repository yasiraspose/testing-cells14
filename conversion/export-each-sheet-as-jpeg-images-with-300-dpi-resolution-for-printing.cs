using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ExportSheetsToJpeg
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Configure image rendering options for JPEG at 300 DPI
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,          // Output format
                HorizontalResolution = 300,          // 300 DPI horizontally
                VerticalResolution = 300,            // 300 DPI vertically
                OnePagePerSheet = false              // Allow multiple pages per sheet if needed
            };

            // Create a SheetRender instance for the current worksheet
            SheetRender renderer = new SheetRender(sheet, options);

            // Export each rendered page of the worksheet as a separate JPEG file
            for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
            {
                // Build a file name that includes the sheet name and page number
                string outputFile = $"{sheet.Name}_Page{pageIndex + 1}.jpg";

                // Render the specified page to the JPEG file
                renderer.ToImage(pageIndex, outputFile);
            }
        }
    }
}