using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Path where the resulting multi‑page TIFF will be saved
            string tiffPath = "output.tiff";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Configure image rendering options for TIFF output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Set the output image type to TIFF
                ImageType = ImageType.Tiff,

                // Use LZW compression to reduce file size without quality loss
                TiffCompression = TiffCompression.CompressionLZW,

                // High resolution preserves visual fidelity
                HorizontalResolution = 300,
                VerticalResolution = 300,

                // Render each worksheet page as a separate page in the TIFF
                OnePagePerSheet = true
            };

            // Create a renderer for the entire workbook
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render the whole workbook to a multi‑page TIFF file
            renderer.ToImage(tiffPath);

            Console.WriteLine($"Workbook successfully rendered to TIFF: {tiffPath}");
        }
    }
}