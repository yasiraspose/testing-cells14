using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class BatchXlsxToPdfConverter
    {
        // Entry point
        public static void Main()
        {
            // Folder containing source XLSX files
            string inputFolder = @"C:\InputXlsx";
            // Folder where PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the input folder (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string sourcePath in xlsxFiles)
            {
                try
                {
                    // Build the destination PDF file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Convert the Excel file to PDF using Aspose.Cells ConversionUtility
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    // Log any conversion errors but continue processing remaining files
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}