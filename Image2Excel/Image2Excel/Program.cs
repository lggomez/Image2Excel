using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

using ImageSharp;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using Color = System.Drawing.Color;
using Image = ImageSharp.Image;
using Shape = Microsoft.Office.Interop.Excel.Shape;

namespace lggomez.Image2Excel
{
    class Program
    {
        private static readonly Stopwatch Watch = Stopwatch.StartNew();

        private static readonly ThreadLocal<Stopwatch> LocalWatch = new ThreadLocal<Stopwatch>(() => Watch);

        private static int previousProgressValue;

        private static int processedPixelCount;

        static void Main(string[] args)
        {
            var imagePath = args[0]; //Default command line arg is C:\\1b.jpg

            GenerateExcelFromImagePath(imagePath, new Progress<int>(ReportProgress)).Wait();
        }

        private static async Task<int> GenerateExcelFromImagePath(string imagePath, IProgress<int> progress)
        {
            const int ExcelMaxRows = 1048576;
            const int ExcelMaxColumns = 16384;

            using (FileStream stream = File.OpenRead(imagePath))
            using (Image image = new Image(stream))
            {
                AdjustImageSize(image, ExcelMaxRows, ExcelMaxColumns);
                return await GenerateExcelWorksheet(image, progress);
            }
        }

        static async Task<int> GenerateExcelWorksheet(Image image, IProgress<int> progress)
        {
            var excelApplication = InitializeExcelApplication();
            var rightmostCell = GenerateExcelColumnUpperBound(image);

            Console.WriteLine("Converting image...");

            int cellCount = await Task.Run<int>(
                                () =>
                                    {
                                        previousProgressValue = 0;
                                        LocalWatch.Value.Restart();

                                        Parallel.For(1L, image.Height + 1,
                                           i => {
                                               //LocalWatch.Value.Start();
                                               Range rowRange = excelApplication?.Range["A" + i, rightmostCell + i];

                                               if (rowRange == null)
                                               {
                                                   throw new InvalidOperationException(
                                                          "Could not get a range.Check to be sure you have the correct versions of the office DLLs.");
                                               }

                                               for (var j = 1; j <= rowRange.Cells.Count; j++)
                                               {
                                                   var cell = rowRange.Cells[j];
                                                   var pixel = image.Pixels[(i - 1) * image.Width + j];

                                                   cell.Interior.Color =
                                                    ColorTranslator.ToOle(Color.FromArgb(pixel.R, pixel.G, pixel.B));
                                               }

                                               Interlocked.Add(ref processedPixelCount, rowRange.Cells.Count);

                                               long progressValue = (processedPixelCount * 100) / (image.Width * image.Height);
                                               if (previousProgressValue != progressValue)
                                               {
                                                   progress?.Report((int)progressValue);
                                                   Interlocked.Exchange(ref previousProgressValue, (int)progressValue);
                                               }
                                           } );

                                        ResizeCells(excelApplication);
                                        AdjustExcelWindow(excelApplication);

                                        LocalWatch.Value.Stop();
                                        return 100;
                                    });

            return cellCount;
        }

        private static void AdjustImageSize(Image image, int excelMaxRows, int excelMaxColumns)
        {
            int newHeigth = image.Height;
            int newWidth = image.Width;
            ValidateImagePixelCount(image);
            bool resize = false;

            if (image.Height > excelMaxRows)
            {
                newHeigth = excelMaxRows;
                newWidth = newHeigth * image.Width / image.Height;
                resize = true;
            }

            if (image.Width > excelMaxColumns)
            {
                newHeigth = excelMaxColumns * newHeigth / newWidth;
                newWidth = excelMaxColumns;
                resize = true;
            }

            if (resize) image.Resize(newHeigth, newWidth);
        }

        private static void ValidateImagePixelCount(Image image)
        {
            int pixelCount = image.Pixels.Length;
            int expectedPixelCount = image.Height * image.Width;

            if (pixelCount != expectedPixelCount)
                Console.WriteLine(
                    $"WARNING: Image pixel count does not match the calculated pixel count (H*W) - expected:{expectedPixelCount} actual:{pixelCount}");

            // throw new ImageProcessingException($"WARNING: Image pixel count does not match the calculated pixel count (H*W) - expected:{expectedPixelCount} actual:{pixelCount}");
        }

        private static void AdjustExcelWindow(Application excelApplication)
        {
            excelApplication.Visible = true;
            excelApplication.WindowState = XlWindowState.xlMaximized;
            excelApplication.ActiveWindow.Zoom = 10;
        }

        private static void ResizeCells(Application excelApplication)
        {
            ((Worksheet)excelApplication.ActiveSheet).Columns.ColumnWidth = 2;
            ((Worksheet)excelApplication.ActiveSheet).Rows.EntireColumn.RowHeight =
                ((Range)((Worksheet)excelApplication.ActiveSheet).Cells[1]).Width;
            foreach (Shape shape in excelApplication.ActiveSheet.Shapes)
            {
                shape.LockAspectRatio = MsoTriState.msoTrue;
            }
        }

        private static Application InitializeExcelApplication()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            Application application = new Application();

            if (application == null)
            {
                throw new InvalidOperationException(
                          "EXCEL could not be started. Check that your office installation and project references are correct.");
            }

            application.Visible = false;

            Workbook workbook = application.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            if (worksheet == null)
            {
                throw new InvalidOperationException(
                          "Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            return application;
        }

        private static string GenerateExcelColumnUpperBound(Image image)
        {
            char[] alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

            var temp = image.Width;
            string rightmostCell = string.Empty;

            while (temp > 0)
            {
                int tempRemainder;
                temp = Math.DivRem(temp, 26, out tempRemainder);
                rightmostCell = alphabet[tempRemainder - 1] + rightmostCell;
            }

            return rightmostCell;
        }

        static void ReportProgress(int value)
        {
            Console.WriteLine($"\tConverting: %{value} (elapsed: {LocalWatch.Value.Elapsed.Minutes}m {LocalWatch.Value.Elapsed.Seconds}s)");
        }
    }
}