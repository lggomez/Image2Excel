using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        private static readonly Stopwatch LocalWatch = Stopwatch.StartNew();

        private static int ImageSize;

        private static long CellCount;

        public class ReportProgressStatus
        {
            public int Percentage;
            public long EllapsedMinutes;
            public long EllapsedSeconds;
            public long ProcessedPixelCount;
            public long PreviousProgressValue;
        }

        private static ReportProgressStatus ProgressStatus { get; } = new ReportProgressStatus();

        static void Main(string[] args)
        {
            if (args.Any())
            {
                var imagePath = args[0]; //Default command line arg is C:\\1b.jpg
                GenerateExcelFromImagePath(imagePath);
            }
            else
            {
                Console.WriteLine("Missing image filepath argument. Please enter an image path and try again.");
                Console.Read();
                Environment.Exit(1);
            }
        }

        private static void GenerateExcelFromImagePath(string imagePath)
        {
            const int ExcelMaxRows = 1048576;
            const int ExcelMaxColumns = 16384;

            using (FileStream stream = File.OpenRead(imagePath))
            using (Image image = new Image(stream))
            {
                AdjustImageSize(image, ExcelMaxRows, ExcelMaxColumns);
                GenerateExcelWorksheet(image);
            }
        }

        private static void GenerateExcelWorksheet(Image image)
        {
            var excelApplication = InitializeExcelApplication();
            var rightmostCell = GenerateExcelColumnUpperBound(image);

            Console.WriteLine("Converting image...");
            LocalWatch.Restart();

            var rowRange = GetRowRange(image, excelApplication, rightmostCell);
            var imageWidth = image.Width;

            Parallel.For(1L, image.Height + 1, i =>
            {
                Debug.WriteLine($"ROW {i} - Starting processing {DateTime.Now.ToLongTimeString()}");
                SetCellColors(image, rowRange, imageWidth, i);
                Debug.WriteLine($"ROW {i} - Finished processing {DateTime.Now.ToLongTimeString()}");
            });

            Marshal.ReleaseComObject(rowRange);
            ResizeCells(excelApplication);
            AdjustExcelWindow(excelApplication);

            LocalWatch.Stop();
        }

        private static Range GetRowRange(Image image, Application excelApplication, string rightmostCell)
        {
            Range rowRange = excelApplication?.Range["A1", rightmostCell + image.Height];

            if (rowRange == null)
            {
                throw new InvalidOperationException(
                    "Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            return rowRange;
        }

        private static void SetCellColors(Image image, Range rowRange, int imageWidth, long i)
        {
            for (var j = 1; j <= imageWidth; j++)
            {
                var rowIndex = (i - 1) * imageWidth + j;
                var cell = rowRange.Cells[rowIndex];
                var pixel = image.Pixels[rowIndex];

                try
                {
                    cell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(pixel.R, pixel.G, pixel.B));
                }
                catch (COMException ex)
                {
                    Debug.Write(ex.ToString());
                }
                Marshal.ReleaseComObject(cell);
            }

            Interlocked.Add(ref CellCount, imageWidth);

            if (CellCount > 200000)
            {
                rowRange.ClearFormats();
                rowRange.Cells.ClearFormats();

                Debug.WriteLine("Starting GC");
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Interlocked.Exchange(ref CellCount, 0);
            }

            ReportProgress(imageWidth);
        }

        private static void AdjustImageSize(Image image, int excelMaxRows, int excelMaxColumns)
        {
            int newHeigth = image.Height;
            int newWidth = image.Width;
            ImageSize = image.Width * image.Height;
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

            if (pixelCount != ImageSize)
                Console.WriteLine(
                    $"WARNING: Image pixel count does not match the calculated pixel count (H*W) - expected:{ImageSize} actual:{pixelCount}");

            // throw new ImageProcessingException($"WARNING: Image pixel count does not match the calculated pixel count (H*W) - expected:{expectedPixelCount} actual:{pixelCount}");
        }

        private static void AdjustExcelWindow(Application excelApplication)
        {
            excelApplication.Visible = true;
            excelApplication.WindowState = XlWindowState.xlMaximized;
            excelApplication.ActiveWindow.Zoom = 10;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(excelApplication.Sheets);
            Marshal.ReleaseComObject(excelApplication.Worksheets);
            Marshal.ReleaseComObject(excelApplication.Workbooks);
            Marshal.ReleaseComObject(excelApplication);
        }

        private static void ResizeCells(Application excelApplication)
        {
            ((Worksheet) excelApplication.ActiveSheet).Columns.ColumnWidth = 2;
            ((Worksheet) excelApplication.ActiveSheet).Rows.EntireColumn.RowHeight =
                ((Range) ((Worksheet) excelApplication.ActiveSheet).Cells[1]).Width;
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
            Worksheet worksheet = (Worksheet) workbook.Worksheets[1];

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

        private static void ReportProgress(int processedCells)
        {
            Debug.WriteLine($"Start reporting {DateTime.Now.ToLongTimeString()}");
            Interlocked.Add(ref ProgressStatus.ProcessedPixelCount, processedCells);
            long progressValue = Interlocked.Read(ref ProgressStatus.ProcessedPixelCount) * 100 / ImageSize;

            if (Interlocked.Read(ref ProgressStatus.PreviousProgressValue) < progressValue)
            {
                if (Interlocked.Read(ref ProgressStatus.EllapsedSeconds) != LocalWatch.Elapsed.Seconds)
                {
                    Interlocked.Exchange(ref ProgressStatus.Percentage, (int) progressValue);
                    Interlocked.Exchange(ref ProgressStatus.EllapsedMinutes, LocalWatch.Elapsed.Minutes);
                    Interlocked.Exchange(ref ProgressStatus.EllapsedSeconds, LocalWatch.Elapsed.Seconds);
                    Console.WriteLine(
                        $"\tConverting: %{(int) progressValue} (elapsed: {Interlocked.Read(ref ProgressStatus.EllapsedMinutes)}m {Interlocked.Read(ref ProgressStatus.EllapsedSeconds)}s)");
                }
                Interlocked.Exchange(ref ProgressStatus.PreviousProgressValue, (int) progressValue);
            }
            Debug.WriteLine($"Finished reporting {DateTime.Now.ToLongTimeString()}");
        }
    }
}