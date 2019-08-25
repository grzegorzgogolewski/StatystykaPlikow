using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace StatystykaPlikow
{
    class Program
    {
        static void Main()
        {
            string startDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ;
            
            FileInfo xlsFile = new FileInfo(Path.Combine(startDirectory ?? throw new InvalidOperationException(), "statystyka.xlsx"));

            File.Delete(xlsFile.Name);

            ExcelPackage xlsWorkbook = new ExcelPackage(xlsFile);
            xlsWorkbook.Workbook.Properties.Title = @"Statystyka plików w folderach";
            xlsWorkbook.Workbook.Properties.Author = @"Grzegorz Gogolewski";
            xlsWorkbook.Workbook.Properties.Company = @"GISNET";

            ExcelWorksheet xlsSheet = xlsWorkbook.Workbook.Worksheets.Add("statystyka");

            xlsSheet.Cells[1, 1].Value = @"KATALOG";
            xlsSheet.Cells[1, 2].Value = @"LICZBA PLIKÓW";

            string[] directories = Directory.GetDirectories(startDirectory, "????_*", SearchOption.TopDirectoryOnly);

            for (int i = 0; i < directories.Length; i++)
            {
                Console.WriteLine(directories[i]);
                xlsSheet.Cells[i + 2, 1].Value = directories[i].Split(Path.DirectorySeparatorChar).Last();
                xlsSheet.Cells[i + 2, 2].Value = Directory.GetFiles(directories[i], "*.pdf", SearchOption.TopDirectoryOnly).Length;
            }

            xlsSheet.View.FreezePanes(2, 1);
            xlsSheet.Cells.Style.Font.Size = 10;
            xlsSheet.Cells.AutoFitColumns(0);
            xlsWorkbook.Save();

            Console.WriteLine("Gotowe");
            Console.ReadKey();
        }
    }
}
