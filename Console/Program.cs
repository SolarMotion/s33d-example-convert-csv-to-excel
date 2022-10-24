using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Console;
using Microsoft.Office.Interop.Excel;

namespace Console
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteLine("Start");
            var inputFilePath = @"C:\Users\weichienyap\Desktop\sample.csv";
            var outputFolderPath = @"C:\Users\weichienyap\Desktop\XXX";

            var destinationFilePath = CopyFile(inputFilePath, outputFolderPath);

            ConvertCsvToExcel(destinationFilePath);

            Write("Finish");
            Read();
        }

        static string CopyFile(string inputPath, string outputFolderPath)
        {
            Directory.CreateDirectory(outputFolderPath);

            var originalFileName = Path.GetFileName(inputPath);
            var destinationFilePath = Path.Combine(outputFolderPath, originalFileName);
            File.Copy(inputPath, destinationFilePath, true);

            return destinationFilePath;
        }

        static void ConvertCsvToExcel(string inputPath)
        {
            var csvFileName = Path.GetFileName(inputPath);
            var csvFileFolderPath = Path.GetDirectoryName(inputPath);

            var recordLines = ReadCsvFile(inputPath);

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workBook = excel.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet sheet = workBook.ActiveSheet;

            for (int i = 0; i < recordLines.Count(); i++)
            {
                var CsvLine = recordLines[i];
                for (int j = 0; j < CsvLine.Count(); j++)
                {
                    sheet.Cells[i + 1, j + 1] = CsvLine[j];
                }
            }

            File.Delete(@"C:\Users\weichienyap\Desktop\XXX\sample.xls");
            workBook.SaveAs(@"C:\Users\weichienyap\Desktop\XXX\sample.xls");
            workBook.Close();
        }

        static List<List<string>> ReadCsvFile(string inputPath)
        {
            var results = new List<List<string>>();

            var fileLines = File.ReadLines(inputPath);

            foreach (string line in fileLines)
            {
                if (line == null || line == "")
                    continue;

                results.Add(line.Split(',').ToList());
            }

            return results;
        }
    }
}
