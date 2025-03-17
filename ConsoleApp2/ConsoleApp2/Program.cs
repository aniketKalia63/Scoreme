using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

class Program
{
    static void Main()
    {
        string pdfPath = @"C:\Users\kalia\Downloads\test3.pdf"; 
        string excelPath = @"D:\output.xlsx";  
        List<List<string>> extractedTable = new List<List<string>>();

        using (PdfDocument pdf = PdfDocument.Open(pdfPath))
        {
            foreach (Page page in pdf.GetPages())
            {
                Console.WriteLine(page.Text);
                extractedTable.AddRange(ExtractTables(page));
            }
        }
        
        SaveToExcel(extractedTable, excelPath);
        Console.WriteLine(" Extraction complete! Data saved to Excel.");
       
    }

    static List<List<string>> ExtractTables(Page page)
    {
        var words = page.GetWords().OrderBy(w => w.BoundingBox.Bottom).ToList();
        if (!words.Any()) return new List<List<string>>();

        var table = new List<List<string>>();
        var columns = DetectColumns(words);
        if (columns.Count == 0) return new List<List<string>>();

        List<string> row = Enumerable.Repeat("", columns.Count).ToList();
        double prevY = words.First().BoundingBox.Bottom;

        foreach (var word in words)
        {
            if (Math.Abs(word.BoundingBox.Bottom - prevY) > 5) // New row
            {
                if (row.Count > 0) table.Add(row);
                row = Enumerable.Repeat("", columns.Count).ToList();
            }

            int colIndex = columns.FindIndex(x => word.BoundingBox.Left >= x - 5 && word.BoundingBox.Left <= x + 5);
            if (colIndex != -1 && colIndex < row.Count)
                row[colIndex] = word.Text;

            prevY = word.BoundingBox.Bottom;
        }

        if (row.Count > 0) table.Add(row);
        return table;
    }


    static List<double> DetectColumns(List<Word> words)
    {
        return words.GroupBy(w => w.BoundingBox.Left)
                    .Select(g => g.Key)
                    .OrderBy(x => x)
                    .ToList();
    }

    static void SaveToExcel(List<List<string>> table, string outputFile)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        using (ExcelPackage package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Transactions");

            for (int row = 0; row < table.Count; row++)
            {
                for (int col = 0; col < table[row].Count; col++)
                {
                    worksheet.Cells[row + 1, col + 1].Value = table[row][col];
                }
            }

            package.SaveAs(new FileInfo(outputFile));
        }
    }
}
