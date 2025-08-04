using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string filePath = @"C:\Users\Dell\Downloads\SampleData100.xlsx";

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];

        int rowCount = worksheet.Dimension.Rows;
        int colCount = worksheet.Dimension.Columns;

        Console.WriteLine("Reading Excel Data:");
        for (int row = 2; row <= rowCount; row++) 
        {
            Console.WriteLine($"Id: {worksheet.Cells[row, 1].Text}, " +
                              $"Name: {worksheet.Cells[row, 2].Text}, " +
                              $"Email: {worksheet.Cells[row, 3].Text}, " +
                              $"Age: {worksheet.Cells[row, 4].Text}, " +
                              $"JoinDate: {worksheet.Cells[row, 5].Text}");
        }

        Console.WriteLine($"\nTotal number of data rows: {rowCount - 1}");
    }
}
