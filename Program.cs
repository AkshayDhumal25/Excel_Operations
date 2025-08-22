using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
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

        List<int> ages = new List<int>();
        List<DateTime> joinDates = new List<DateTime>();

        for (int row = 2; row <= rowCount; row++)
        {
            if (int.TryParse(worksheet.Cells[row, 4].Text, out int age))
                ages.Add(age);

            if (DateTime.TryParse(worksheet.Cells[row, 5].Text, out DateTime joinDate))
                joinDates.Add(joinDate);
        }

        
        var ageGroups = ages
            .GroupBy(a => a / 10) 
            .OrderBy(g => g.Key)
            .ToDictionary(g => $"{g.Key * 10}-{g.Key * 10 + 9}", g => g.Count());

        Console.WriteLine("\nAge Group Distribution:");
        foreach (var group in ageGroups)
            Console.WriteLine($"{group.Key}: {group.Value} users");

        Console.WriteLine($"Average Age: {ages.Average():F2}");

        var joinYearGroups = joinDates
            .GroupBy(d => d.Year)
            .OrderBy(g => g.Key)
            .ToDictionary(g => g.Key, g => g.Count());

        Console.WriteLine("\nJoin Year Distribution:");
        foreach (var group in joinYearGroups)
            Console.WriteLine($"{group.Key}: {group.Value} users");

        Console.WriteLine($"Earliest Join Date: {joinDates.Min():d}");
        Console.WriteLine($"Latest Join Date: {joinDates.Max():d}");
    }
}
