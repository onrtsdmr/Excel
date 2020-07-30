using System;
using ReadExcel.Extensions;

namespace ReadExcel
{
    internal class Program
    {
        private const string DummyExcelFilePath =
            @"C:\Users\onurt\Desktop\.NET WEB API\ReadExcel\ReadExcel\Data\DummyExcelFile.xlsx";

        private static void Main(string[] args)
        {
            DummyExcelFilePath.ReadExcel();
            Console.ReadKey();
        }
    }
}