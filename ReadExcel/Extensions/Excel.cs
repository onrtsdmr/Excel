using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader;
using ReadExcel.Models;

namespace ReadExcel.Extensions
{
    public static class Excel
    {
        private static readonly List<DummyModel> Data = new List<DummyModel>();

        public static void ReadExcel(this string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                            Data.Add(new DummyModel
                            {
                                Name = reader.GetString(1),
                                Price = reader.GetString(2),
                                Quantity = reader.GetString(3)
                            });
                    } while (reader.NextResult());
                }
            }

            foreach (var item in Data)
            {
                Console.Write(item.Name);
                Console.Write("   |   ");
                Console.Write(item.Price);
                Console.Write("   |   ");
                Console.Write(item.Quantity);
                Console.Write("   |   \n");
            }
        }
    }
}