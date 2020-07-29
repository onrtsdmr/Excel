using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using ReadExcel.Models;

namespace ReadExcel
{
    public static class Excel
    {
        public static List<DummyModel> Data = new List<DummyModel>();

        public static void ReadExcel(this string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                            Data.Add(new DummyModel()
                            {
                                Name = reader.GetString(1),
                                Price = reader.GetString(2),
                                Quantity = reader.GetString(3),
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

/*
            IExcelDataReader excelReader;
            List liste = new List();
            int counter = 0;
 
            /*yeni sürümlerde bu kaldırıldığı için kapatıldı.
            //Datasete atarken ilk satırın başlık olacağını belirtiyor.
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();*/

//Veriler okunmaya başlıyor.
/*
while (excelReader.Read())
{
    counter++;

    //ilk satır başlık olduğu için 2.satırdan okumaya başlıyorum.
    if (counter > 1)
    {
        liste.Add("Ad =" + excelReader.GetString(0));
        liste.Add("Soyad =" + excelReader.GetString(1));
    }
}

//Okuma bitiriliyor.
excelReader.Close();

//Veriler ekrana basılıyor.
foreach (var item in liste)
{
    Console.WriteLine(item);
}

Console.ReadKey();
}
}
}*/