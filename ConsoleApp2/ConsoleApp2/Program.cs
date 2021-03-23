using OfficeOpenXml;
using System.IO;
using System;
using System.Linq;
using System.Configuration;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Data;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {

            SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\HP\source\repos\ConsoleApp2\ConsoleApp2\Database.mdf;Integrated Security=True");
            connection.Open();



            string path = "C:\\Users\\HP\\Downloads\\User.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            // get number of rows and columns in the sheet

            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            worksheet.Cells[2, 3].Value = "GizliDegil";

            string[,] array = new string[rows + 2, columns + 1];
            var differentRowColumns = new List<int>();

            // Excel tablosunu 2 boyutlu dizi içerisine yerlestirdim.

            for (int i = 2; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    object content = worksheet.Cells[i, j].Value;
                    if (worksheet.Cells[i, j].Value == null)
                    {
                        array[i - 2, j - 1] = "";
                        continue;
                    }
                    else
                    {                      
                        array[i - 2, j - 1] = worksheet.Cells[i, j].Value.ToString();
                    }
                }
            }

            for(int i = 0;i<rows; i++)
            {
                for(int j = 0; j < columns; j++)
                {
                    Console.WriteLine(array[i, j]);
                }
            }

            //Console.WriteLine(ColumnHelper<string>("4", "0", array));
            


            List<string> differentID = new List<string>();

            // Create the Command and Parameter objects.
            SqlCommand command = new SqlCommand("SELECT * FROM routertable", connection);

            // Create and execute the DataReader..

            SqlDataReader reader = command.ExecuteReader();
            int k = 0,t=0;
            int differentCounter = 0;
            while (reader.Read())
            {
                var rec = new List<string>();
                for (int i = 1; i < reader.FieldCount-1; i++) //The mathematical formula for reading the next fields must be <=
                {
                    rec.Add(reader.GetString(i).ToString());
                    if (i == 1)
                    {
                        if (!reader.GetString(i).Equals(array[k, 0]))
                        {
                            differentID.Add(array[k, 0]);
                            differentCounter++; // id'ler farklı ise degeri 1 arttır.
                            continue;
                        }                       
                    }
                    if (!reader.GetString(i).Equals(array[k, t]))
                    {
                        Console.WriteLine("Farkli olan deger: " + array[k,t]);
                        differentRowColumns.Add(k);
                        differentRowColumns.Add(t);
                    }
                    t++;
                }
                t = 0;
                k++;
            }
            Console.WriteLine("***********");
            Console.WriteLine("kaç tanesinin id karşılaştırması farklı geldi: " + differentCounter);
            Console.WriteLine("***********");
            
            Console.WriteLine("Kaç farklı hücrede değişiklik yapıldı: " + (differentRowColumns.Count / 2));

            //List<string> differentIdCount = differentRowColumns.Distinct().ToList().ToString();
            List<string> differentIdCount = new List<string>();


            // Hangi satırda değişik yapıldıysa o satırların id'sini verir.
            for (int i = 0; i < differentRowColumns.Count; i += 2)
            {
                Console.WriteLine("different row and columns id: "
                + ColumnHelper<string>(differentRowColumns[i].ToString(), "0", array));
                differentIdCount.Add(ColumnHelper<string>(differentRowColumns[i].ToString(), "0", array));
            }


            // Aynı id üzerinde birden fazla hücrede işlem yapıldıysa, aynı id'yi 3-4 defa dönme 1 kere dön.
            List<string> sayacArray = differentIdCount.Distinct().ToList();
            Console.WriteLine("sayac array count: " + sayacArray.Count);

            // Id'ler içinde router: şeklinde başlayan ifadeyi yok etmek için buradaki for ile düzenledim.
            for(int j = 0; j<sayacArray.Count; j++)
            {
                sayacArray[j] = sayacArray[0].Substring(sayacArray[0].IndexOf(":") + 1);
            }
            Console.WriteLine("Yeni deger: " + sayacArray[0]);


            


            //for (int i = 0; i <= 499; i++)
            //{
            //    try
            //    {
            //        //var command1 = new SqlCommand("INSERT INTO routertable(tabloid,adi,gizliliksinifi,saklamayeri,butunluk,erisilebilirlik,gizlilik,varlikdegeri,konum,varlikmuhafizi,varliksahibi,ortam,sirket) Values('" + array[i, 0] + "','" + array[i, 1] + "','" + array[i, 2] + "','" + array[i, 3] + "','" + array[i, 4] + "','" + array[i, 5] + "','" + array[i, 6] + "','" + array[i, 7] + "','" + array[i, 8] + "','" + array[i, 9] + "','" + array[i, 10] + "','" + array[i, 11] + "','" + array[i, 12] + "')", connection);
            //        //command1.ExecuteNonQuery();


            //    }
            //    catch (SqlException e)
            //    {
            //        throw;
            //    }
            //}
        }

        public static T ColumnHelper<T>(T rowIndex, T columnName, T[,] array)
        {
            var rowIndexx = Convert.ToInt32(rowIndex.ToString());
            var columnNamee = Convert.ToInt32(columnName.ToString());
            return array[rowIndexx,columnNamee];
        }
    }
}


