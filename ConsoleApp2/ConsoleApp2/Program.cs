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



            string path = "C:\\Users\\HP\\Downloads\\Routerr.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            // get number of rows and columns in the sheet

            int rows = worksheet.Dimension.Rows; 
            int columns = worksheet.Dimension.Columns;

            worksheet.Cells[2, 3].Value = "GizliDegil";

            string[,] array = new string[rows+2, columns+1];
            var differentRowColumns = new List<int>();


            string kontrol;
            // loop through the worksheet rows and columns

            for (int i = 2; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                   
                        object content = worksheet.Cells[i, j].Value;
                        if(worksheet.Cells[i, j].Value == null)
                        {
                            array[i - 2, j - 1] = "";
                            continue;
                        }
                        else
                        {
                            //Console.WriteLine("content: " + content);
                            array[i - 2, j - 1] = worksheet.Cells[i, j].Value.ToString();
                        }
                                         
                }
            }

            List<List<String>> ResultSet = new List<List<String>>();


            // Create the Command and Parameter objects.
            SqlCommand command = new SqlCommand("SELECT * FROM routertable", connection);

            // Create and execute the DataReader..

            SqlDataReader reader = command.ExecuteReader();
            int k = 0,t=0;
            while (reader.Read())
            {
                var rec = new List<string>();
                for (int i = 1; i < reader.FieldCount-1; i++) //The mathematical formula for reading the next fields must be <=
                {
                    rec.Add(reader.GetString(i).ToString());
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
            foreach(var i in differentRowColumns)
            {
                Console.WriteLine(i);
            }


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
    }
}


