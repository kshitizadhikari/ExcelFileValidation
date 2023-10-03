using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExcelValidation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            HashSet<string> dbEmailHashSet = new HashSet<string>();
            List<string> excelEmailList = new List<string>();
            Excel baseFile = new Excel(@"D:\DotNet\ExcelValidation\abc.xlsx", 1);

            string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Server=DESKTOP-72AIJKL;Initial Catalog=EmailCheckDb;Integrated Security=True";
            SqlConnection connection = new SqlConnection(connectionString);

            try
            {
                connection.Open();
                Console.WriteLine("Connection established\n");

                string selectQuery = "SELECT email FROM Users"; // Adjust your SQL query as needed
                SqlCommand command = new SqlCommand(selectQuery, connection);
                SqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string email = reader["email"].ToString();
                    dbEmailHashSet.Add(email);
                }
                reader.Close();

                for (int i = 1; i <= 1; i++)
                {
                    for (int j = 2; j < 10; j++)
                    {
                        string cellValue = baseFile.ReadCellsFromColumn(j, i).ToString();
                        if (dbEmailHashSet.Contains(cellValue))
                        {
                            baseFile.WriteInCell(j, i + 2, "duplicate");
                            excelEmailList.Add(cellValue);
                        }
                    }
                }

                baseFile.SaveWorkBook();

                Console.WriteLine("Emails in Database :");
                foreach (var email in dbEmailHashSet)
                {
                    Console.WriteLine(email);
                }

                Console.WriteLine("\nDuplicate Emails in Excel File :");
                foreach (var email in excelEmailList)
                {
                    Console.WriteLine(email);
                }

            }
            catch (SqlException ex)
            {
                Console.WriteLine("Error connecting to the database. \n" + ex.Message);
            }
            finally
            {
                connection.Close();
            }

            Console.ReadLine();
        }
    }
}
