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
            int count = 0;
            int emailColumn = 1;
            //Excel baseFile = new Excel(@"D:\DotNet\ExcelValidation\abc.xlsx", 1);
            char choice;
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
                    count++;
                    dbEmailHashSet.Add(email);
                }
                reader.Close();

                Console.WriteLine("Select the appropriate file format : \n");
                Console.WriteLine("Excel File => 1");
                Console.WriteLine("CSV File => 2");
                choice = Console.ReadKey().KeyChar;
                Console.WriteLine("\nPress Enter");
                Console.ReadLine();

                switch (choice)
                {
                    case '1':

                        Console.WriteLine("Please enter the path or the filename that you want to use : ");
                        string filePath = Console.ReadLine();

                        if (System.IO.File.Exists(filePath))
                        {
                            Console.WriteLine($"Selected File => {filePath}");
                            Excel baseFile = new Excel(filePath, 1);
                            for (int i = 1; i <= count; i++)
                            {

                                string cellValue = baseFile.ReadCellsFromColumn(i, emailColumn);
                                if (dbEmailHashSet.Contains(cellValue))
                                {
                                    baseFile.WriteInCell(i, emailColumn + 2, "duplicate");
                                    excelEmailList.Add(cellValue);
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
                        else
                        {
                            Console.WriteLine("The specified file path does not exist.");
                        }

                        Console.ReadLine();
                        break;

                    case '2':
                        break;

                    default:
                        Console.WriteLine("Invalid Choice");
                        break;
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
