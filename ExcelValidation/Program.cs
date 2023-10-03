using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;

namespace ExcelValidation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            HashSet<string> dbEmailHashSet = new HashSet<string>();
            List<string> duplicateEmailList = new List<string>();
            int count = 0;
            int emailColumn = 1;
            char choice;
            string filePath;
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
                        Console.WriteLine("Please enter the absolute path along with the filename and extension that you want to use : ");
                        filePath = Console.ReadLine();

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
                                    duplicateEmailList.Add(cellValue);
                                }
                            }

                            baseFile.SaveWorkBook();
                           
                            Console.WriteLine("\nDuplicate Emails in the Excel File are:");
                            foreach (var email in duplicateEmailList)
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
                        List<string> emailsInCsvFile = new List<string>();

                        Console.WriteLine("Please enter the absolute path along with the filename and extension that you want to use : ");
                        filePath = Console.ReadLine();
                        string newLine;
                        if (System.IO.File.Exists(filePath))
                        {
                            using (StreamReader readerA = new StreamReader(filePath))
                            {
                                // Create a StreamWriter to write changes back to the file
                                using (StreamWriter writer = new StreamWriter("temp.csv"))
                                {
                                    while (!readerA.EndOfStream)
                                    {
                                        string line = readerA.ReadLine();
                                        string[] values = line.Split(',');

                                        if (values.Length > 0)
                                        {
                                            string email = values[0];
                                            if (dbEmailHashSet.Contains(email))
                                            {
                                                duplicateEmailList.Add(email);
                                                values[2] = "duplicate";
                                                newLine = string.Join(",", values); //new line out of bounds of array
                                                writer.WriteLine(newLine);
                                            }
                                            else
                                            {
                                                newLine = $"{email},{values[1]}";
                                                writer.WriteLine(newLine);
                                            }
                                            emailsInCsvFile.Add(email);
                                        }
                                    }
                                }
                            }

                            File.Delete(filePath);
                            File.Move("temp.csv", filePath);

                            Console.WriteLine("\nDuplicate emails in the CSV file are :");
                            foreach (string email in duplicateEmailList)
                            {
                                Console.WriteLine(email);
                            }
                        }
                        else
                        {
                            Console.WriteLine("The specified file path does not exist.");
                        }



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
