using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelValidation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            HashSet<string> emailHashSet = new HashSet<string>();
            string[] result = new string[11];
            Excel baseFile = new Excel(@"D:\DotNet\ExcelValidation\ExcelFile2.xlsx", 1);
            Excel newFile = new Excel(@"D:\DotNet\ExcelValidation\Test1.xlsx", 1);

            for (int i = 1; i <= 1; i++)
            {
                for (int j = 2; j < 10; j++)
                {
                    result[j] = baseFile.ReadCellsFromColumn(j, i);
                    if (emailHashSet.Contains(result[j]))
                    {
                        newFile.WriteInCell(j, i, result[j]);

                        newFile.WriteInCell(j, i+2, "duplicate");
                    }
                    else
                    {
                        emailHashSet.Add(result[j]);
                        newFile.WriteInCell(j,i, result[j]);
                        newFile.WriteInCell(j, i+2, "valid");
                    }
                }
            }
            newFile.SaveWorkBook();
            baseFile.CloseFile();
            foreach (var email in emailHashSet)
            {
                Console.WriteLine(email);
            }
            Console.ReadLine();
        }
    }
}
