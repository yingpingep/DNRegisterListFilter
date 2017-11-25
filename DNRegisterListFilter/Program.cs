using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Add by EP.
using Excel = Microsoft.Office.Interop.Excel;

namespace DNRegisterListFilter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Please enter your file path(should be *.xlsx): ");
            string filePath = Console.ReadLine();
            string outputPath = filePath.Substring(0, filePath.Length - 5) + ".txt";

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCounts = xlRange.Rows.Count;
            int colCounts = xlRange.Columns.Count;

            Console.WriteLine("Strat!");
            List<Profile> listProfiles = new List<Profile>();
            StreamWriter sw = new StreamWriter(File.Create(outputPath));
            for (int r = 2; r <= rowCounts; r++)
            {
                int index = listProfiles.FindIndex(x => x.Name == xlRange.Cells[r, 1].Value2.ToString());
                if (index != -1)
                {
                    listProfiles[index].Count += 1;
                }
                else
                {
                    Profile p = new Profile();
                    p.Name = xlRange.Cells[r, 1].Value2.ToString();
                    p.Email = xlRange.Cells[r, 2].Value2.ToString();
                    p.School = xlRange.Cells[r, 3].Value2.ToString();
                    p.Major = xlRange.Cells[r, 4].Value2.ToString();
                    p.Grade = xlRange.Cells[r, 5].Value2.ToString();
                    listProfiles.Add(p);
                }
                Console.WriteLine("Done " + r);
            }

            xlWorkbook.Close();
            xlApp.Quit();

            foreach (var p in listProfiles)
            {
                sw.WriteLine(p.Name + "\t" + p.Email + "\t" + p.School + "\t" + p.Major + "\t" + p.Grade + "\t" + p.Count);
            }

            sw.Close();
            Console.WriteLine("Done");
        }
    }
}
