using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
using System.Linq;

namespace LogTMUpdate
{
    class Program
    {
        static void Main(string[] args)
        {

            // otwieranie TM update log xlsx i dopisywanie wierszy do odpowiednich zakladek,
            // na podstawie sciezki z ktore odpalono skrypt pod PPM
            FileInfo TMlogFile = new FileInfo(@"C:\Users\zawii\Desktop\ttest\TmUpdate.xlsx");


            string projectPath = @"D:\_work\CPH\Desktop\1624444_Got_Volvo_SDAS_DASDAS_DAS\HO9\post\_rdy";


            string projectNumberPattern = @"(?<=\\)\d{7}(?=_)";
            string HOPattern = @"(?<=\\HO)\d+?(?=\\)";


            Regex findProjectNumber = new Regex(projectNumberPattern, RegexOptions.IgnoreCase);
            Match matchedProjectNumber = findProjectNumber.Match(projectPath);

            Regex findHONumber = new Regex(HOPattern, RegexOptions.IgnoreCase);
            Match matchedHONumber = findHONumber.Match(projectPath);


            Console.WriteLine("Project number: " + matchedProjectNumber);
            Console.WriteLine("HO: " + matchedHONumber);

            Console.WriteLine("Languages: ");
            int liczbaJezykow = 0;
            foreach (string langDir in Directory.GetDirectories(projectPath, "*-*"))
            {
                string langFolder = langDir.Substring(langDir.LastIndexOf('\\') + 1);
                Console.WriteLine(langFolder);
                liczbaJezykow++;
            }
            Console.WriteLine("Number of languages: " +liczbaJezykow);

            Console.WriteLine("Which client? (Volvo, Thule, ...)");
            string Client = Console.ReadLine();
            Console.WriteLine("Which TM? (Trucks, Buses, ... [none])");
            string TM = Console.ReadLine();
            Console.WriteLine("TM status? (Proofread, ClientApproved)");
            string TMStatus = Console.ReadLine();
            Console.WriteLine("Additional info? (Comments)");
            string AddInfo = Console.ReadLine();


            using (ExcelPackage package = new ExcelPackage(TMlogFile))
            {
                //ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                foreach (var worksheet in package.Workbook.Worksheets)
                {

                    if (worksheet.Name == Client)
                    {
                        int lastNotEmptyRow = GetLastUsedRow(worksheet);

                        int emptyRow = lastNotEmptyRow + 1;

                        worksheet.Cells[emptyRow, 1].Value = matchedProjectNumber;
                        worksheet.Cells[emptyRow, 2].Value = matchedHONumber;

                        worksheet.Cells[emptyRow, 4].Value = DateTime.Now.ToShortDateString();
                        worksheet.Cells[emptyRow, 5].Value = Client;
                        worksheet.Cells[emptyRow, 6].Value = TM;
                        worksheet.Cells[emptyRow, 7].Value = TMStatus;
                        worksheet.Cells[emptyRow, 8].Value = Environment.UserName;



                        //Console.WriteLine(lastRow);
                        //Console.WriteLine(worksheet);
                    }

                }


                package.Save();
            }

            
            Console.ReadKey();

            

        }

        static int GetLastUsedRow(ExcelWorksheet sheet)
        {
            var row = sheet.Dimension.End.Row;
            while (row >= 1)
            {
                var range = sheet.Cells[row, 1, row, sheet.Dimension.End.Column];
                if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
                {
                    break;
                }
                row--;
            }
            return row;
        }
    }
}
