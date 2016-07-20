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
using Microsoft.Win32;
using System.Windows.Forms;
using System.Threading;

namespace LogTMUpdate
{
    class Program
    {
        static void Main(string[] args)
        {


            if (args.Length == 0)
            {
                //if (Registry.ClassesRoot.GetValue("HKEY_CLASSES_ROOT\\batfile\\shell\\PopulateFoldersWithXLZ\\command", null) == null)
                if (Registry.GetValue("HKEY_CLASSES_ROOT\\Directory\\shell\\UpdateTMLog\\command", "", null) == null)
                {
                    Registry.SetValue("HKEY_CLASSES_ROOT\\Directory\\shell\\UpdateTMLog\\command", "", System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName + " \"%1\"");
                    MessageBox.Show("Script added to directory context menu!", "Script Added!");
                }
                else
                {
                    MessageBox.Show("Script already installed!", "Script detected!");
                }

            }
            else
            {
                string dir = Path.GetFullPath(args[0]);
                UpdateTMLog(dir);

            }
            Thread.Sleep(2000);

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

        static void UpdateTMLog(string arg)
        {

            // otwieranie TM update log xlsx i dopisywanie wierszy do odpowiednich zakladek,
            // na podstawie sciezki z ktore odpalono skrypt pod PPM
            FileInfo TMlogFile = new FileInfo(@"\\waw-fs01\K_ENG\_Gothenburg\Volvo\_TM_update_log\TM_UPDATE_LOG.xlsx");


            string projectPath = arg;


            string projectNumberPattern = @"(?<=\\)\d{7}(?=_)";
            string HOPattern = @"(?<=\\HO)\d+?(?=\\)";


            Regex findProjectNumber = new Regex(projectNumberPattern, RegexOptions.IgnoreCase);
            Match matchedProjectNumber = findProjectNumber.Match(projectPath);
            int projNr = Int32.Parse(matchedProjectNumber.ToString());

            Regex findHONumber = new Regex(HOPattern, RegexOptions.IgnoreCase);
            Match matchedHONumber = findHONumber.Match(projectPath);
            int hoNr = Int32.Parse(matchedHONumber.ToString());

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
            Console.WriteLine("Number of languages: " + liczbaJezykow);

            Console.WriteLine("Which client? (Volvo, Thule, ...)");
            string Client = Console.ReadLine();
            Console.WriteLine("Which TM? (Trucks, Buses, ... [none])");
            string TM = Console.ReadLine();
            Console.WriteLine("TM status? (Proofread, ClientApproved)");
            string TMStatus = Console.ReadLine();
            //Console.WriteLine("Additional info? (Comments)");
            //string AddInfo = Console.ReadLine();

            int dodanychWpisow = 0;
            using (ExcelPackage package = new ExcelPackage(TMlogFile))
            {
                //ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                foreach (var worksheet in package.Workbook.Worksheets)
                {

                    if (worksheet.Name.ToLower() == Client.ToLower())
                    {
                        int lastNotEmptyRow = GetLastUsedRow(worksheet);

                        int emptyRow = lastNotEmptyRow + 1;
                        //int emptyLangRow = emptyRow;
                        foreach (string langDir in Directory.GetDirectories(projectPath, "*-*"))
                        {

                            string langFolder = langDir.Substring(langDir.LastIndexOf('\\') + 1);

                            worksheet.Cells["A:B"].Style.Numberformat.Format = "";
                      

                            worksheet.Cells[emptyRow, 1].Value = projNr;
                            worksheet.Cells[emptyRow, 2].Value = hoNr;
                            worksheet.Cells[emptyRow, 3].Value = langFolder;
                            worksheet.Cells[emptyRow, 4].Value = DateTime.Now.ToShortDateString();
                            worksheet.Cells[emptyRow, 5].Value = Client;
                            worksheet.Cells[emptyRow, 6].Value = TM;
                            worksheet.Cells[emptyRow, 7].Value = TMStatus;
                            worksheet.Cells[emptyRow, 8].Value = Environment.UserName;
                            emptyRow++;
                            dodanychWpisow++;
                        }


                        //Console.WriteLine(lastRow);
                        //Console.WriteLine(worksheet);
                    }

                }


                package.Save();
            }

            Console.WriteLine("Dodane wpisy: " + dodanychWpisow);
         


        }
    }
}