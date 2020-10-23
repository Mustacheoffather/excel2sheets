using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
//using System.Timers;
//using System.Threading.Tasks;

namespace excel2sheets
{
    class Program
    { 
        private static System.Timers.Timer aTimer;
        static void Main(string[] args)
        {
            Spreadsheet sSheet = new Spreadsheet();
            Console.WriteLine("Введите путь к файлу и название самого файла через слэш");
            string file = Console.ReadLine();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fi = new FileInfo(file);

            Dictionary<string, List<string>> excelData = new Dictionary<string, List<string>>();

            byte[] bin = File.ReadAllBytes(fi.ToString());

            ExcelPackage excelPackage1 = new ExcelPackage(new FileInfo(fi.ToString()));
            var myWorksheet = excelPackage1.Workbook.Worksheets.First();
            int totalRows = myWorksheet.Dimension.End.Row+1;
            int totalColumns = myWorksheet.Dimension.End.Column+1;
            string[,] cellsVal = new string[totalRows, totalColumns];

            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                var worksheet = excelPackage.Workbook.Worksheets.First(); //select sheet here
                for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                {
                    //Console.WriteLine($"here \t{i}");
                    for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                    {
                        if (worksheet.Cells[i, j].Value == null)
                        {
                            worksheet.Cells[i, j].Value = " ";
                        }
                        else
                        {
                            cellsVal[i, j] = worksheet.Cells[i, j].Value.ToString();
                            if (worksheet.Cells[i, j].Value.ToString() == null)
                            {
                                Console.WriteLine("null");
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Введите ID вашей таблицы");
            sSheet.SpreadSheetId = Console.ReadLine();

            Console.WriteLine("Get Creds");
            var credentials = sSheet.GetCredential();

            Console.WriteLine("Get service");
            var service = sSheet.GetService(credentials);

            Console.WriteLine("Fill Data");
            sSheet.FillTheSpreadSheet(service, sSheet.SpreadSheetId, cellsVal);

            Console.WriteLine("Done");
            
            Console.WriteLine(
                "Введите директорию в которой хотите создать новый файл и название нового файла через слэш и укажите расширение файла xlsx");

            string Directory = Console.ReadLine();
            
            /*Timer timer = new Timer(5000);
            timer.Elapsed += async ( sender, e ) => await UpdateNewFile();
            timer.Start();
            Console.Write("Press any key to exit... ");
            Console.ReadKey();
            timer.Stop();*/
            
            //FileInfo fillFile = new FileInfo(Directory ?? string.Empty);
            FillNewFile(cellsVal, Directory, totalRows, totalColumns);
            Console.WriteLine("File was created");
        }
        private static void FillNewFile(string[,] data, string Directory, int totalRows, int totalColumns)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "User";
                excelPackage.Workbook.Properties.Title = " ";
                excelPackage.Workbook.Properties.Subject = " ";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                var worksheeet = excelPackage.Workbook.Worksheets.First();
                for (int i = 1; i < totalRows; i++)
                {
                    for (int j = 1; j < totalColumns; j++)
                    {
                        worksheet.Cells[i, j].Value = data[i,j];
                    }
                }
                
                FileInfo fi = new FileInfo(Directory);
                excelPackage.SaveAs(fi);
            }
        }
    }
}