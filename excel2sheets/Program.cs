using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using OfficeOpenXml;
using System.Text;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Dynamic;
using Newtonsoft.Json;
using Dropbox.Api;
using Dropbox.Api.Files;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace excel2sheets
{
    class Program
    {
        private static string ClientSecret = "credentials.json";
        private static string[] Scopes = {SheetsService.Scope.Spreadsheets};
        private static string ApplicationName = "excel2sheets";
        private static string SpreadSheetId;
        private const string Range = "'Лист1'!A1:A";
        static int k = 0;

        static void Main(string[] args)
        {

            Console.WriteLine("Введите путь к файлу и название самого файла через слэш");
            string file = Console.ReadLine();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fi = new FileInfo(file);
            
            List<string> excelData = new List<string>();

            byte[] bin = File.ReadAllBytes(fi.ToString());

            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            if (worksheet.Cells[i,j].Value != null)
                            {
                                excelData.Add(worksheet.Cells[i,j].Value.ToString());
                            }
                            k++;
                        }
                    }
                }
                
            }

            Console.WriteLine("Введите ID вашей таблицы");
            SpreadSheetId = Console.ReadLine();

            Console.WriteLine("Get Creds");
            var credentials = GetCredential();

            Console.WriteLine("Get service");
            var service = GetService(credentials);

            Console.WriteLine("Fill Data");
            FillSpreadSheet(service, SpreadSheetId, excelData, k);

            Console.WriteLine("Getting result");
            string result = GetFirstCell(service, Range, SpreadSheetId);
            //Console.WriteLine("result: {0}", result);

            Console.WriteLine("Done");
            //Console.ReadLine();

            Console.WriteLine("Введите директорию в которой хотите создать новый файл и название нового файла через слэш");
            string Directory = Console.ReadLine();
            SpreadsheetDocument spreadsheetDocument =
            SpreadsheetDocument.Create(Directory, SpreadsheetDocumentType.Workbook);

            /*Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);*/

            FillNewFile(Directory, excelData);


        }

        public static UserCredential GetCredential()
        {
            using (var stream = new FileStream(ClientSecret, FileMode.Open, FileAccess.Read))
            {
                var credPath = Path.Combine(Directory.GetCurrentDirectory(), "sheetsCreds.json");

                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)
                ).Result;
            }
        }

        public static SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
        }

        private static void FillSpreadSheet(SheetsService service, string SpreadSheetId, List<string> sb, int k)
        {
            List<Request> requests = new List<Request>();
            //Console.WriteLine(sb.Length);
            List<CellData> values = new List<CellData>();
            foreach (var i in sb) 
            {
                values.Add
                (
                    new CellData
                    {
                        UserEnteredValue = new ExtendedValue
                        {
                            StringValue = i
                        }
                    }
                );
            }
                    
            for(int i=0; i < k; i++)
            {
                requests.Add(
                    new Request
                    {
                        UpdateCells = new UpdateCellsRequest
                        {
                            Start = new GridCoordinate
                            {
                                SheetId = 0,
                                RowIndex = i,
                                ColumnIndex = 0
                            },

                            Rows = new List<RowData> { 
                                new RowData { 
                                    Values= values
                                } 
                            },
                            Fields = "UserEnteredValue"
                        }
                    }
                );
                Console.WriteLine(requests);
            }

            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest();

            busr.Requests = requests;

            service.Spreadsheets.BatchUpdate(busr, SpreadSheetId).Execute();
        }

        private static string GetFirstCell(SheetsService service, string range, string SpreadSheetId)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SpreadSheetId, range);
            ValueRange response = request.Execute();

            string result = null;

            foreach (var value in response.Values)
            {
                result += " " + value[0];
            }

            return result;
        }

        private static void FillNewFile(string filePath, List<string> sb)
        {
            /*List<CellData> values = new List<CellData>();
            foreach (var i in sb)
            {
                values.Add
                (
                    new CellData
                    {
                        UserEnteredValue = new ExtendedValue
                        {
                            StringValue = i
                        }
                    }
                );
            }*/

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //create a WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                //add all the content from the List<Book> collection, starting at cell A1
                worksheet.Cells["A1:Z"].LoadFromCollection(sb);
            }

        }
    }
}