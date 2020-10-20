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
        private static readonly string[] Scopes = {SheetsService.Scope.Spreadsheets};
        private static readonly string ApplicationName = "excel2sheets";
        private static string SpreadSheetId;
        //private const string Range = "'Sheet1'!B1:B";
        //static int k = 0;
        //static int a = 0;

        static void Main(string[] args)
        {
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
            //Console.WriteLine(totalRows);
            //Console.WriteLine(totalColumns);

            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                var worksheet = excelPackage.Workbook.Worksheets.First(); //select sheet here

                //foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                //{

                    //Console.WriteLine("here3");
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //if (worksheet.Row(i)==null)
                        //{
                          //  i++;
                        //}
                        //Console.WriteLine("here2");
                        //a = worksheet.Dimension.End.Row;
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {

                            //Console.WriteLine("here1");
                            if (worksheet.Cells[i, j].Value == null)
                            {
                                worksheet.Cells[i, j].Value = " ";
                                //Console.WriteLine("here");
                                //excelData.Add(String.Concat(i.ToString(), j.ToString()),
                                  //  new List<string> {worksheet.Cells[i, j].Value.ToString()});
                                //Console.WriteLine(worksheet.Cells[i, j].Value);
                                //var cell = w;
                                

                                //Console.WriteLine(cellsVal[i,j]);
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
                //}

                //Console.WriteLine("The data is done");
            }

            Console.WriteLine("Введите ID вашей таблицы");
            SpreadSheetId = Console.ReadLine();

            Console.WriteLine("Get Creds");
            var credentials = GetCredential();

            Console.WriteLine("Get service");
            var service = GetService(credentials);

            Console.WriteLine("Fill Data");
            FillTheSpreadSheet(service, SpreadSheetId, cellsVal);

            //Console.WriteLine("Getting result");
            //string result = GetFirstCell(service, Range, SpreadSheetId);
            //Console.WriteLine("result: {0}", result);

            Console.WriteLine("Done");
            //Console.ReadLine();

            Console.WriteLine(
                "Введите директорию в которой хотите создать новый файл и название нового файла через слэш");
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

            FillNewFile(excelData);
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

        /*private static void FillSpreadSheet(SheetsService service, string SpreadSheetId, /*Dictionary<string, List<string>> sb, int a, int k)
        {
            List<Request> requests = new List<Request>();
            //Console.WriteLine(sb.Length);
            List<CellData> values = new List<CellData>();
            foreach (KeyValuePair<string, List<string>> keyValue in sb) 
            {
                if (Convert.ToInt32(keyValue.Key)%100 < (Convert.ToInt32(keyValue.Key) % 100)+1) 
                {
                    values.Add
                    (
                        new CellData
                        {
                             UserEnteredValue = new ExtendedValue
                             {
                                StringValue  = keyValue.Value.ToString()
                             }
                        }
                    );
                }
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
                                RowIndex = 0,
                                ColumnIndex = i
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
                Console.WriteLine(requests.ToString());
            }

            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest();

            busr.Requests = requests;

            service.Spreadsheets.BatchUpdate(busr, SpreadSheetId).Execute();
        }*/

        private static void FillTheSpreadSheet(SheetsService service, string SpreadSheetId, string[,] data)
        {
            List<Request> requests = new List<Request>();

            for (int i = 1; i < data.GetLength(1); i++)
            {
                List<CellData> values = new List<CellData>();
                for (int j = 1; j < data.GetLength(1); j++)
                {
                    values.Add
                    (
                        new CellData
                        {
                            UserEnteredValue = new ExtendedValue
                            {
                                StringValue = data[i, j]
                            }
                        }
                    );
                }

                requests.Add(
                    new Request
                    {
                        UpdateCells = new UpdateCellsRequest
                        {
                            Start = new GridCoordinate
                            {
                                SheetId = 0,
                                RowIndex = i-1,
                                ColumnIndex = 0
                            },

                            Rows = new List<RowData>
                            {
                                new RowData
                                {
                                    Values = values
                                }
                            },
                            Fields = "UserEnteredValue"
                        }
                    }
                );
            }

            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest();

            busr.Requests = requests;

            service.Spreadsheets.BatchUpdate(busr, SpreadSheetId).Execute();
        }

        private static string GetFirstCell(SheetsService service, string range, string SpreadSheetId)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(SpreadSheetId, range);

            ValueRange response = request.Execute(); 
            //ValueRange response = request.Execute() ?? throw new ArgumentNullException("request.Execute()");

            string result = null;

            foreach (var value in response.Values)
            {
                result += " " + value;
            }

            return result;

            /*return response.Values.Aggregate<IList<object>, string>(null,
                (current, value) =>
                {
                    if (value != null) return current + (" " + value[0]);
                    return null;
                });*/
        }

        private static void FillNewFile(Dictionary<string, List<string>> sb)
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