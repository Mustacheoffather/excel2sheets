using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace excel2sheets
{
    public class Spreadsheet
    {
        private static string ClientSecret = "credentials.json";
        private static readonly string[] Scopes = {SheetsService.Scope.Spreadsheets};
        private static readonly string ApplicationName = "excel2sheets";
        public string SpreadSheetId { get; set; }

        public UserCredential GetCredential()
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

        public SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
        }    

        public void FillTheSpreadSheet(SheetsService service, string SpreadSheetId, string[,] data)
        {
            List<Request> requests = new List<Request>();

            for (int i = 1; i < data.GetLength(0); i++)
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
                //Console.WriteLine($"added \t{i} row");

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

    }
}