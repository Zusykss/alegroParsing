using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ParserAlegro;
using System.Reflection;

public class GoogleSheetsHelper
{
    private readonly string _spreadSheetId;
    private SheetsService _sheetsService;
    private List<string> HeaderOrder;
    private int sheetIndex = 1;
    private const int MAX_ROW_COUNT = 25000;




    public GoogleSheetsHelper(string spreadSheetId, string clientId, string clientSecret)
    {
        ServiceInit(clientId, clientSecret);
        _spreadSheetId = spreadSheetId;

    }

    private Sheet CreateOrGetNewSheet(int sheetIndex = 0)
    {
        var sheetTitle = DateTime.Now.ToString("d-M-yyyy");
        sheetTitle += "_" + (sheetIndex);

        var spreadSheetRequest = _sheetsService.Spreadsheets.Get(_spreadSheetId);
        var spreadSheetInCloud = spreadSheetRequest.Execute();
        var sheet = spreadSheetInCloud.Sheets.FirstOrDefault(s => s.Properties.Title == sheetTitle);
        if (sheet != null)
        {
            return sheet;
        }


        var batches = new BatchUpdateSpreadsheetRequest();
        batches.IncludeSpreadsheetInResponse = true;
        batches.Requests = new List<Request>();
        var request = new Request();
        var addSheetRequest = new AddSheetRequest()
        {
            Properties = new SheetProperties()
            {
                Title = sheetTitle
            }
        };
        request.AddSheet = addSheetRequest;
        batches.Requests.Add(request);

        var result = _sheetsService.Spreadsheets.BatchUpdate(batches, _spreadSheetId).Execute();

        sheet = result.UpdatedSpreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == sheetTitle);







        return sheet;
    }

    private void FormatCellsHorizontalSize(Sheet sheet)
    {
        var batchFormatRequest = new BatchUpdateSpreadsheetRequest();
        batchFormatRequest.Requests = new List<Request>();

        var formatRequest = new Request();
        var resizeDimentions = new UpdateDimensionPropertiesRequest();
        resizeDimentions.Fields = "pixelSize";
        resizeDimentions.Properties = new DimensionProperties();
        resizeDimentions.Properties.PixelSize = 125;
        resizeDimentions.Range = new DimensionRange();
        resizeDimentions.Range.Dimension = "Columns";
        resizeDimentions.Range.StartIndex = 0;
        resizeDimentions.Range.EndIndex = 10;
        resizeDimentions.Range.SheetId = sheet.Properties.SheetId;

        formatRequest.UpdateDimensionProperties = resizeDimentions;
        batchFormatRequest.Requests.Add(formatRequest);
        _sheetsService.Spreadsheets.BatchUpdate(batchFormatRequest, _spreadSheetId).Execute();
    }

    private void ServiceInit(string clientId, string clientSecret)
    {
        string[] Scopes = { SheetsService.Scope.Spreadsheets };

        var credentials = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
        {
            ClientId = clientId,
            ClientSecret = clientSecret
        },
            Scopes,
            "user",
            CancellationToken.None
            ).Result;
        var service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credentials,
            ApplicationName = "Sheets api test"
        });
        this._sheetsService = service;
    }

    private void WriteDataToSheet(Sheet sheet, IEnumerable<ObjectAlegro> objects)
    {
        var chunked = objects.Chunk(1000).ToList();


        foreach (var chunk in chunked)
        {
            var rangeBody = new ValueRange();
            rangeBody.Values = new List<IList<object>>();


            var nextSheetObjects = new List<ObjectAlegro>();


            var sheetCount = GetSheetCount(sheet);

            foreach (var alergoObject in chunk)
            {
                var row = new List<object>();

                string photos = "";
                for (int k = 0; k < alergoObject.Photos.Count; k++)
                {
                    //Console.WriteLine(obj["medium"].ToString().Replace("\\u002F", "/"));
                    //objectAlegro.Photos.Add(obj["medium"].ToString().Replace("\\u002F", "/"));
                    photos += alergoObject.Photos[k].Replace("\\u002F", "/").Replace("/s360/", "/s2048/");
                    if (k + 1 != alergoObject.Photos.Count) photos += ",";
                }
                string join = String.Join(',', alergoObject.Photos).Replace("/s360/", "/s2048/");
                row.Add(alergoObject.Producer);
                if (string.IsNullOrEmpty(alergoObject.Url)) row.Add(alergoObject.NumberLot);
                else row.Add(alergoObject.Url.Replace("\\u002F", "/"));
                row.Add(alergoObject.CatalogNumber);
                row.Add(Convert.ToDouble(alergoObject.Price.Replace(".", ",")));
                row.Add(alergoObject.Quantity);
                row.Add(photos);
                row.Add("used");


                if (rangeBody.Values.Count >= MAX_ROW_COUNT - sheetCount)
                {
                    nextSheetObjects.Add(alergoObject);
                }
                else
                {
                    rangeBody.Values.Add(row);
                }
            }

            var request = _sheetsService.Spreadsheets.Values.Append(rangeBody, _spreadSheetId, $"{sheet.Properties.Title}!A:L");
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            request.Execute();
            if (nextSheetObjects.Count != 0)
            {
                sheetIndex++;
                while (true)
                {
                    sheet = CreateOrGetNewSheet(sheetIndex);
                    sheetCount = GetSheetCount(sheet);

                    if (sheetCount >= MAX_ROW_COUNT)
                    {
                        continue;
                    }
                    FormatCellsHorizontalSize(sheet);
                    WriteDataToSheet(sheet, nextSheetObjects);
                }
            }
        }
    }

    public int GetSheetCount(Sheet sheet)
    {
        var getCountRequset = _sheetsService.Spreadsheets.Values.Get(_spreadSheetId, $"{sheet.Properties.Title}!A:G");
        var count = getCountRequset.Execute().Values?.Count ?? 0;
        return count;
    }
    public void WriteList(List<ObjectAlegro> objects)
    {
        while (true)
        {
            var sheet = CreateOrGetNewSheet(sheetIndex);

            var count = GetSheetCount(sheet);
            if (count >= MAX_ROW_COUNT)
            {
                sheetIndex++;
                continue;
            }

            FormatCellsHorizontalSize(sheet);
            WriteDataToSheet(sheet, objects);
            break;
        }
    }
}