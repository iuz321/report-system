using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Text;

public class GoogleSheetService
{
    private readonly SheetsService _service;
    private readonly string _spreadsheetId = "1akc8lPUQ9_6BB1sxOpQbuzqKSDGWNSZJL-m5pnBvdeM";

    public GoogleSheetService()
    {
        var json = Environment.GetEnvironmentVariable("GOOGLE_CREDENTIAL_JSON");

        GoogleCredential credential;

        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
        {
            credential = GoogleCredential.FromStream(stream)
                .CreateScoped(SheetsService.Scope.Spreadsheets);
        }

        _service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "ReportSystem"
        });
    }

    public IList<IList<object>> GetAll()
    {
        var request = _service.Spreadsheets.Values.Get(_spreadsheetId, "work!A2:F");
        var response = request.Execute();
        return response.Values ?? new List<IList<object>>();
    }

    public void AddRow(List<object> row)
    {
        var body = new ValueRange { Values = new List<IList<object>> { row } };

        var request = _service.Spreadsheets.Values.Append(body, _spreadsheetId, "work!A:F");
        request.ValueInputOption =
            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

        request.Execute();
    }

    public void UpdateRow(int rowIndex, List<object> row)
    {
        var range = $"work!A{rowIndex}:F{rowIndex}";

        var body = new ValueRange { Values = new List<IList<object>> { row } };

        var request = _service.Spreadsheets.Values.Update(body, _spreadsheetId, range);
        request.ValueInputOption =
            SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

        request.Execute();
    }
}