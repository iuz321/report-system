using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

public class GoogleSheetService
{
    private readonly SheetsService _service;

    private readonly string _spreadsheetId = "1akc8lPUQ9_6BB1sxOpQbuzqKSDGWNSZJL-m5pnBvdeM";

    public GoogleSheetService()
    {
        GoogleCredential credential;

        // 🔥 先讀 Render 環境變數
        var json = Environment.GetEnvironmentVariable("GOOGLE_JSON");

        if (!string.IsNullOrEmpty(json))
        {
            using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json)))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }
        }
        else
        {
            // 🔹 本機 fallback
            var path = Path.Combine(Directory.GetCurrentDirectory(), "Credentials", "google.json");

            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }
        }

        _service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "ReportSystem"
        });
    }

    // 🔹 讀取（A~F）
    public IList<IList<object>> GetAll()
    {
        var request = _service.Spreadsheets.Values.Get(_spreadsheetId, "work!A2:F");
        var response = request.Execute();
        return response.Values ?? new List<IList<object>>();
    }

    // 🔹 新增（含日期）
    public void AddRow(List<object> row)
    {
        // 自動加日期
        row.Add(DateTime.Now.ToString("yyyy-MM-dd"));

        var body = new ValueRange
        {
            Values = new List<IList<object>> { row }
        };

        var request = _service.Spreadsheets.Values.Append(body, _spreadsheetId, "work!A:F");
        request.ValueInputOption =
            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

        request.Execute();
    }

    // 🔹 更新（含日期）
    public void UpdateRow(int rowIndex, List<object> row)
    {
        // 更新日期
        row.Add(DateTime.Now.ToString("yyyy-MM-dd"));

        var range = $"work!A{rowIndex}:F{rowIndex}";

        var body = new ValueRange
        {
            Values = new List<IList<object>> { row }
        };

        var request = _service.Spreadsheets.Values.Update(body, _spreadsheetId, range);
        request.ValueInputOption =
            SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

        request.Execute();
    }
}