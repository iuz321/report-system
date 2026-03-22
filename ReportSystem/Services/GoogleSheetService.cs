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

        var json = Environment.GetEnvironmentVariable("GOOGLE_JSON");

        // 🔥 ===== 驗證開始（重點）=====
        if (string.IsNullOrEmpty(json))
        {
            throw new Exception("❌ GOOGLE_JSON 沒設定");
        }

        if (!json.Contains("private_key"))
        {
            throw new Exception("❌ GOOGLE_JSON 格式錯誤（缺少 private_key）");
        }
        // 🔥 ===== 驗證結束 =====

        using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json)))
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

    // 🔹 測試讀取
    public IList<IList<object>> GetAll()
    {
        try
        {
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, "work!A2:F");
            var response = request.Execute();
            return response.Values ?? new List<IList<object>>();
        }
        catch (Exception ex)
        {
            throw new Exception("❌ 讀取 Google Sheet 失敗：" + ex.Message);
        }
    }

    // 🔹 新增
    public void AddRow(List<object> row)
    {
        try
        {
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
        catch (Exception ex)
        {
            throw new Exception("❌ 新增失敗：" + ex.Message);
        }
    }

    // 🔹 更新
    public void UpdateRow(int rowIndex, List<object> row)
    {
        try
        {
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
        catch (Exception ex)
        {
            throw new Exception("❌ 更新失敗：" + ex.Message);
        }
    }
}