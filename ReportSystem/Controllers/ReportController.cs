using Microsoft.AspNetCore.Mvc;

public class ReportController : Controller
{
    private readonly GoogleSheetService _sheet;

    public ReportController()
    {
        _sheet = new GoogleSheetService();
    }

    public IActionResult Index()
    {
        var data = _sheet.GetAll();
        return View(data);
    }

    [HttpPost]
    public IActionResult Save(
        int rowIndex,
        string Customer,
        string Type,
        string Equipment,
        string Content,
        string Owner)
    {
        // 🔥 防爆處理（關鍵）
        Content = Content?
            .Replace("\r\n", " ")
            .Replace("\r", " ")
            .Replace("\n", " ")
            .Replace("\t", " ")
            .Trim();

        if (!string.IsNullOrEmpty(Content) && Content.Length > 500)
        {
            Content = Content.Substring(0, 500);
        }

        var today = DateTime.Now.ToString("yyyy-MM-dd");

        var row = new List<object>
        {
            Customer ?? "",
            Type ?? "",
            Equipment ?? "",
            Content ?? "",
            Owner ?? "",
            today
        };

        if (rowIndex == 0)
        {
            _sheet.AddRow(row);
        }
        else
        {
            _sheet.UpdateRow(rowIndex, row);
        }

        return RedirectToAction("Index");
    }
}