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
    public IActionResult Save(string Customer, string Type, string Equipment, string Content, string Owner, int rowIndex)
    {
        var today = DateTime.Now.ToString("yyyy-MM-dd");

        var row = new List<object>
        {
            Customer,
            Type,
            Equipment,
            Content,
            Owner,
            today
        };

        if (rowIndex == 0)
        {
            // 🔥 直接新增（不檢查重複）
            _sheet.AddRow(row);
        }
        else
        {
            _sheet.UpdateRow(rowIndex, row);
        }

        return RedirectToAction("Index");
    }
}