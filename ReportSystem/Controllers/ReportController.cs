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

    // ⭐ 單一入口（新增 + 更新）
    [HttpPost]
    public IActionResult Add(string Customer, string Type, string Equipment, string Content, string Owner, int? rowIndex)
    {
        if (rowIndex.HasValue)
        {
            // 更新
            _sheet.UpdateRow(rowIndex.Value, new List<object>
            {
                Customer,
                Type,
                Equipment,
                Content,
                Owner
            });
        }
        else
        {
            // 新增
            _sheet.AddRow(new List<object>
            {
                Customer,
                Type,
                Equipment,
                Content,
                Owner
            });
        }

        return RedirectToAction("Index");
    }
}