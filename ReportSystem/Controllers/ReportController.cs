using Microsoft.AspNetCore.Mvc;

public class ReportController : Controller
{
    private readonly GoogleSheetService _sheet;

    public ReportController()
    {
        _sheet = new GoogleSheetService();
    }

    // 🔹 首頁
    public IActionResult Index()
    {
        var data = _sheet.GetAll();
        return View(data);
    }

    // 🔹 新增 / 更新
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

        var all = _sheet.GetAll();

        // 🔥 新增時檢查是否重複
        if (rowIndex == 0)
        {
            foreach (var r in all)
            {
                var existCustomer = r.Count > 0 ? r[0]?.ToString() : "";

                if (existCustomer == Customer)
                {
                    TempData["Error"] = "❌ 客戶已存在，不能新增";
                    return RedirectToAction("Index");
                }
            }

            _sheet.AddRow(row);
        }
        else
        {
            // 🔹 編輯允許
            _sheet.UpdateRow(rowIndex, row);
        }

        return RedirectToAction("Index");
    }
}