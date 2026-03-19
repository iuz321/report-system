using Microsoft.AspNetCore.Mvc;

public class ReportController : Controller
{
    private readonly GoogleSheetService _sheet;

    public ReportController()
    {
        _sheet = new GoogleSheetService();
    }

    // 🔹 單頁（顯示 + 表單）
    public IActionResult Index()
    {
        var data = _sheet.GetAll();
        return View(data);
    }

    // 🔹 新增（同頁送出）
    [HttpPost]
    public IActionResult Add(string Customer, string Type, string Equipment, string Content, string Owner)
    {
        _sheet.AddRow(new List<object>
        {
            Customer,
            Type,
            Equipment, // ⭐ 新增
            Content,
            Owner
        });

        return RedirectToAction("Index");
    }

    // 🔹 編輯（點擊）
    [HttpPost]
    public IActionResult Update(int rowIndex, string Customer, string Type, string Equipment, string Content, string Owner)
    {
        _sheet.UpdateRow(rowIndex, new List<object>
        {
            Customer,
            Type,
            Equipment, // ⭐ 新增
            Content,
            Owner
        });

        return RedirectToAction("Index");
    }
}