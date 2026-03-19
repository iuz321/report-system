using Microsoft.AspNetCore.Mvc;

public class ReportController : Controller
{
    private readonly GoogleSheetService _sheet;

    public ReportController()
    {
        _sheet = new GoogleSheetService();
    }

    // 📄 顯示資料
    public IActionResult Index()
    {
        var data = _sheet.GetAll();
        return View(data);
    }

    // ➕ 顯示新增頁
    public IActionResult Create()
    {
        return View();
    }

    // ➕ 新增資料
    [HttpPost]
    public IActionResult Create(string Customer, string Type, string Equipment, string Content, string Owner)
    {
        _sheet.AddRow(new List<object>
        {
            Customer,
            Type,
            Equipment, // ⭐ 新增欄位
            Content,
            Owner
        });

        return RedirectToAction("Index");
    }
}