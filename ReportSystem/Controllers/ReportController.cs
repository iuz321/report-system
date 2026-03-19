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
    public IActionResult Add(string Customer, string Type, string Equipment, string Content, string Owner, int? rowIndex)
    {
        var today = DateTime.Now.ToString("yyyy-MM-dd"); // ⭐ 只日期

        if (rowIndex.HasValue)
        {
            _sheet.UpdateRow(rowIndex.Value, new List<object>
            {
                Customer,
                Type,
                Equipment,
                Content,
                Owner,
                today
            });
        }
        else
        {
            _sheet.AddRow(new List<object>
            {
                Customer,
                Type,
                Equipment,
                Content,
                Owner,
                today
            });
        }

        return RedirectToAction("Index");
    }
}