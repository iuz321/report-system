using Microsoft.AspNetCore.Mvc;

public class ReportController : Controller
{
    private readonly GoogleSheetService _sheetService;

    public ReportController()
    {
        _sheetService = new GoogleSheetService();
    }

    public IActionResult Index()
    {
        var data = _sheetService.GetAll();
        return View(data);
    }

    [HttpPost]
    public IActionResult Save(string Customer, string Type, string Equipment, string Content, string Owner, int RowIndex)
    {
        // 🔥 修正跨裝置換行問題（重點）
        Content = Content?
            .Replace("\r\n", "\n")
            .Replace("\r", "\n")
            .Trim();

        Customer = Customer?.Trim();
        Type = Type?.Trim();
        Equipment = Equipment?.Trim();
        Owner = Owner?.Trim();

        var row = new List<object>
        {
            Customer,
            Type,
            Equipment,
            Content,
            Owner
        };

        if (RowIndex == 0)
        {
            _sheetService.AddRow(row);
        }
        else
        {
            _sheetService.UpdateRow(RowIndex, row);
        }

        return RedirectToAction("Index");
    }
}