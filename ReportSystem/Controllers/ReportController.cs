using Microsoft.AspNetCore.Mvc;

public class ReportController : Controller
{
    private readonly GoogleSheetService _sheetService;

    public ReportController(GoogleSheetService sheetService)
    {
        _sheetService = sheetService;
    }

    // 顯示清單
    public IActionResult Index()
    {
        var data = _sheetService.GetAll();

        var list = data.Select(r => new Report
        {
            Customer = r.ElementAtOrDefault(0)?.ToString(),
            Type = r.ElementAtOrDefault(1)?.ToString(),
            Content = r.ElementAtOrDefault(2)?.ToString(),
            Owner = r.ElementAtOrDefault(3)?.ToString()
        }).ToList();

        return View(list);
    }

    // 新增或更新
    [HttpPost]
    public IActionResult Save(Report report)
    {
        var data = _sheetService.GetAll();

        int rowIndex = 2;
        bool found = false;

        foreach (var row in data)
        {
            if (row.Count > 0 && row[0]?.ToString() == report.Customer)
            {
                found = true;
                break;
            }
            rowIndex++;
        }

        var newRow = new List<object>
        {
            report.Customer,
            report.Type,
            report.Content,
            report.Owner
        };

        if (found)
        {
            _sheetService.UpdateRow(rowIndex, newRow);
        }
        else
        {
            _sheetService.AddRow(newRow);
        }

        return RedirectToAction("Index");
    }
}