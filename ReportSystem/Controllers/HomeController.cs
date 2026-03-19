using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using ReportSystem.Models;

namespace ReportSystem.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        // 👉 首頁直接導向 Report
        public IActionResult Index()
        {
            return RedirectToAction("Index", "Report");
        }

        // 👉 隱私頁（可有可無）
        public IActionResult Privacy()
        {
            return View();
        }

        // 👉 錯誤頁（避免 500 崩掉）
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel
            {
                RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier
            });
        }
    }
}