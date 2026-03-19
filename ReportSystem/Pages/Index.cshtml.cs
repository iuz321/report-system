using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace ReportSystem.Pages
{
    public class IndexModel : PageModel
    {
        [BindProperty]
        public string CustomerName { get; set; } = "";


        [BindProperty]
        public string ReportType { get; set; } = "";

        [BindProperty]
        public string ReportContent { get; set; } = "";

        [BindProperty]
        public string Owner { get; set; } = "";

        [BindProperty]
        public IFormFile? UploadFile { get; set; }

        public async Task OnPostAsync()
        {
            var uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads");

            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }

            if (UploadFile != null)
            {
                var fileName = DateTime.Now.Ticks + "_" + UploadFile.FileName;
                var filePath = Path.Combine(uploadPath, fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await UploadFile.CopyToAsync(stream);
                }
            }
        }
    }


}
