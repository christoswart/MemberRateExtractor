using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Threading.Tasks;
using System.Text.Json;

namespace MemberRateExtractor.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly IWebHostEnvironment _environment;

        public IndexModel(ILogger<IndexModel> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
        }

        [BindProperty]
        public IFormFile excelFile { get; set; }

        public string UploadMessage { get; set; }
        public bool UploadSuccess { get; set; }

        public string ProcessMessage { get; set; }
        public bool ProcessSuccess { get; set; }
        public string ProcessedFilePath { get; set; } // Store the path of the processed file


        public void OnGet()
        {
        }

        public IActionResult OnPost()
        {
            if (excelFile != null && excelFile.Length > 0)
            {
                try
                {
                    string uploadsFolder = Path.Combine(AppContext.BaseDirectory, "Spreadsheets");

                    if (!Directory.Exists(uploadsFolder))
                    {
                        Directory.CreateDirectory(uploadsFolder);
                    }

                    string filePath = Path.Combine(uploadsFolder, excelFile.FileName);
                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        excelFile.CopyTo(fileStream);
                    }

                    UploadMessage = $"File '{excelFile.FileName}' uploaded successfully.";
                    UploadSuccess = true;
                }
                catch (Exception ex)
                {
                    UploadMessage = $"Error uploading file: {ex.Message}";
                    UploadSuccess = false;
                }
            }
            else
            {
                UploadMessage = "Please select a file to upload.";
                UploadSuccess = false;
            }

            return Page();
        }

        public async Task<IActionResult> OnGetProcess()
        {
            try
            {
                string uploadsFolder = Path.Combine(AppContext.BaseDirectory, "Spreadsheets");
                string[] files = Directory.GetFiles(uploadsFolder, "*.xlsx"); // or *.xls
                if (files.Length > 0)
                {
                    var processResult = AutomationProcess.InitiateProcess();
                    ProcessMessage = processResult;
                    ProcessSuccess = true;
                    ProcessedFilePath = processResult.Contains("Finished") ? files[0] : string.Empty;
                    string filePath = files[0];
                    string processedFileName = Path.GetFileNameWithoutExtension(filePath) + "_processed" + Path.GetExtension(filePath);
                    string downloadsFolder = Path.Combine(_environment.WebRootPath, "Spreadsheets");
                    ProcessedFilePath = Path.Combine(downloadsFolder, processedFileName);
                    System.IO.File.Copy(filePath, ProcessedFilePath, true);
                }
                else
                {
                    ProcessMessage = "No excel file found to process.";
                    ProcessSuccess = false;
                }
            }
            catch (Exception ex)
            {
                ProcessMessage = $"Error processing file: {ex.Message}";
                ProcessSuccess = false;
            }
            return new JsonResult(new { success = ProcessSuccess, message = ProcessMessage, filePath = ProcessedFilePath });
        }

        public IActionResult DownloadFile(string filePath)
        {
            if (System.IO.File.Exists(filePath))
            {
                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                var fileName = Path.GetFileName(filePath);
                return File(fileBytes, "application/octet-stream", fileName);
            }
            else
            {
                return NotFound();
            }
        }
    }
}
