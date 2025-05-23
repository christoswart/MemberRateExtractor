using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Diagnostics;
using System.Drawing;

namespace MemberRateExtractor
{
    public class AutomationProcess
    {
        private static int _chromeProcessId = 0;
        public static string InitiateProcess()
        {
            // Search for Chrome Browser exe file
            string chromePath = GetChromePath();
            if (!string.IsNullOrEmpty(chromePath))
            {
                ChromeOptions chromeOptions = StartChromeDebug(chromePath);

                HotelSheetRow[] hotelRows = Array.Empty<HotelSheetRow>();
                string otaSpreadsheet = GetHotelsFromSpreadSheetData(ref hotelRows);

                LoopOverUrls(chromeOptions, hotelRows);

                SaveScreenshotLinksToSpreadsheet(otaSpreadsheet, hotelRows);
                Console.WriteLine("Finished processing.");
                Process.GetProcessById(_chromeProcessId).Kill();
            }
            else
            {
                Console.WriteLine("Chrome not found. Unable to start scraping process.");
            }
            return "Automation Process Finished";
        }

        private static string GetHotelsFromSpreadSheetData(ref HotelSheetRow[] hotelRows)
        {
            string spreadsheetDir = GetSpreadSheetDirectory();
            string firstSpreadsheet = Directory.GetFiles(spreadsheetDir, "*.xlsx").FirstOrDefault();

            if (firstSpreadsheet != null)
            {
                Console.WriteLine($"Reading first worksheet: {firstSpreadsheet}");
                hotelRows = GetUrlsFromSpreadsheet(firstSpreadsheet);
            }
            else
            {
                Console.WriteLine("No spreadsheets found in the directory.");
            }

            return firstSpreadsheet;
        }

        private static ChromeOptions StartChromeDebug(string chromePath)
        {
            // Run the Chrome application
            ProcessStartInfo chromeProcessInfo = new ProcessStartInfo
            {
                FileName = chromePath,
                UseShellExecute = true,
                ArgumentList = { "--remote-debugging-port=9222", "--user-data-dir=C:\\ChromeDebug" }
            };
            _chromeProcessId = Process.Start(chromeProcessInfo).Id;

            // Connect to the Chrome instance using ChromeDriver
            var chromeOptions = new ChromeOptions();
            chromeOptions.DebuggerAddress = "localhost:9222";
            //chromeOptions.AddArgument("--headless");
            return chromeOptions;
        }

        private static void LoopOverUrls(ChromeOptions chromeOptions, HotelSheetRow[] hotelRows)
        {
            using (var driver = new ChromeDriver(chromeOptions))
            {
                Console.WriteLine("Connected to Chrome instance and navigated to OTA website.");
                string screenshotsDir = GetScreenshotDirectory();

                foreach (var hotel in hotelRows)
                {
                    // Switch to the newly opened tab
                    List<string> handles = new List<string>(driver.WindowHandles);
                    driver.SwitchTo().Window(handles[handles.Count - 1]);

                    // Navigate to the URL in the new tab
                    driver.Navigate().GoToUrl(hotel.Url);

                    // Take a screenshot
                    CaptureFullPageScreenshot(driver, screenshotsDir, hotel);

                    string pageMarkdown = ConvertWebsiteToMarkdown(driver);

                    // Open a new tab
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");

                    // Log for debugging 
                    Console.WriteLine($"Opened URL: {hotel.Url}");
                }
            }
        }

        public static void CaptureFullPageScreenshot(IWebDriver driver, string outputPath, HotelSheetRow hotel)
        {
            // Create Directory for Hotel Screenshots
            var hotelOutputPath = Path.Combine(outputPath, $"{hotel.HotelName}");
            if (!Directory.Exists(hotelOutputPath))
            {
                Directory.CreateDirectory(hotelOutputPath);
            }

            // Get the total height and width of the page
            long totalHeight = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.scrollHeight");
            long totalWidth = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.scrollWidth");
            long viewportHeight = (long)((IJavaScriptExecutor)driver).ExecuteScript("return window.innerHeight");

            int partCount = 1;
            // Create a bitmap to store the full screenshot
            using (Bitmap fullScreenshot = new Bitmap((int)totalWidth, (int)totalHeight))
            {
                using (Graphics graphics = Graphics.FromImage(fullScreenshot))
                {
                    long currentY = 0;

                    while (currentY < totalHeight)
                    {
                        // Scroll to the current position
                        ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, {currentY})");
                        System.Threading.Thread.Sleep(200); // Wait for the page to render

                        // Capture the screenshot of the current viewport
                        Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                        screenshot.SaveAsFile(Path.Combine(hotelOutputPath, $"part{partCount}.png"));
                        using (MemoryStream ms = new MemoryStream(screenshot.AsByteArray))
                        {
                            using (Bitmap viewportScreenshot = new Bitmap(ms))
                            {
                                // Draw the viewport screenshot onto the full screenshot
                                graphics.DrawImage(viewportScreenshot, 0, (int)currentY);
                            }
                        }

                        partCount++;
                        currentY += viewportHeight;
                    }
                }

                // Save the full screenshot to a file
                string newFileName = Path.Combine(hotelOutputPath, $"screenshot_{Guid.NewGuid()}.png");
                fullScreenshot.Save(newFileName);
                hotel.ScreenshotLink = newFileName;
                Console.WriteLine($"Full page screenshot saved to: {newFileName}");
            }
        }

        private static HotelSheetRow[] GetUrlsFromSpreadsheet(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    var hotelNameColumn = worksheet.Cells["1:1"].FirstOrDefault(cell => cell.Text == "Hotel Name")?.Start.Column ?? -1;
                    if (hotelNameColumn == -1)
                        Console.WriteLine("Hotel Name column not found.");

                    var countryColumn = worksheet.Cells["1:1"].FirstOrDefault(cell => cell.Text == "Hotel Name")?.Start.Column ?? -1;
                    if (countryColumn == -1)
                        Console.WriteLine("Country column not found.");

                    var urlColumn = worksheet.Cells["1:1"].FirstOrDefault(cell => cell.Text == "URL")?.Start.Column ?? -1;
                    if (urlColumn == -1)
                    {
                        Console.WriteLine("URL column not found.");
                        return Array.Empty<HotelSheetRow>();
                    }

                    var hotelSheetRows = new List<HotelSheetRow>();
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var url = worksheet.Cells[row, urlColumn].Text;
                        if (!string.IsNullOrEmpty(url))
                        {
                            hotelSheetRows.Add(new HotelSheetRow
                            {
                                HotelName = worksheet.Cells[row, hotelNameColumn].Text ?? "",
                                Country = worksheet.Cells[row, countryColumn].Text ?? "",
                                Url = worksheet.Cells[row, urlColumn].Text,
                                ScreenshotLink = "",
                                MemberRate = ""
                            });
                        }
                    }

                    return hotelSheetRows.ToArray();
                }
            }

            return Array.Empty<HotelSheetRow>();
        }

        private static void SaveScreenshotLinksToSpreadsheet(string filePath, HotelSheetRow[] hotelRows)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    var urlColumn = worksheet.Cells["1:1"].FirstOrDefault(cell => cell.Text == "URL")?.Start.Column ?? -1;
                    var screenshotLinkColumn = worksheet.Cells["1:1"].FirstOrDefault(cell => cell.Text == "Screenshot Link")?.Start.Column ?? -1;

                    if (screenshotLinkColumn == -1)
                    {
                        screenshotLinkColumn = worksheet.Dimension.End.Column + 1;
                        worksheet.Cells[1, screenshotLinkColumn].Value = "Screenshot Link";
                    }

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var url = worksheet.Cells[row, urlColumn].Text;
                        var hotelRow = hotelRows.FirstOrDefault(h => h.Url == url);
                        if (hotelRow != null)
                        {
                            worksheet.Cells[row, screenshotLinkColumn].Value = hotelRow.ScreenshotLink;
                        }
                    }

                    package.Save();
                }
            }
        }

        private static string GetScreenshotDirectory()
        {
            // Create directory for screenshots if it doesn't exist
            string screenshotsDir = Path.Combine(AppContext.BaseDirectory, "Screenshots");
            if (!Directory.Exists(screenshotsDir))
            {
                Directory.CreateDirectory(screenshotsDir);
            }

            return screenshotsDir;
        }

        private static string GetSpreadSheetDirectory()
        {
            // Create directory for screenshots if it doesn't exist
            string screenshotsDir = Path.Combine(AppContext.BaseDirectory, "Spreadsheets");
            if (!Directory.Exists(screenshotsDir))
            {
                Directory.CreateDirectory(screenshotsDir);
            }

            return screenshotsDir;
        }

        private static string GetChromePath()
        {
            string[] possiblePaths = {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Google", "Chrome", "Application", "chrome.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Google", "Chrome", "Application", "chrome.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google", "Chrome", "Application", "chrome.exe")
        };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    Console.WriteLine($"Chrome found at: {path}");
                    return path;
                }
            }

            return string.Empty;
        }
        private static string ConvertWebsiteToMarkdown(IWebDriver driver)
        {
            // Get the website's HTML content
            string htmlContent = driver.PageSource;

            // Use a library like HtmlAgilityPack to parse the HTML and convert it to Markdown
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlContent);

            var markdown = new System.Text.StringBuilder();
            foreach (var node in doc.DocumentNode.ChildNodes)
            {
                ConvertNodeToMarkdown(node, markdown);
            }

            return markdown.ToString();
        }

        private static void ConvertNodeToMarkdown(HtmlAgilityPack.HtmlNode node, System.Text.StringBuilder markdown)
        {
            switch (node.Name)
            {
                case "h1":
                    markdown.AppendLine($"# {node.InnerText}");
                    break;
                case "h2":
                    markdown.AppendLine($"## {node.InnerText}");
                    break;
                case "h3":
                    markdown.AppendLine($"### {node.InnerText}");
                    break;
                case "p":
                    markdown.AppendLine(node.InnerText);
                    break;
                case "ul":
                    foreach (var li in node.SelectNodes("li"))
                    {
                        markdown.AppendLine($"- {li.InnerText}");
                    }
                    break;
                case "ol":
                    int index = 1;
                    foreach (var li in node.SelectNodes("li"))
                    {
                        markdown.AppendLine($"{index}. {li.InnerText}");
                        index++;
                    }
                    break;
                default:
                    if (node.HasChildNodes)
                    {
                        foreach (var child in node.ChildNodes)
                        {
                            ConvertNodeToMarkdown(child, markdown);
                        }
                    }
                    break;
            }
        }
    }
}
