# Member Rate Extractor 
## Member Rate Extractor is an ASP.NET Core Razor Pages web application that automates the extraction and processing of data from Excel spreadsheets and websites.
### Features
•	Excel File Upload: Users can upload Excel files (.xlsx or .xls) through the web interface.
•	Automated Processing: After uploading, users can trigger an automated process that reads the uploaded spreadsheet, extracts relevant data (such as URLs), and performs automated web tasks using Selenium WebDriver and Chrome.
•	Website Content Extraction: The application visits websites listed in the spreadsheet, retrieves their HTML content, and converts it to Markdown format using HtmlAgilityPack.
•	Processed File Download: Once processing is complete, users can download the processed results directly from the web interface.
•	User Feedback: The application provides clear success/error messages for both upload and processing steps.
### Technologies Used
•	ASP.NET Core Razor Pages (.NET 8)
•	EPPlus (for Excel file handling)
•	Selenium WebDriver (for browser automation)
•	HtmlAgilityPack (for HTML parsing and Markdown conversion)
•	Bootstrap (for UI styling)
### Typical Workflow
1.	User uploads an Excel file containing URLs.
2.	User clicks "Process File" to start automated extraction.
3.	The application visits each URL, extracts and converts content, and saves the results.
4.	User downloads the processed file.
