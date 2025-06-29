using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using System.ComponentModel.DataAnnotations;
using ClosedXML.Excel;
using System.IO;
using ESHApp.Models;

namespace ESHApp.Pages
{
	public class ProcessPayComFileModel : PageModel
    {
        private readonly ILogger<ProcessPayComFileModel> _logger;
        //private async Task GenerateOutputFilesAsync(List<PaycomRecord> paycomRecords, Dictionary<string, ...> customerMap, Dictionary<string, ...> employeeMap);
        private Dictionary<string, (string SegmentCode, string Customer, string Department)>? _customerSegmentData;
        private Dictionary<string, (string ExternalId, string Title, string Name)>? _employeesTitleData;

        private DateTime cacheExpiry;

        private readonly string[] RequiredColumns = new[]
        {
            "Distribution","Employee_Code","Employee_Name","Intern_Code", "DOL_Status", "Gross_Hours(DR1)","SALARY(DR1)", "FICA(DR1)", "NY_Metro_CTM(DR1)", "401K_MATCH(DR1)"
        };

        public Dictionary<string, (string SegmentCode, string Customer, string Department)> CustomerSegmentData
        {
            get
            {
                if (_customerSegmentData == null || DateTime.UtcNow > cacheExpiry)
                {
                    _customerSegmentData = LoadCustomerSegmentData();
                    cacheExpiry = DateTime.UtcNow.AddHours(24);
                }
                return _customerSegmentData;
            }
        }

        public Dictionary<string, (string ExternalId, string Title, string Name)> EmployeesTitleData
        {
            get
            {
                if (_employeesTitleData == null || DateTime.UtcNow > cacheExpiry)
                {
                    _employeesTitleData = LoadEmployeeTitleData();
                    cacheExpiry = DateTime.UtcNow.AddHours(24);
                }
                return _employeesTitleData;
            }
        }

        public ProcessPayComFileModel(ILogger<ProcessPayComFileModel> logger)
        {
            _logger = logger;
        }

        // -- Bind properties to the form --
        [BindProperty]
        [Required]
        public IFormFile PaycomFile { get; set; }

        [BindProperty]
        [DataType(DataType.Date)]
        [Required]
        public DateTime? PayPeriodStart { get; set; }

        [BindProperty]
        [DataType(DataType.Date)]
        [Required]
        public DateTime? PayPeriodEnd { get; set; }

        // -- messages to pass back to the page --
        public string ErrorMessage { get; set; }
        public string SuccessMessage { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPost()
        {
            _logger.LogInformation("Processing Paycom File ...");

            // basic validation
            if (!ModelState.IsValid)
            {
                ErrorMessage = "Please ensure all fields are filled out correctly.";
                return Page();
            }

            if (PaycomFile == null || PaycomFile.Length == 0)
            {
                ErrorMessage = "No Paycom file selected.";
                return Page();
            }

            if (PayPeriodStart == null || PayPeriodEnd == null)
            {
                ErrorMessage = "Please specify both pay period dates.";
                    return Page();
            }

            if (!ValidatePaycomFileAsync())
            {
                _logger.LogInformation("Upload file is missing required columns, not in the right format.");
                ErrorMessage = "Upload file is missing required columns.";
                return Page();
            }

            // load reference data
            LoadCustomerSegmentData();
            var employeesTitleData = LoadEmployeeTitleData();

            // load paycom records
            var paycomRecords = LoadPaycomRecords();

            // generate 4 output csv files
            //GenerateOutputFilesAsync(paycomRecords, customerMap, employeeMap);

            SuccessMessage = "Processing complete!";
            return Page();
        }

        private bool ValidatePaycomFileAsync ()
        {
            using var stream = PaycomFile.OpenReadStream();
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.First();

            // Read header row (assume first row)
            var headerRow = worksheet.Row(1);
            var columnsInFile = headerRow.CellsUsed().Select(c => c.GetString().Trim()).ToList();

            // Check if all required column exist
            foreach (var requiredCol in RequiredColumns)
            {
                if (!columnsInFile.Contains(requiredCol))
                {
                    return false;
                }
            }
            _logger.LogInformation("Paycom file validated.");
            return true;
        }

        private Dictionary<string, (string SegmentCode, string Customer, string Department)> LoadCustomerSegmentData()
        {
            var result = new Dictionary<string, (string SegmentCode, string Customer, string Department)>();
            var referenceFilePath = Path.Combine(Environment.CurrentDirectory, "Data", "Customer & Segment Code list.xlsx");
            using var workbook = new XLWorkbook(referenceFilePath);            
            var worksheet = workbook.Worksheet("Segmented Department list");

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var segmentCode = row.Cell(1).GetString().Trim();                
                var customer = row.Cell(2).GetString().Trim();
                var department = row.Cell(3).GetString().Trim();
                result[segmentCode] = ( segmentCode, customer, department );
            }
            return result;
        }

        private Dictionary<string, (string ExternalId, string Title, string Name)> LoadEmployeeTitleData()
        {
            var result = new Dictionary<string, (string ExternalId, string Title, string Name)>();
            var referenceFilePath = Path.Combine(Environment.CurrentDirectory, "Data", "EmployeeTitle.xlsx");
            using var workbook = new XLWorkbook(referenceFilePath);
            var worksheet = workbook.Worksheets.First();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var externalID = row.Cell(1).GetString().Trim();
                var title = row.Cell(2).GetString().Trim();
                var name = row.Cell(3).GetString().Trim();
                result[name] = ( externalID, title, name );
            }
            return result;
        }

        private List<PaycomRecord> LoadPaycomRecords()
        {
            return null;
        }
    }
}
