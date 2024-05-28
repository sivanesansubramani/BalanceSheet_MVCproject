using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.IO;
using System.Web;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using CsvHelper;
using System.Globalization;

namespace BalanceSheet.Controllers
{
    public class ExcelController : Controller
    {
        // GET: Excel
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                // Validate the file type
                if (!file.FileName.EndsWith(".csv") || file.ContentType != "text/csv")
                {
                    TempData["Error"] = "Please upload a valid CSV file.";
                    return RedirectToAction("Index");
                }

                try
                {
                    using (var stream = new MemoryStream())
                    {
                        file.CopyTo(stream);
                        stream.Position = 0;
                        using (var reader = new StreamReader(stream))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            var records = csv.GetRecords<CsvData>();
                            var processedData = ProcessCsvData(records);

                            // Generate CSV for download
                            using (var memoryStream = new MemoryStream())
                            using (var writer = new StreamWriter(memoryStream))
                            using (var csvWriter = new CsvWriter(writer, CultureInfo.InvariantCulture))
                            {
                                csvWriter.WriteRecords(processedData);
                                writer.Flush();
                                memoryStream.Position = 0;

                                return File(memoryStream.ToArray(), "text/csv", "ProcessedFile.csv");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    TempData["Error"] = $"An error occurred while processing the file: {ex.Message}";
                    return RedirectToAction("Index");
                }
            }

            // If no file is selected, redirect to index with an error
            TempData["Error"] = "Please select a file to upload.";
            return RedirectToAction("Index");
        }

        private List<ProcessedData> ProcessCsvData(IEnumerable<CsvData> records)
        {
            var processedData = new List<ProcessedData>();
            decimal openingBalance = 0;
            decimal closingBalance = 0;
            DateTime currentDate = DateTime.MinValue.Date;

            foreach (var record in records)
            {
                var ChoosenDate = DateTime.ParseExact(record.ChoosenDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                var head = record.Head_Id;
                var narration = record.Narration;
                var credit = string.IsNullOrWhiteSpace(record.Credit) ? 0 : decimal.Parse(record.Credit);
                var debit = string.IsNullOrWhiteSpace(record.Debit) ? 0 : decimal.Parse(record.Debit);

                // If currentDate is not MinValue and it's a new date, add Closing Balance for the previous date
                if (currentDate != DateTime.MinValue.Date && currentDate != ChoosenDate)
                {
                    processedData.Add(new ProcessedData
                    {
                        ChoosenDate = currentDate,
                        Head_Id = "",
                        Narration = "Closing Balance",
                        Credit = 0,
                        Debit = closingBalance
                    });

                    // Set the opening balance for the new date as the closing balance of the previous date
                    openingBalance = closingBalance;
                }

                // Add Opening Balance for the new date
                if (currentDate != ChoosenDate)
                {
                    currentDate = ChoosenDate;
                    processedData.Add(new ProcessedData
                    {
                        ChoosenDate = currentDate,
                        Head_Id = "",
                        Narration = "Opening Balance",
                        Credit = openingBalance,
                        Debit = 0
                    });
                }

                // Add the record
                processedData.Add(new ProcessedData
                {
                    ChoosenDate = ChoosenDate,
                    Head_Id = head,
                    Narration = narration,
                    Credit = credit,
                    Debit = debit
                });

                // Calculate the closing balance for the current date
                closingBalance = openingBalance + credit - debit;

                // Today's closing balance becomes tomorrow's opening balance
                openingBalance = closingBalance;
            }

            // Add Closing Balance for the last date
            if (currentDate != DateTime.MinValue)
            {
                processedData.Add(new ProcessedData
                {
                    ChoosenDate = currentDate,
                    Head_Id = "",
                    Narration = "Closing Balance",
                    Credit = 0,
                    Debit = closingBalance
                });
            }

            return processedData;
        }


    }

    public class CsvData
    {
        public string ChoosenDate { get; set; }
        public string Head_Id { get; set; }
        public string Narration { get; set; }
        public string Credit { get; set; }
        public string Debit { get; set; }
    }

    public class ProcessedData
    {
        public DateTime ChoosenDate { get; set; }
        public string Head_Id { get; set; }
        public string Narration { get; set; }
        public decimal Credit { get; set; }
        public decimal Debit { get; set; }
    }


}
