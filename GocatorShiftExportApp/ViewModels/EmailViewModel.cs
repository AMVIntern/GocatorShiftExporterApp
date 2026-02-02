using GocatorShiftExportApp.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace GocatorShiftExportApp.ViewModels
{
    public class EmailViewModel
    {
        private readonly EmailData _emailData;

        public EmailViewModel(EmailData emailData)
        {
            _emailData = emailData ?? throw new ArgumentNullException(nameof(emailData));
        }

        public void GenerateAndSendReport()
        {
            try
            {
                string topFolder = @"C:\AMV\Gocator\Top";
                string bottomFolder = @"C:\AMV\Gocator\Bottom";
                string combinedFolder = @"C:\AMV\Gocator\Combined";

                // Create combined folder if it doesn't exist
                Directory.CreateDirectory(combinedFolder);

                // Find most recent CSV file containing "values" in Top folder
                string topFile = Directory.GetFiles(topFolder, "*.csv")
                    .Where(f => Path.GetFileName(f).ToLower().Contains("values"))
                    .OrderByDescending(f => File.GetLastWriteTime(f))
                    .FirstOrDefault();

                if (topFile == null)
                {
                    Console.WriteLine("No CSV file containing 'values' found in Top folder.");
                    return;
                }

                // Find most recent CSV file containing "values" in Bottom folder
                string bottomFile = Directory.GetFiles(bottomFolder, "*.csv")
                    .Where(f => Path.GetFileName(f).ToLower().Contains("values"))
                    .OrderByDescending(f => File.GetLastWriteTime(f))
                    .FirstOrDefault();

                if (bottomFile == null)
                {
                    Console.WriteLine("No CSV file containing 'values' found in Bottom folder.");
                    return;
                }

                // Read Top CSV with dynamic column handling
                var topData = ReadCsvFile(topFile, "Top");
                if (topData == null || topData.Rows.Count == 0)
                {
                    Console.WriteLine("Top CSV file has insufficient data rows.");
                    return;
                }

                // Read Bottom CSV with dynamic column handling
                var bottomData = ReadCsvFile(bottomFile, "Bottom");
                if (bottomData == null || bottomData.Rows.Count == 0)
                {
                    Console.WriteLine("Bottom CSV file has insufficient data rows.");
                    return;
                }

                // Find date and timestamp columns dynamically
                string topDateCol = FindColumnByName(topData.Headers, new[] { "top:date" });
                string topTimestampCol = FindColumnByName(topData.Headers, new[] { "top:timestamp" });
                string bottomDateCol = FindColumnByName(bottomData.Headers, new[] { "bot:date" });
                string bottomTimestampCol = FindColumnByName(bottomData.Headers, new[] { "bot:timestamp" });

                if (string.IsNullOrEmpty(topDateCol) || string.IsNullOrEmpty(topTimestampCol) ||
                    string.IsNullOrEmpty(bottomDateCol) || string.IsNullOrEmpty(bottomTimestampCol))
                {
                    Console.WriteLine("Could not find required date/timestamp columns in CSV files.");
                    return;
                }

                // Calculate full timestamps for matching
                CalculateFullTimestamps(topData.Rows, topDateCol, topTimestampCol);
                CalculateFullTimestamps(bottomData.Rows, bottomDateCol, bottomTimestampCol);

                // Sort both lists by FullTimestamp
                topData.Rows = topData.Rows.OrderBy(r => r.FullTimestamp).ToList();
                bottomData.Rows = bottomData.Rows.OrderBy(r => r.FullTimestamp).ToList();

                // Find columns for calculations (look for "Overall" or "Pass" or "Result" patterns)
                string topOverallCol = FindColumnByName(topData.Headers, new[] { "top:overall pass" }, true);
                string bottomOverallCol = FindColumnByName(bottomData.Headers, new[] { "bot:overall_result" }, true);

                // Combine headers - preserve all columns from both files
                var combinedHeaders = new List<string>(topData.Headers);
                // Add bottom headers, avoiding duplicates and excluding Bot:Date and Bot:Timestamp
                string[] excludedColumns = { "Bot:Date", "Bot:Timestamp" };
                foreach (var header in bottomData.Headers)
                {
                    if (!combinedHeaders.Contains(header, StringComparer.OrdinalIgnoreCase) &&
                        !excludedColumns.Contains(header, StringComparer.OrdinalIgnoreCase))
                    {
                        combinedHeaders.Add(header);
                    }
                }
                combinedHeaders.Add("Assured_Result");

                // Prepare combined data
                List<Dictionary<string, string>> combinedRows = new List<Dictionary<string, string>>();

                // Match only rows within 1.5 seconds, ignoring unmatched rows, and calculate Assured_Result
                int i = 0, j = 0;
                while (i < topData.Rows.Count && j < bottomData.Rows.Count)
                {
                    if (!topData.Rows[i].FullTimestamp.HasValue || !bottomData.Rows[j].FullTimestamp.HasValue)
                    {
                        i++;
                        j++;
                        continue;
                    }

                    double diff = Math.Abs((topData.Rows[i].FullTimestamp.Value - bottomData.Rows[j].FullTimestamp.Value).TotalSeconds);
                    if (diff < 1.5) // Match within 1.5 seconds
                    {
                        var combinedRow = new Dictionary<string, string>(topData.Rows[i].Data, StringComparer.OrdinalIgnoreCase);

                        // Merge bottom data, avoiding duplicates
                        foreach (var kvp in bottomData.Rows[j].Data)
                        {
                            if (!combinedRow.ContainsKey(kvp.Key))
                            {
                                combinedRow[kvp.Key] = kvp.Value;
                            }
                        }

                        // Calculate Assured_Result if both overall columns exist
                        if (!string.IsNullOrEmpty(topOverallCol) && !string.IsNullOrEmpty(bottomOverallCol) &&
                            topData.Rows[i].Data.ContainsKey(topOverallCol) && bottomData.Rows[j].Data.ContainsKey(bottomOverallCol))
                        {
                            if (double.TryParse(topData.Rows[i].Data[topOverallCol], NumberStyles.Any, CultureInfo.InvariantCulture, out double topVal) &&
                                double.TryParse(bottomData.Rows[j].Data[bottomOverallCol], NumberStyles.Any, CultureInfo.InvariantCulture, out double bottomVal))
                            {
                                double assuredResult = topVal * bottomVal;
                                combinedRow["Assured_Result"] = assuredResult.ToString(CultureInfo.InvariantCulture);
                            }
                        }
                        else
                        {
                            combinedRow["Assured_Result"] = "";
                        }

                        combinedRows.Add(combinedRow);
                        i++;
                        j++;
                    }
                    else if (topData.Rows[i].FullTimestamp < bottomData.Rows[j].FullTimestamp)
                    {
                        i++; // Skip unmatched Top row
                    }
                    else
                    {
                        j++; // Skip unmatched Bottom row
                    }
                }

                if (combinedRows.Count == 0)
                {
                    Console.WriteLine("No matching rows found between Top and Bottom CSV files.");
                    return;
                }

                // Save combined CSV
                string shiftCol = FindColumnByName(combinedHeaders, new[] { "shift" });
                string dateCol = FindColumnByName(combinedHeaders, new[] { "top:date" });

                string shiftValue = combinedRows[0].ContainsKey(shiftCol) ? combinedRows[0][shiftCol] : "Unknown";
                string dateValue = combinedRows[0].ContainsKey(dateCol) ? combinedRows[0][dateCol] : DateTime.Now.ToString("dd-MMM-yyyy");

                string combinedFile = Path.Combine(combinedFolder, $"Gocator_Report_Shift_{shiftValue}_{dateValue}.csv");

                using (StreamWriter writer = new StreamWriter(combinedFile))
                {
                    // Write header
                    writer.WriteLine(string.Join(",", combinedHeaders));

                    // Write rows
                    foreach (var row in combinedRows)
                    {
                        var rowValues = combinedHeaders.Select(h => row.ContainsKey(h) ? row[h] : "").ToArray();
                        writer.WriteLine(string.Join(",", rowValues));
                    }
                }
                Console.WriteLine($"Combined CSV saved to: {combinedFile}");

                // Extract date and shift for email
                string combinedShift = shiftValue;
                string combinedDate = dateValue;

                // Set attachment path and update email content with combined file data
                _emailData.AttachmentPath = combinedFile;
                _emailData.Subject = string.Format(_emailData.Subject, combinedDate, combinedShift);
                _emailData.Body = $"Please find attached the Gocator Report for {combinedDate} corresponding to Shift {combinedShift}.";
                SendEmail();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating report: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private CsvData ReadCsvFile(string filePath, string sourceName)
        {
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                if (lines.Length < 2)
                {
                    Console.WriteLine($"{sourceName} CSV file has insufficient data rows.");
                    return null;
                }

                string[] headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();
                var rows = new List<CsvRow>();

                for (int i = 1; i < lines.Length; i++)
                {
                    string[] values = lines[i].Split(',');
                    if (values.Length != headers.Length)
                    {
                        Console.WriteLine($"Row {i} in {sourceName} CSV has {values.Length} columns, expected {headers.Length}. Skipping.");
                        continue;
                    }

                    var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int j = 0; j < headers.Length; j++)
                    {
                        rowData[headers[j]] = values[j].Trim();
                    }

                    rows.Add(new CsvRow { Data = rowData });
                }

                return new CsvData { Headers = headers, Rows = rows };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading {sourceName} CSV file: {ex.Message}");
                return null;
            }
        }

        private string FindColumnByName(string[] headers, string[] possibleNames, bool containsMatch = false)
        {
            foreach (var header in headers)
            {
                string headerLower = header.ToLower();
                foreach (var name in possibleNames)
                {
                    if (containsMatch)
                    {
                        if (headerLower.Contains(name.ToLower()))
                            return header;
                    }
                    else
                    {
                        if (headerLower.Equals(name.ToLower()) || headerLower.Replace(":", "").Equals(name.ToLower()))
                            return header;
                    }
                }
            }
            return null;
        }

        private string FindColumnByName(List<string> headers, string[] possibleNames, bool containsMatch = false)
        {
            return FindColumnByName(headers.ToArray(), possibleNames, containsMatch);
        }

        private void CalculateFullTimestamps(List<CsvRow> rows, string dateCol, string timestampCol)
        {
            foreach (var row in rows)
            {
                if (!row.Data.ContainsKey(dateCol) || !row.Data.ContainsKey(timestampCol))
                    continue;

                string dateStr = row.Data[dateCol];
                string timeStr = row.Data[timestampCol].Trim();

                bool parsed = false;

                // Try parsing date first
                DateTime date;
                if (!DateTime.TryParseExact(dateStr, "dd-MMM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    if (!DateTime.TryParse(dateStr, out date))
                    {
                        continue; // Skip if date can't be parsed
                    }
                }

                // Check if timestamp contains AM/PM (12-hour format)
                if (timeStr.Contains("AM", StringComparison.OrdinalIgnoreCase) ||
                    timeStr.Contains("PM", StringComparison.OrdinalIgnoreCase))
                {
                    // Try 12-hour format with AM/PM: "hh:mm:ss tt" or "h:mm:ss tt"
                    string[] formats12Hour = {
                    "h:mm:ss tt",
                    "hh:mm:ss tt",
                    "h:mm:ss.fff tt",
                    "hh:mm:ss.fff tt"
                };

                    foreach (var format in formats12Hour)
                    {
                        if (DateTime.TryParseExact(timeStr, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime timeOnly))
                        {
                            row.FullTimestamp = date.Date + timeOnly.TimeOfDay;
                            parsed = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Try 24-hour format with milliseconds: "hh:mm:ss.fff"
                    if (TimeSpan.TryParseExact(timeStr, @"hh\:mm\:ss\.fff", CultureInfo.InvariantCulture, out TimeSpan time))
                    {
                        row.FullTimestamp = date.Date + time;
                        parsed = true;
                    }
                    // Try 24-hour format without milliseconds: "hh:mm:ss"
                    else if (TimeSpan.TryParseExact(timeStr, @"hh\:mm\:ss", CultureInfo.InvariantCulture, out time))
                    {
                        row.FullTimestamp = date.Date + time;
                        parsed = true;
                    }
                }

                // Fallback: Try generic TimeSpan parsing
                if (!parsed && TimeSpan.TryParse(timeStr, out TimeSpan fallbackTime))
                {
                    row.FullTimestamp = date.Date + fallbackTime;
                    parsed = true;
                }

                // If still not parsed, try parsing as full DateTime string
                if (!parsed)
                {
                    string combinedDateTime = $"{dateStr} {timeStr}";
                    if (DateTime.TryParse(combinedDateTime, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fullDateTime))
                    {
                        row.FullTimestamp = fullDateTime;
                    }
                }
            }
        }

        public bool SendEmail()
        {
            try
            {
                // Create the mail message
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(_emailData.FromEmail); // Corrected to FromEmail
                                                                       // Add multiple recipients from ToEmails list
                    if (_emailData.ToEmails != null)
                    {
                        foreach (var toEmail in _emailData.ToEmails)
                        {
                            mail.To.Add(toEmail);
                        }
                    }
                    else
                    {
                        Console.WriteLine("No recipients specified.");
                        return false;
                    }
                    // CC
                    if (_emailData.CcEmails != null && _emailData.CcEmails.Any())
                    {
                        foreach (var ccEmail in _emailData.CcEmails)
                        {
                            mail.CC.Add(ccEmail);
                        }
                    }
                    mail.Subject = _emailData.Subject;
                    mail.Body = _emailData.Body;
                    mail.IsBodyHtml = false;

                    // Attach the file if it exists
                    if (!string.IsNullOrEmpty(_emailData.AttachmentPath) && File.Exists(_emailData.AttachmentPath))
                    {
                        Attachment attachment = new Attachment(_emailData.AttachmentPath);
                        mail.Attachments.Add(attachment);
                    }
                    else if (!string.IsNullOrEmpty(_emailData.AttachmentPath))
                    {
                        Console.WriteLine("Attachment file not found.");
                        return false;
                    }

                    // Configure the SMTP client
                    using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587))
                    {
                        smtpClient.EnableSsl = true;
                        smtpClient.UseDefaultCredentials = false;
                        smtpClient.Credentials = new NetworkCredential(_emailData.FromEmail, _emailData.AppPassword);
                        smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                        // Send the email
                        smtpClient.Send(mail);
                        Console.WriteLine("Email sent successfully!");
                        return true;
                    }
                }
            }
            catch (SmtpException ex)
            {
                Console.WriteLine($"SMTP Error: {ex.Message}");
                Console.WriteLine($"Status Code: {ex.StatusCode}");
                if (ex.InnerException != null)
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General Error: {ex.Message}");
                return false;
            }
        }

        private static string GetCurrentShift(DateTime now)
        {
            int hour = now.Hour;
            if (hour >= 6 && hour < 14) return "1"; // 6 AM to 2 PM
            else if (hour >= 14 && hour < 22) return "2"; // 2 PM to 10 PM
            else return "3"; // 10 PM to 6 AM
        }
    }

    // Column-agnostic data structures
    public class CsvData
    {
        public string[] Headers { get; set; }
        public List<CsvRow> Rows { get; set; } = new List<CsvRow>();
    }

    public class CsvRow
    {
        public Dictionary<string, string> Data { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public DateTime? FullTimestamp { get; set; }
    }
}
