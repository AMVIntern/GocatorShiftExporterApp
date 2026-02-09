using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace GocatorShiftExportApp.ViewModels
{
    public class CombinedExcelGenerator
    {
        private readonly string _combinedFolder;
        private readonly string _s1Folder;
        private readonly string _s2Folder;

        public CombinedExcelGenerator(
            string combinedFolder = @"E:\AMV\GocatorShiftExportApp\Csvs\Combined",
            string s1Folder = @"E:\AMV\GocatorShiftExportApp\Csvs\S1_DLDataLogs",
            string s2Folder = @"E:\AMV\GocatorShiftExportApp\Csvs\S2_DLDataLogs")
        {
            _combinedFolder = combinedFolder;
            _s1Folder = s1Folder;
            _s2Folder = s2Folder;
        }

        public string GenerateCombinedExcelReport()
        {
            try
            {
                // Set EPPlus license context (required for non-commercial use)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Find the most recent combined Gocator CSV file
                string gocatorCsvFile = Directory.GetFiles(_combinedFolder, "Gocator_Report_*.csv")
                    .OrderByDescending(f => File.GetLastWriteTime(f))
                    .FirstOrDefault();

                if (gocatorCsvFile == null)
                {
                    Console.WriteLine("No combined Gocator CSV file found in Combined folder.");
                    return null;
                }

                Console.WriteLine($"Processing Gocator file: {Path.GetFileName(gocatorCsvFile)}");

                // Read combined Gocator CSV
                var gocatorData = ReadCsvFile(gocatorCsvFile);
                if (gocatorData == null || gocatorData.Rows.Count == 0)
                {
                    Console.WriteLine("Gocator CSV file has insufficient data rows.");
                    return null;
                }

                // Extract shift and date from filename or data
                string shift = ExtractShiftFromFile(gocatorCsvFile, gocatorData);
                string date = ExtractDateFromFile(gocatorCsvFile, gocatorData);

                // Find corresponding S1 and S2 shift files
                string s1File = FindShiftFile(_s1Folder, shift, date);
                string s2File = FindShiftFile(_s2Folder, shift, date);

                if (s1File == null && s2File == null)
                {
                    Console.WriteLine($"No shift files found for Shift {shift} and Date {date}.");
                    return null;
                }

                // Read shift files
                ShiftData s1Data = null;
                ShiftData s2Data = null;

                if (s1File != null)
                {
                    Console.WriteLine($"Processing S1 file: {Path.GetFileName(s1File)}");
                    s1Data = ReadShiftFile(s1File, "S1");
                }

                if (s2File != null)
                {
                    Console.WriteLine($"Processing S2 file: {Path.GetFileName(s2File)}");
                    s2Data = ReadShiftFile(s2File, "S2");
                }

                // Calculate timestamps for Gocator data
                CalculateGocatorTimestamps(gocatorData);

                // Calculate timestamps for shift data
                if (s1Data != null)
                {
                    CalculateShiftTimestamps(s1Data);
                }
                if (s2Data != null)
                {
                    CalculateShiftTimestamps(s2Data);
                }

                // Create Excel file
                string excelFileName = Path.Combine(_combinedFolder, $"Combined_Report_Shift_{shift}_{date}.xlsx");
                CreateExcelFile(excelFileName, gocatorData, s1Data, s2Data);

                Console.WriteLine($"Combined Excel file saved to: {excelFileName}");
                return excelFileName;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating combined Excel report: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        private CsvData ReadCsvFile(string filePath)
        {
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                if (lines.Length < 2)
                {
                    Console.WriteLine($"CSV file has insufficient data rows: {filePath}");
                    return null;
                }

                string[] headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();
                var rows = new List<CsvRow>();

                for (int i = 1; i < lines.Length; i++)
                {
                    string[] values = lines[i].Split(',');
                    if (values.Length != headers.Length)
                    {
                        Console.WriteLine($"Row {i} in CSV has {values.Length} columns, expected {headers.Length}. Skipping.");
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
                Console.WriteLine($"Error reading CSV file {filePath}: {ex.Message}");
                return null;
            }
        }

        private ShiftData ReadShiftFile(string filePath, string station)
        {
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                if (lines.Length < 3) // Shift files have 2 header rows
                {
                    Console.WriteLine($"Shift file has insufficient data rows: {filePath}");
                    return null;
                }

                // First line is headers, second line is sub-headers (TLB1, TIB1, etc.), data starts from line 3
                string[] headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();
                string[] subHeaders = null;
                if (lines.Length >= 2)
                {
                    subHeaders = lines[1].Split(',').Select(h => h.Trim()).ToArray();
                    if (subHeaders.Length != headers.Length)
                    {
                        // Pad or trim to match header count
                        var list = new List<string>(subHeaders);
                        while (list.Count < headers.Length) list.Add("");
                        if (list.Count > headers.Length) list = list.Take(headers.Length).ToList();
                        subHeaders = list.ToArray();
                    }
                }

                var rows = new List<ShiftRow>();

                for (int i = 2; i < lines.Length; i++)
                {
                    string[] values = lines[i].Split(',');
                    if (values.Length != headers.Length)
                    {
                        Console.WriteLine($"Row {i} in shift file has {values.Length} columns, expected {headers.Length}. Skipping.");
                        continue;
                    }

                    // Store data by index to handle duplicate column names
                    var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    var rowValues = new string[headers.Length]; // Store values by index
                    
                    for (int j = 0; j < headers.Length; j++)
                    {
                        rowValues[j] = values[j].Trim();
                        // Also store in dictionary for key column lookups (Date, Timestamp, etc.)
                        // For duplicate columns, only store the first occurrence in dictionary
                        if (!rowData.ContainsKey(headers[j]))
                        {
                            rowData[headers[j]] = values[j].Trim();
                        }
                    }

                    rows.Add(new ShiftRow { Data = rowData, Values = rowValues });
                }

                return new ShiftData { Headers = headers, SubHeaders = subHeaders, Rows = rows, Station = station };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading shift file {filePath}: {ex.Message}");
                return null;
            }
        }

        private string FindShiftFile(string folder, string shift, string date)
        {
            if (!Directory.Exists(folder))
                return null;

            // Format: S1_Report_Shift_1_28-Jan-2026.csv or S2_Report_Shift_1_28-Jan-2026.csv
            string pattern = $"*_Report_Shift_{shift}_*.csv";
            var files = Directory.GetFiles(folder, pattern);

            // Try to match by date in filename
            foreach (var file in files)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                if (fileName.Contains(date, StringComparison.OrdinalIgnoreCase))
                {
                    return file;
                }
            }

            // Return first match if no date match found
            return files.FirstOrDefault();
        }

        private string ExtractShiftFromFile(string filePath, CsvData data)
        {
            // Try to extract from filename first
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            // Format: Gocator_Report_Shift_1_28-JAN-2026
            if (fileName.Contains("Shift_"))
            {
                int shiftIndex = fileName.IndexOf("Shift_") + 6;
                int underscoreIndex = fileName.IndexOf("_", shiftIndex);
                if (underscoreIndex > shiftIndex)
                {
                    return fileName.Substring(shiftIndex, underscoreIndex - shiftIndex);
                }
            }

            // Fallback: try to get from data
            if (data != null && data.Rows.Count > 0)
            {
                string shiftCol = FindColumnByName(data.Headers, new[] { "shift" });
                if (!string.IsNullOrEmpty(shiftCol) && data.Rows[0].Data.ContainsKey(shiftCol))
                {
                    return data.Rows[0].Data[shiftCol];
                }
            }

            return "Unknown";
        }

        private string ExtractDateFromFile(string filePath, CsvData data)
        {
            // Try to extract from filename first
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            // Format: Gocator_Report_Shift_1_28-JAN-2026
            int lastUnderscore = fileName.LastIndexOf("_");
            if (lastUnderscore > 0)
            {
                string datePart = fileName.Substring(lastUnderscore + 1);
                return datePart;
            }

            // Fallback: try to get from data
            if (data != null && data.Rows.Count > 0)
            {
                string dateCol = FindColumnByName(data.Headers, new[] { "top:date" });
                if (!string.IsNullOrEmpty(dateCol) && data.Rows[0].Data.ContainsKey(dateCol))
                {
                    return data.Rows[0].Data[dateCol];
                }
            }

            return DateTime.Now.ToString("dd-MMM-yyyy");
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

        private void CalculateGocatorTimestamps(CsvData gocatorData)
        {
            string dateCol = FindColumnByName(gocatorData.Headers, new[] { "top:date" });
            string timestampCol = FindColumnByName(gocatorData.Headers, new[] { "top:timestamp" });

            if (string.IsNullOrEmpty(dateCol) || string.IsNullOrEmpty(timestampCol))
            {
                Console.WriteLine("Could not find Top:Date or Top:Timestamp columns in Gocator CSV.");
                return;
            }

            CalculateFullTimestamps(gocatorData.Rows, dateCol, timestampCol);
        }

        private void CalculateShiftTimestamps(ShiftData shiftData)
        {
            string dateCol = FindColumnByName(shiftData.Headers, new[] { "date" });
            string timestampCol = FindColumnByName(shiftData.Headers, new[] { "timestamp" });

            if (string.IsNullOrEmpty(dateCol) || string.IsNullOrEmpty(timestampCol))
            {
                Console.WriteLine($"Could not find Date or Timestamp columns in {shiftData.Station} shift file.");
                return;
            }

            CalculateFullTimestamps(shiftData.Rows, dateCol, timestampCol);
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
                    if (!DateTime.TryParseExact(dateStr, "dd-MMM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    {
                        if (!DateTime.TryParse(dateStr, out date))
                        {
                            continue; // Skip if date can't be parsed
                        }
                    }
                }

                // Check if timestamp contains AM/PM (12-hour format)
                if (timeStr.Contains("AM", StringComparison.OrdinalIgnoreCase) ||
                    timeStr.Contains("PM", StringComparison.OrdinalIgnoreCase))
                {
                    // Try 12-hour format with AM/PM
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

        private void CalculateFullTimestamps(List<ShiftRow> rows, string dateCol, string timestampCol)
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
                    if (!DateTime.TryParseExact(dateStr, "dd-MMM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    {
                        if (!DateTime.TryParse(dateStr, out date))
                        {
                            continue; // Skip if date can't be parsed
                        }
                    }
                }

                // Shift files use 24-hour format without milliseconds: "hh:mm:ss"
                if (TimeSpan.TryParseExact(timeStr, @"hh\:mm\:ss", CultureInfo.InvariantCulture, out TimeSpan time))
                {
                    row.FullTimestamp = date.Date + time;
                    parsed = true;
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

        private void CreateExcelFile(string filePath, CsvData gocatorData, ShiftData? s1Data, ShiftData? s2Data)
        {
            using (var package = new ExcelPackage())
            {
                // Sheet 1: Gocator Combined Data
                var gocatorSheet = package.Workbook.Worksheets.Add("Gocator_Combined");
                WriteCsvDataToSheet(gocatorSheet, gocatorData);

                // Sheet 2: S1 Shift Data (matched with Gocator)
                // Always create S1 sheet - if data is missing, fill with 0s
                var s1Sheet = package.Workbook.Worksheets.Add("S1_Shift_Data");
                if (s1Data != null)
                {
                    MatchAndWriteShiftData(s1Sheet, gocatorData, s1Data);
                }
                else
                {
                    // If S1 file doesn't exist, try to get headers from a sample file or create default structure
                    WriteEmptyShiftSheet(s1Sheet, gocatorData, "S1");
                }

                // Sheet 3: S2 Shift Data (matched with Gocator)
                // Always create S2 sheet - if data is missing, fill with 0s
                var s2Sheet = package.Workbook.Worksheets.Add("S2_Shift_Data");
                if (s2Data != null)
                {
                    MatchAndWriteShiftData(s2Sheet, gocatorData, s2Data);
                }
                else
                {
                    // If S2 file doesn't exist, try to get headers from a sample file or create default structure
                    WriteEmptyShiftSheet(s2Sheet, gocatorData, "S2");
                }

                // Save the Excel file
                package.SaveAs(new FileInfo(filePath));
            }
        }

        private void WriteCsvDataToSheet(ExcelWorksheet sheet, CsvData data)
        {
            // Write headers
            for (int col = 1; col <= data.Headers.Length; col++)
            {
                sheet.Cells[1, col].Value = data.Headers[col - 1];
            }

            // Write data rows
            for (int row = 0; row < data.Rows.Count; row++)
            {
                for (int col = 1; col <= data.Headers.Length; col++)
                {
                    string header = data.Headers[col - 1];
                    string value = data.Rows[row].Data.ContainsKey(header) ? data.Rows[row].Data[header] : "";
                    sheet.Cells[row + 2, col].Value = value;
                }
            }
        }

        private void WriteEmptyShiftSheet(ExcelWorksheet sheet, CsvData gocatorData, string station)
        {
            // Try to get headers and sub-headers from a sample shift file
            string folder = station == "S1" ? _s1Folder : _s2Folder;
            string[]? headers = null;
            string[]? subHeaders = null;

            if (Directory.Exists(folder))
            {
                var sampleFile = Directory.GetFiles(folder, "*.csv").FirstOrDefault();
                if (sampleFile != null)
                {
                    var sampleData = ReadShiftFile(sampleFile, station);
                    if (sampleData != null && sampleData.Headers != null)
                    {
                        headers = sampleData.Headers;
                        subHeaders = sampleData.SubHeaders;
                    }
                }
            }

            // If no headers found, use default structure
            if (headers == null || headers.Length == 0)
            {
                // Default headers based on typical shift file structure
                if (station == "S2")
                {
                    headers = new[] { "CHEP_PALLET_ID", "Date", "Timestamp", "Shift", "Station" };
                }
                else
                {
                    headers = new[] { "Date", "Timestamp", "Shift", "Station" };
                }
            }

            // Reorder headers: Station, Date, Timestamp, then rest
            var reorderedHeaders = new List<string>();
            var processedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            string stationHeader = headers.FirstOrDefault(h => h.Equals("Station", StringComparison.OrdinalIgnoreCase));
            if (stationHeader != null)
            {
                reorderedHeaders.Add(stationHeader);
                processedHeaders.Add(stationHeader);
            }
            
            string dateHeader = headers.FirstOrDefault(h => h.Equals("Date", StringComparison.OrdinalIgnoreCase));
            if (dateHeader != null)
            {
                reorderedHeaders.Add(dateHeader);
                processedHeaders.Add(dateHeader);
            }
            
            string timestampHeader = headers.FirstOrDefault(h => h.Equals("Timestamp", StringComparison.OrdinalIgnoreCase));
            if (timestampHeader != null)
            {
                reorderedHeaders.Add(timestampHeader);
                processedHeaders.Add(timestampHeader);
            }
            
            foreach (var header in headers)
            {
                if (!processedHeaders.Contains(header))
                {
                    reorderedHeaders.Add(header);
                }
            }

            // Write reordered headers (row 1)
            for (int col = 1; col <= reorderedHeaders.Count; col++)
            {
                sheet.Cells[1, col].Value = reorderedHeaders[col - 1];
            }

            // Write sub-headers (row 2) when available
            if (subHeaders != null && subHeaders.Length > 0)
            {
                var reorderedSubHeaders = GetReorderedSubHeaders(reorderedHeaders, headers, subHeaders);
                for (int col = 1; col <= reorderedSubHeaders.Count; col++)
                {
                    sheet.Cells[2, col].Value = reorderedSubHeaders[col - 1];
                }
            }

            // Identify key columns
            var keyColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Date", "Timestamp", "Shift", "Station", "CHEP_PALLET_ID"
            };

            // Write rows for each Gocator row, filling with 0s - data starts at row 3
            int excelRow = 3;
            foreach (var gocatorRow in gocatorData.Rows)
            {
                for (int col = 1; col <= reorderedHeaders.Count; col++)
                {
                    string header = reorderedHeaders[col - 1];

                    if (keyColumns.Contains(header))
                    {
                        // Fill key columns from Gocator data if available
                        if (header.Equals("Date", StringComparison.OrdinalIgnoreCase))
                        {
                            string gocatorDateCol = FindColumnByName(gocatorData.Headers, new[] { "top:date" });
                            if (!string.IsNullOrEmpty(gocatorDateCol) && gocatorRow.Data.ContainsKey(gocatorDateCol))
                            {
                                sheet.Cells[excelRow, col].Value = gocatorRow.Data[gocatorDateCol];
                            }
                            else
                            {
                                sheet.Cells[excelRow, col].Value = "";
                            }
                        }
                        else if (header.Equals("Timestamp", StringComparison.OrdinalIgnoreCase))
                        {
                            string gocatorTimestampCol = FindColumnByName(gocatorData.Headers, new[] { "top:timestamp" });
                            if (!string.IsNullOrEmpty(gocatorTimestampCol) && gocatorRow.Data.ContainsKey(gocatorTimestampCol))
                            {
                                sheet.Cells[excelRow, col].Value = gocatorRow.Data[gocatorTimestampCol];
                            }
                            else
                            {
                                sheet.Cells[excelRow, col].Value = "";
                            }
                        }
                        else if (header.Equals("Shift", StringComparison.OrdinalIgnoreCase))
                        {
                            string gocatorShiftCol = FindColumnByName(gocatorData.Headers, new[] { "shift" });
                            if (!string.IsNullOrEmpty(gocatorShiftCol) && gocatorRow.Data.ContainsKey(gocatorShiftCol))
                            {
                                sheet.Cells[excelRow, col].Value = gocatorRow.Data[gocatorShiftCol];
                            }
                            else
                            {
                                sheet.Cells[excelRow, col].Value = "";
                            }
                        }
                        else if (header.Equals("Station", StringComparison.OrdinalIgnoreCase))
                        {
                            sheet.Cells[excelRow, col].Value = station;
                        }
                        else
                        {
                            sheet.Cells[excelRow, col].Value = "";
                        }
                    }
                    else
                    {
                        // All data columns - fill with 0
                        sheet.Cells[excelRow, col].Value = "0";
                    }
                }
                excelRow++;
            }
        }

        /// <summary>
        /// Builds sub-headers in the same column order as reorderedHeaders, using original Headers/SubHeaders.
        /// </summary>
        private static List<string> GetReorderedSubHeaders(List<string> reorderedHeaders, string[] headers, string[] subHeaders)
        {
            var result = new List<string>(reorderedHeaders.Count);
            for (int col = 0; col < reorderedHeaders.Count; col++)
            {
                string header = reorderedHeaders[col];
                int headerCount = 0;
                for (int k = 0; k < col; k++)
                {
                    if (string.Equals(reorderedHeaders[k], header, StringComparison.OrdinalIgnoreCase))
                        headerCount++;
                }
                int occurrenceCount = 0;
                int originalIndex = -1;
                for (int idx = 0; idx < headers.Length; idx++)
                {
                    if (string.Equals(headers[idx], header, StringComparison.OrdinalIgnoreCase))
                    {
                        if (occurrenceCount == headerCount)
                        {
                            originalIndex = idx;
                            break;
                        }
                        occurrenceCount++;
                    }
                }
                if (originalIndex >= 0 && originalIndex < subHeaders.Length)
                    result.Add(subHeaders[originalIndex] ?? "");
                else
                    result.Add("");
            }
            return result;
        }

        private void MatchAndWriteShiftData(ExcelWorksheet sheet, CsvData gocatorData, ShiftData shiftData)
        {
            // Sort shift data by timestamp for efficient matching
            var sortedShiftRows = shiftData.Rows
                .Where(r => r.FullTimestamp.HasValue)
                .OrderBy(r => r.FullTimestamp)
                .ToList();

            // Reorder headers: Station, Date, Timestamp, then rest
            var reorderedHeaders = new List<string>();
            var processedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            // Add Station first
            string stationHeader = shiftData.Headers.FirstOrDefault(h => h.Equals("Station", StringComparison.OrdinalIgnoreCase));
            if (stationHeader != null)
            {
                reorderedHeaders.Add(stationHeader);
                processedHeaders.Add(stationHeader);
            }
            
            // Add Date second
            string dateHeader = shiftData.Headers.FirstOrDefault(h => h.Equals("Date", StringComparison.OrdinalIgnoreCase));
            if (dateHeader != null)
            {
                reorderedHeaders.Add(dateHeader);
                processedHeaders.Add(dateHeader);
            }
            
            // Add Timestamp third
            string timestampHeader = shiftData.Headers.FirstOrDefault(h => h.Equals("Timestamp", StringComparison.OrdinalIgnoreCase));
            if (timestampHeader != null)
            {
                reorderedHeaders.Add(timestampHeader);
                processedHeaders.Add(timestampHeader);
            }
            
            // Add rest of the headers in original order
            foreach (var header in shiftData.Headers)
            {
                if (!processedHeaders.Contains(header))
                {
                    reorderedHeaders.Add(header);
                }
            }

            // Write reordered headers (row 1)
            for (int col = 1; col <= reorderedHeaders.Count; col++)
            {
                sheet.Cells[1, col].Value = reorderedHeaders[col - 1];
            }

            // Write sub-headers (row 2: TLB1, TIB1, etc.) in same column order as headers
            if (shiftData.SubHeaders != null && shiftData.SubHeaders.Length > 0)
            {
                var reorderedSubHeaders = GetReorderedSubHeaders(reorderedHeaders, shiftData.Headers, shiftData.SubHeaders);
                for (int col = 1; col <= reorderedSubHeaders.Count; col++)
                {
                    sheet.Cells[2, col].Value = reorderedSubHeaders[col - 1];
                }
            }

            // Identify key columns that should be preserved (not filled with 0)
            var keyColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Date", "Timestamp", "Shift", "Station", "CHEP_PALLET_ID"
            };

            // Match each Gocator row with shift data (Gocator is source of truth) - data starts at row 3
            int excelRow = 3;
            foreach (var gocatorRow in gocatorData.Rows)
            {
                if (!gocatorRow.FullTimestamp.HasValue)
                {
                    // Write row with 0s for data columns if no timestamp
                    for (int col = 1; col <= reorderedHeaders.Count; col++)
                    {
                        string header = reorderedHeaders[col - 1];
                        // Fill data columns with 0, keep key columns empty
                        if (keyColumns.Contains(header))
                        {
                            sheet.Cells[excelRow, col].Value = "";
                        }
                        else
                        {
                            sheet.Cells[excelRow, col].Value = "0";
                        }
                    }
                    excelRow++;
                    continue;
                }

                DateTime gocatorTime = gocatorRow.FullTimestamp.Value;

                // Find the closest shift timestamp that is <= Gocator timestamp (within 10 seconds window)
                ShiftRow matchedRow = null;
                double minDiff = double.MaxValue;

                foreach (var shiftRow in sortedShiftRows)
                {
                    if (!shiftRow.FullTimestamp.HasValue)
                        continue;

                    DateTime shiftTime = shiftRow.FullTimestamp.Value;
                    double diff = (gocatorTime - shiftTime).TotalSeconds;

                    // Shift timestamp should be earlier than or equal to Gocator timestamp
                    // Within 10 seconds window
                    if (diff >= 0 && diff <= 10 && diff < minDiff)
                    {
                        minDiff = diff;
                        matchedRow = shiftRow;
                    }
                }

                // Write matched row or row with 0s for unmatched data
                if (matchedRow != null)
                {
                    // Write matched data - but replace Timestamp with Gocator timestamp
                    for (int col = 1; col <= reorderedHeaders.Count; col++)
                    {
                        string header = reorderedHeaders[col - 1];
                        
                        // Replace Timestamp with Gocator timestamp (Gocator is source of truth)
                        if (header.Equals("Timestamp", StringComparison.OrdinalIgnoreCase))
                        {
                            string gocatorTimestampCol = FindColumnByName(gocatorData.Headers, new[] { "top:timestamp" });
                            if (!string.IsNullOrEmpty(gocatorTimestampCol) && gocatorRow.Data.ContainsKey(gocatorTimestampCol))
                            {
                                sheet.Cells[excelRow, col].Value = gocatorRow.Data[gocatorTimestampCol];
                            }
                            else
                            {
                                sheet.Cells[excelRow, col].Value = "";
                            }
                        }
                        else
                        {
                            // Find the column index in original headers for this reordered header
                            int originalIndex = -1;
                            for (int idx = 0; idx < shiftData.Headers.Length; idx++)
                            {
                                if (string.Equals(shiftData.Headers[idx], header, StringComparison.OrdinalIgnoreCase))
                                {
                                    // For duplicate headers, find the one that matches the reordered position
                                    // Count how many times this header appears before the current position in reordered list
                                    int headerCount = 0;
                                    for (int k = 0; k < col - 1; k++)
                                    {
                                        if (string.Equals(reorderedHeaders[k], header, StringComparison.OrdinalIgnoreCase))
                                        {
                                            headerCount++;
                                        }
                                    }
                                    
                                    // Find the nth occurrence of this header in original headers
                                    int occurrenceCount = 0;
                                    for (int origIdx = 0; origIdx < shiftData.Headers.Length; origIdx++)
                                    {
                                        if (string.Equals(shiftData.Headers[origIdx], header, StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (occurrenceCount == headerCount)
                                            {
                                                originalIndex = origIdx;
                                                break;
                                            }
                                            occurrenceCount++;
                                        }
                                    }
                                    break;
                                }
                            }
                            
                            // Get value from matched row using index
                            string value = "";
                            if (originalIndex >= 0 && matchedRow.Values != null && originalIndex < matchedRow.Values.Length)
                            {
                                value = matchedRow.Values[originalIndex];
                            }
                            else if (matchedRow.Data.ContainsKey(header))
                            {
                                // Fallback to dictionary lookup for key columns
                                value = matchedRow.Data[header];
                            }
                            
                            sheet.Cells[excelRow, col].Value = value;
                        }
                    }
                }
                else
                {
                    // No match found - fill data columns with 0, preserve key columns if possible
                    for (int col = 1; col <= reorderedHeaders.Count; col++)
                    {
                        string header = reorderedHeaders[col - 1];
                        
                        if (keyColumns.Contains(header))
                        {
                            // For key columns, try to get from Gocator or leave empty
                            if (header.Equals("Date", StringComparison.OrdinalIgnoreCase))
                            {
                                string gocatorDateCol = FindColumnByName(gocatorData.Headers, new[] { "top:date" });
                                if (!string.IsNullOrEmpty(gocatorDateCol) && gocatorRow.Data.ContainsKey(gocatorDateCol))
                                {
                                    sheet.Cells[excelRow, col].Value = gocatorRow.Data[gocatorDateCol];
                                }
                                else
                                {
                                    sheet.Cells[excelRow, col].Value = "";
                                }
                            }
                            else if (header.Equals("Timestamp", StringComparison.OrdinalIgnoreCase))
                            {
                                string gocatorTimestampCol = FindColumnByName(gocatorData.Headers, new[] { "top:timestamp" });
                                if (!string.IsNullOrEmpty(gocatorTimestampCol) && gocatorRow.Data.ContainsKey(gocatorTimestampCol))
                                {
                                    sheet.Cells[excelRow, col].Value = gocatorRow.Data[gocatorTimestampCol];
                                }
                                else
                                {
                                    sheet.Cells[excelRow, col].Value = "";
                                }
                            }
                            else if (header.Equals("Shift", StringComparison.OrdinalIgnoreCase))
                            {
                                string gocatorShiftCol = FindColumnByName(gocatorData.Headers, new[] { "shift" });
                                if (!string.IsNullOrEmpty(gocatorShiftCol) && gocatorRow.Data.ContainsKey(gocatorShiftCol))
                                {
                                    sheet.Cells[excelRow, col].Value = gocatorRow.Data[gocatorShiftCol];
                                }
                                else
                                {
                                    sheet.Cells[excelRow, col].Value = "";
                                }
                            }
                            else if (header.Equals("Station", StringComparison.OrdinalIgnoreCase))
                            {
                                // Station is S1 or S2, keep empty or set based on shiftData.Station
                                sheet.Cells[excelRow, col].Value = shiftData.Station ?? "";
                            }
                            else
                            {
                                // Other key columns (like CHEP_PALLET_ID) - leave empty
                                sheet.Cells[excelRow, col].Value = "";
                            }
                        }
                        else
                        {
                            // All data columns (RN, PN, PL, ST, FN, TN, etc.) - fill with 0
                            sheet.Cells[excelRow, col].Value = "0";
                        }
                    }
                }

                excelRow++;
            }
        }

        // Helper classes
        private class ShiftData
        {
            public string[] Headers { get; set; }
            public string[] SubHeaders { get; set; } // Second row: TLB1, TIB1, TIB2, etc.
            public List<ShiftRow> Rows { get; set; } = new List<ShiftRow>();
            public string Station { get; set; }
        }

        private class ShiftRow
        {
            public Dictionary<string, string> Data { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public string[] Values { get; set; } // Store values by index to handle duplicate column names
            public DateTime? FullTimestamp { get; set; }
        }
    }
}
