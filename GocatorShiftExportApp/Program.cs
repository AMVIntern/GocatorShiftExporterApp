using System;
using System.IO;
using System.Threading;
using System.Text.Json;
using GocatorShiftExportApp.Models;
using GocatorShiftExportApp.ViewModels;
class Program
{
    static void Main(string[] args)
    {
        // Load configuration from JSON file
        string configPath = @"C:\AMV\Gocator\Config\config.json"; // Adjust path as needed
        EmailData emailData = LoadConfig(configPath);
        if (emailData == null)
        {
            Console.WriteLine("Failed to load configuration. Exiting.");
            return;
        }

        // Initialize ViewModel
        var viewModel = new EmailViewModel(emailData);

        // Run in an infinite loop to wait for scheduled times
        while (true)
        {
            var now = DateTime.Now;
            // Find the next scheduled time
            DateTime nextScheduledTime = GetNextValidScheduledTime(now);

            //// Check if the next scheduled time falls on a Saturday or Sunday
            //while (nextScheduledTime.DayOfWeek == DayOfWeek.Saturday || nextScheduledTime.DayOfWeek == DayOfWeek.Sunday)
            //{
            //    // Move to the next Monday by adding days (6 if Sunday, 1 if Saturday)
            //    //int daysToAdd = nextScheduledTime.DayOfWeek == DayOfWeek.Sunday ? 1 : (8 - (int)nextScheduledTime.DayOfWeek);
            //    //nextScheduledTime = nextScheduledTime.AddDays(daysToAdd);
            //    //nextScheduledTime = GetNextScheduledTime(nextScheduledTime, scheduledTimeSpans);
            //}

            // Calculate initial delay to next scheduled time
            TimeSpan delay = nextScheduledTime - now;
            Console.WriteLine($"Next report scheduled at {nextScheduledTime}. Waiting for {delay.TotalMinutes:F2} minutes...");

            // Update counter every 10 seconds until the scheduled time
            while (now < nextScheduledTime)
            {
                TimeSpan remaining = nextScheduledTime - DateTime.Now;
                Console.Write($"\rRemaining time: {remaining.TotalMinutes:F2} minutes "); // \r for overwrite
                Thread.Sleep(10000); // Update every 10 seconds
                now = DateTime.Now;
            }
            Console.WriteLine(); // New line after countdown

            // Generate combined CSV and send email (shift and date from CSV)
            viewModel.GenerateAndSendReport();
        }
    }
    private static DateTime GetNextValidScheduledTime(DateTime now)
    {
        while (true)
        {
            var day = now.DayOfWeek;

            // Saturday: skip entirely
            if (day == DayOfWeek.Saturday)
            {
                now = now.Date.AddDays(1); // move to Sunday 00:00
                continue;
            }

            // Sunday: only allow 22:00
            if (day == DayOfWeek.Sunday)
            {
                DateTime sundaySlot = now.Date.AddHours(22);

                if (sundaySlot > now)
                    return sundaySlot;

                // Sunday 22:00 already passed → move to Monday
                now = now.Date.AddDays(1);
                continue;
            }

            // Weekdays (Mon–Fri): 06:00, 14:00, 22:00
            TimeSpan[] weekdaySlots =
            {
            new TimeSpan(6, 0, 0),
            new TimeSpan(14, 0, 0),
            new TimeSpan(22, 0, 0)
        };

            foreach (var slot in weekdaySlots)
            {
                DateTime candidate = now.Date + slot;
                if (candidate > now)
                    return candidate;
            }

            // All slots today passed → move to next day
            now = now.Date.AddDays(1);
        }
    }
    private static string GetCurrentShift(DateTime now)
    {
        int hour = now.Hour;
        if (hour >= 6 && hour < 14) return "1"; // 6 AM to 2 PM
        else if (hour >= 14 && hour < 22) return "2"; // 2 PM to 10 PM
        else return "3"; // 10 PM to 6 AM
    }

    private static EmailData LoadConfig(string configPath)
    {
        string jsonString = null; // Declare outside try block for reuse
        try
        {
            if (!File.Exists(configPath))
            {
                // Create directory if it doesn't exist
                string directory = Path.GetDirectoryName(configPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                // Create default configuration
                var defaultConfig = new EmailConfig
                {
                    Settings = new EmailData
                    {
                        FromEmail = "amvgocatorreport@gmail.com",
                        AppPassword = "zoyr xkfl zxlk dhqy",
                        ToEmails = new List<string> { "vikrant@amvco.com.au" },
                        Subject = "AMV Gocator Report",
                        Body = "Please find attached the Gocator Report for {0} corresponding to Shift {1}."
                    }
                };
                jsonString = JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(configPath, jsonString);
                Console.WriteLine($"Created default configuration file at {configPath}.");
            }
            jsonString = File.ReadAllText(configPath); // Reuse jsonString
            var config = JsonSerializer.Deserialize<EmailConfig>(jsonString);
            if (config?.Settings != null)
            {
                // Ensure ToEmails is initialized if null
                if (config.Settings.ToEmails == null)
                {
                    config.Settings.ToEmails = new List<string>();
                }
                return config.Settings;
            }
            Console.WriteLine($"Configuration file invalid at {configPath}.");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading or creating configuration: {ex.Message}");
            return null;
        }
    }
}