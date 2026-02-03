using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
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
        var gocatorDataExporter = new GocatorDataExporter(emailData);

        // Run in an infinite loop to wait for scheduled times
        while (true)
        {
            var now = DateTime.Now;
            // Find the next scheduled time
            DateTime nextScheduledTime = GetNextValidScheduledTime(now);

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
            gocatorDataExporter.GenerateAndSendReport();

            // Generate combined Excel file with Gocator and Shift data
            var excelGenerator = new CombinedExcelGenerator();
            string excelFilePath = excelGenerator.GenerateCombinedExcelReport();

            // Send email with the combined Excel file
            if (!string.IsNullOrEmpty(excelFilePath) && File.Exists(excelFilePath))
            {
                // Extract shift and date from filename for email subject/body
                string fileName = Path.GetFileNameWithoutExtension(excelFilePath);
                // Format: Combined_Report_Shift_1_28-JAN-2026
                string shift = "Unknown";
                string date = DateTime.Now.ToString("dd-MMM-yyyy");
                
                if (fileName.Contains("Shift_"))
                {
                    int shiftIndex = fileName.IndexOf("Shift_") + 6;
                    int underscoreIndex = fileName.IndexOf("_", shiftIndex);
                    if (underscoreIndex > shiftIndex)
                    {
                        shift = fileName.Substring(shiftIndex, underscoreIndex - shiftIndex);
                    }
                    int lastUnderscore = fileName.LastIndexOf("_");
                    if (lastUnderscore > 0)
                    {
                        date = fileName.Substring(lastUnderscore + 1);
                    }
                }

                // Update email data for combined report
                emailData.AttachmentPath = excelFilePath;
                emailData.Subject = $"AMV Combined Report - Shift {shift} - {date}";
                emailData.Body = $"Please find attached the Combined Report (Gocator + Shift Data) for {date} corresponding to Shift {shift}.";

                // Send email
                SendEmail(emailData);
            }
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

    private static bool SendEmail(EmailData emailData)
    {
        try
        {
            // Create the mail message
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(emailData.FromEmail);
                
                // Add multiple recipients from ToEmails list
                if (emailData.ToEmails != null && emailData.ToEmails.Any())
                {
                    foreach (var toEmail in emailData.ToEmails)
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
                if (emailData.CcEmails != null && emailData.CcEmails.Any())
                {
                    foreach (var ccEmail in emailData.CcEmails)
                    {
                        mail.CC.Add(ccEmail);
                    }
                }

                mail.Subject = emailData.Subject;
                mail.Body = emailData.Body;
                mail.IsBodyHtml = false;

                // Attach the file if it exists
                if (!string.IsNullOrEmpty(emailData.AttachmentPath) && File.Exists(emailData.AttachmentPath))
                {
                    Attachment attachment = new Attachment(emailData.AttachmentPath);
                    mail.Attachments.Add(attachment);
                }
                else if (!string.IsNullOrEmpty(emailData.AttachmentPath))
                {
                    Console.WriteLine("Attachment file not found.");
                    return false;
                }

                // Configure the SMTP client
                using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtpClient.EnableSsl = true;
                    smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new NetworkCredential(emailData.FromEmail, emailData.AppPassword);
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
}