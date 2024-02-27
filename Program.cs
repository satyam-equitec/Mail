using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string pass = ConfigurationManager.AppSettings["EmailPass"];
            string filePath = @"C:\Users\Equi-Pc3\Downloads\Book1.xlsx";
            Excel.Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Worksheets[1];
            Range range = worksheet.Range["B1:B4"];
            object[,] values =  range.Value;


            List<string> emails = new List<string>();
            foreach (var value in values)
            {
                if (value != null)
                {
                    emails.Add(value.ToString());
                }
                
            }
            foreach (var email in emails)
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("example@example.com");
                mail.To.Add(email);
                mail.Subject = "Test Email";
                mail.Body = "This is a test email";
                SmtpClient smtp = new SmtpClient("smtp.gmail.com");
                smtp.Port = 587;
                smtp.Credentials = new NetworkCredential("example@example.com", pass);
                smtp.EnableSsl = true;
                try
                {
                    smtp.Send(mail);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }

        }

    }
}