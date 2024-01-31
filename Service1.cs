using System;
using System.Collections.Generic;
using System.Configuration;

using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Reflection;
using static System.Net.WebRequestMethods;

namespace BirthdayMailSender
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer;
        public Service1()
        {
            InitializeComponent();
        }


        protected override void OnStart(string[] args)
        {
            try
            {
                timer = new Timer();
                timer.Elapsed += new ElapsedEventHandler(DoWork);
                SetTimerInterval();
                timer.Start();
                WriteToFile("Service started successfully.");
            }
            catch (Exception ex)
            {
                // Log the exception
                WriteToFile($"Error in OnStart: {ex.Message}");
            }
        }

        private void SetTimerInterval()
        {
            
            try
            {
                
                // Get current time
                DateTime now = DateTime.Now;

                // Set the target time (e.g., next 9:00 AM)
                DateTime targetTime = new DateTime(now.Year, now.Month, now.Day, 14, 31, 0);

                // If the target time has already passed for today, set it for tomorrow
                if (now > targetTime)
                {
                    targetTime = targetTime.AddDays(1);
                }

                // Calculate the interval until the target time
                TimeSpan interval = targetTime - now;

                // Ensure the interval is positive
                if (interval.TotalMilliseconds <= 0)
                {
                    // If the interval is zero or negative, set it to a default value (e.g., one day)
                    interval = TimeSpan.FromDays(1);
                }

                // Set the Timer interval
                timer.Interval = interval.TotalMilliseconds;
            }
            catch (Exception ex)
            {
                // Log any exception during interval calculation
                WriteToFile($"Error in SetTimerInterval: {ex.Message}");
            }


        }
        //private void SetTimerInterval()
        //{
        //    // Şuanki saat bilgisini al
        //    DateTime now = DateTime.Now;
        //    DateTime nextNineAm = new DateTime(now.Year, now.Month, now.Day, 12, 51, 0);

        //    if (now > nextNineAm)
        //    {
        //        nextNineAm = nextNineAm.AddDays(1);
        //    }
        //    //// Sabah 8'i belirtilen tarih olarak ayarla
        //    //DateTime morning8 = new DateTime(now.Year, now.Month, now.Day, 8, 0, 0);

        //    //// Akşam 8'i belirtilen tarih olarak ayarla
        //    //DateTime evening8 = new DateTime(now.Year, now.Month, now.Day, 20, 0, 0);

        //    //// Eğer şu anki zaman, bugünün 20:00 (akşam 8) saatinden sonra ise, bir sonraki günün 8:00'ine ayarla
        //    //if (now > evening8)
        //    //{
        //    //    morning8 = morning8.AddDays(1);
        //    //    evening8 = evening8.AddDays(1);
        //    //}

        //    // Timer'ın ilk çalışma zamanını ayarla
        //    double interval = (nextNineAm - now).TotalMilliseconds;
        //    timer.Interval = interval;
        //}


        // Timer'ın periyodunu bir gün olarak ayarla
        //    timer.Interval = (evening8 - morning8).TotalMilliseconds; // Sabah 8 ile akşam 8 arasındaki süre
        //    timer.Start();
        //}


        protected override void OnStop()
        {
            timer.Stop();
            timer.Dispose();
        }
        private void DoWork(object sender, ElapsedEventArgs e)
        {
            WriteToFile("Dowork Method Triggered ");

            try
            {
                string constr = ConfigurationManager.ConnectionStrings["DivanDevConnectionString"].ConnectionString;

                using (var connection = new SqlConnection(constr))
                {
                    WriteToFile("Trying to connect to the database...");

                    connection.Open();
                    WriteToFile("Database connection successful.");

                    var query = "SELECT NameSurname, BDate, mAdress FROM BirthDate WHERE MONTH(BDate) = MONTH(GETDATE()) AND DAY(BDate) = DAY(GETDATE())";

                    using (var command = new SqlCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var NameSurname = reader.GetString(0);
                                var BDate = reader.GetDateTime(1);
                                var mAddress = reader.GetString(2);
                                WriteToFile($"Retrieved data: NameSurname={NameSurname}, BDate={BDate}, mAddress={mAddress}");

                                // Mail gönderme işlemlerini burada gerçekleştirin
                                SendBirthdayMail(NameSurname, BDate, mAddress);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile($"Error in DoWork: {ex.Message}");
            }

            // Timer'ın bir sonraki çalışma zamanını ayarla
            SetTimerInterval();
        }

        //private void DoWork(object sender, ElapsedEventArgs e)
        //{
        //    WriteToFile("Dowork Method Triggered ");

        //    string constr = ConfigurationManager.ConnectionStrings["DivanDevConnectionString"].ConnectionString;
        //    using(var connection = new SqlConnection("DivanDevConnectionString"))

        //    {

        //        WriteToFile("database e bağlanma başarılı ");
        //        connection.Open();

        //        var query = "SELECT NameSurname, BDate, mAdress FROM BirthDate WHERE MONTH(BDate) = MONTH(GETDATE()) AND DAY(BDate) = DAY(GETDATE())";

        //        using (var command = new SqlCommand(query, connection))
        //        {
        //            using (var reader = command.ExecuteReader())
        //            {
        //                while (reader.Read())
        //                {
        //                    var NameSurname = reader.GetString(0);
        //                    var BDate = reader.GetDateTime(1);
        //                    var mAddress = reader.GetString(2);
        //                    WriteToFile($"Retrieved data: NameSurname={NameSurname}, BDate={BDate}, mAddress={mAddress}");
        //                    // Mail gönderme işlemlerini burada gerçekleştirin
        //                    SendBirthdayMail(NameSurname, BDate, mAddress);
        //                }

        //            }
        //        }
        //    }

        //    // Timer'ın bir sonraki çalışma zamanını ayarla
        //    SetTimerInterval();
        //}
        private void SendBirthdayMail(string NameSurname, DateTime BDate, string mAddress)
        {
            WriteToFile("SendBirthdayMail method triggered.");
            string imagepath = "D:\\userdata\\nergis.armagan\\desktop\\background";

            try
            {

                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("10.2.254.24"); // SMTP sunucu bilgisi

                mail.From = new MailAddress("murat.oztin@divan.com.tr"); // Gönderen e-posta adresi
                mail.To.Add(mAddress); 
                mail.Subject = "Happy Birthday!";
                //mail.Body = $"Dear {NameSurname},\n\nHappy Birthday!\n\nBest wishes,\nBirthday Mail Sender Service";
                string body = $@"<html>
                                <body style='background-image:url{imagepath}); background-size:cover;'>
                                    <div> style='text-align:center; color:white; padding:50px;'>
                                    <p>Dear {NameSurname},</p>
                                    <p>Happy Birthday!</p>
                                    <p>Best wishes</p>
                                </body>
                                </html>";
                mail.Body = body;
                mail.IsBodyHtml= true;
                SmtpServer.Port = 25;
                //SmtpServer.Credentials = new NetworkCredential("destek@divan.com.tr"); // GÖNDEREN MAİL ADRESİ
                //SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
                WriteToFile($"Birthday mail sent to {NameSurname} ({mAddress})");
            }
            catch (Exception ex)
            {
                WriteToFile($"Error in SendBirthdayMail: {ex.Message}");
            }
        }
        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            using (System.IO.StreamWriter sw = System.IO.File.AppendText(filepath))
            {
                sw.WriteLine(Message);
            }
        }
    }  
}
