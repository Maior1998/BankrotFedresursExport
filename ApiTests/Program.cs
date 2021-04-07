using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

using BankruptFedresursClient;

using BankruptFedresursModel;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace ApiTests
{
    class Program
    {

        static void Main(string[] args)
        {
            DateTime testDate = new(2021, 04, 01);
            const string reportFileName = "report.txt";
            if (!File.Exists(reportFileName))
                File.Create(reportFileName).Close();
            File.AppendAllLines(reportFileName, new[] { $"Report from {DateTime.Now:dd.MM.yyyy HH mm}" });
            for (ushort delayInMsec = 3000; delayInMsec >= 1500; delayInMsec -= 200)
            {
                Stopwatch watch = new();
                watch.Start();
                ClientSettings.Settings.MinRequestDelayInMsec = delayInMsec;
                ClientSettings.Settings.MaxRequestDelayInMsec = (ushort)(delayInMsec + 50);
                DebtorMessage[] messages = BankrotClient.GetMessages(
                    testDate,
                    testDate,
                    BankrotClient.SupportedMessageTypes.First());
                BankrotClient.ExportMessagesToExcel(messages);
                watch.Stop();
                double elapsedSeconds = watch.Elapsed.TotalSeconds;
                string result = $"Messages count: {messages.Length}. Delay: {delayInMsec} ms. Time elapsed: {elapsedSeconds} s.";
                Console.WriteLine(result);
                File.AppendAllLines(reportFileName, new[] { result });
                Console.WriteLine($"Waiting before {DateTime.Now.AddHours(1):dd.MM.yyyy HH mm}");
                Thread.Sleep(3600 * 1000);
            }
        }

    }

}
