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
using Cookie = System.Net.Cookie;

namespace ApiTests
{
    class Program
    {

        static void Main(string[] args)
        {
            DateTime date = new(2021, 04, 01);
            File.AppendAllLines("new_delay_tests.log",new []{$"Report from {DateTime.Now:dd.MM.yyyy HH.mm.ss}"});
            
            for (ushort delay = 3000; delay >= 1000; delay -= 100)
            {
                Stopwatch stopwatch = new();
                stopwatch.Start();
                DebtorMessage[] messages = BankrotClient.GetMessagesWithBirthDates(date, BankrotClient.SupportedMessageTypes);
                BankrotClient.ExportMessagesToExcel(messages);
                stopwatch.Stop();
                string bufferResult =
                    $"Delay: {delay} ms. Time: {stopwatch.Elapsed.TotalSeconds} s. Messages: {messages.Length}";
                Console.WriteLine(bufferResult);
                File.AppendAllLines("new_delay_tests.log",new []{bufferResult});
                Thread.Sleep(1800 * 1000);
            }
        }
    }

}
