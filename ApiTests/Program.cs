using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        private const int ИндексРеализацииИмущества = 19;
        private const int ИндексРеструктуризацииДолгов = 18;
        static void Main(string[] args)
        {
            DebtorMessage[] messages = BankrotClient.GetMessages(
                DateTime.Today.AddDays(0),
                DateTime.Today.AddDays(1),
                ИндексРеализацииИмущества);
            MemoryStream stream = BankrotClient.ExportMessagesToExcel(messages);
            File.WriteAllBytes("output.xlsx", stream.ToArray());
            stream.Close();
            Console.WriteLine();

        }

    }

}
