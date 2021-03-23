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
        
        static void Main(string[] args)
        {
            DebtorMessage[] messages = BankrotClient.GetMessages(
                DateTime.Today.AddDays(-2),
                DateTime.Today.AddDays(-2),
                BankrotClient.SupportedMessageTypes.First());
            MemoryStream stream = BankrotClient.ExportMessagesToExcel(messages);
            File.WriteAllBytes("output.xlsx", stream.ToArray());
            stream.Close();
            Console.WriteLine();

        }

    }

}
