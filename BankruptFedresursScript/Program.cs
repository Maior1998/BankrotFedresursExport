using System;
using System.IO;

using BankruptFedresursModel;

namespace BankruptFedresursScript
{
    class Program
    {
        static void Main(string[] args)
        {
            BankruptFedresursClient.BankrotClient.ProgressChanged += BankrotClient_ProgressChanged;
            DateTime dateExport = DateTime.Today.AddDays(-1);
            DebtorMessage[] messages = BankruptFedresursClient.BankrotClient.GetMessagesWithBirthDates(dateExport,
                BankruptFedresursClient.BankrotClient.SupportedMessageTypes);
            MemoryStream stream = BankruptFedresursClient.BankrotClient.ExportMessagesToExcel(messages);
            File.WriteAllBytes($"{DateTime.Now:dd.MM.yyyy HH:mm:ss}.xlsx", stream.ToArray());
        }

        private static void BankrotClient_ProgressChanged(BankruptFedresursClient.ExportStage obj)
        {
            Console.WriteLine($"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {obj.Name}{(obj.AllCount == 0 ? string.Empty : $" {obj.Done} / {obj.AllCount} {obj.Done * 100f / obj.AllCount:P1}")}");
        }
    }
}
