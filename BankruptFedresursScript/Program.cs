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
            DateTime startDate = DateTime.Today.AddDays(-ScriptSettings.Settings.DateStartOffset);
            DateTime endDate = DateTime.Today.AddDays(-ScriptSettings.Settings.DateEndOffset);
            Console.WriteLine($"Начало выгрузки от {startDate:dd.MM.yyyy} до {endDate:dd.MM.yyyy}");
            DebtorMessage[] messages = BankruptFedresursClient.BankrotClient.GetMessagesWithBirthDates(
                startDate,
                endDate,
                BankruptFedresursClient.BankrotClient.SupportedMessageTypes);
            MemoryStream stream = BankruptFedresursClient.BankrotClient.ExportMessagesToExcel(messages);
            File.WriteAllBytes(
                Path.Combine(ScriptSettings.Settings.ExcelExportFilePath,
                    $"Выгрузка сообщений от {DateTime.Now:dd.MM.yyyy HH mm ss} [{startDate:dd.MM.yyyy}-{endDate:dd.MM.yyyy}].xlsx"),
                stream.ToArray());
        }

        private static void BankrotClient_ProgressChanged(BankruptFedresursClient.ExportStage obj)
        {
            Console.WriteLine($"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {obj.Name}{(obj.AllCount == 0 ? string.Empty : $" {obj.Done} / {obj.AllCount} {(float)obj.Done / obj.AllCount:P1}")}");
        }
    }
}
