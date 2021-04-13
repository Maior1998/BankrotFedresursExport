using System;
using System.IO;
using System.Linq;

using BankruptFedresursClient;

using BankruptFedresursModel;

using NUnit.Framework;

namespace Tests
{
    [TestFixture]
    public class Tests
    {
        [SetUp]
        public void Setup()
        { }

        [Test]
        public void SimpleDateFilterTest()
        {
            Random rand = new Random();
            DateTime from = new DateTime
                (
                rand.Next(2020, 2022),
                rand.Next(1, 13),
                rand.Next(1, 25)
                );
            DateTime to = from.AddDays(2);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            var messages = BankrotClient.GetMessages(from, to, type);
            if (messages.Length != 0)
            {
                Assert.GreaterOrEqual(messages.Select(x => x.DatePublished.Date).Min(), from);
                Assert.LessOrEqual(messages.Select(x => x.DatePublished.Date).Max(), to);
            }

        }
        public static string ConvertXlsToXlsx(FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return xlsxFile;
        }
        [TestCase(2021, 3, 1,ExpectedResult = 727)]
        [TestCase(2020, 12, 10, ExpectedResult = 696)]
        [TestCase(2021, 2, 23, ExpectedResult = 163)]
        public int TestGetMessages(int y, int m, int d)
        {
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            
            DebtorMessage[] messages = BankrotClient.GetMessages(new DateTime(y, m, d), type);
            return messages.Length;
        }
        [Test]
        public void TestInputInvalidDate()
        {
            DateTime start = new DateTime(2021, 3, 13);
            DateTime end = new DateTime(2021, 3, 12);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            Exception ex = Assert.Throws<Exception>(() => BankrotClient.GetMessages(start, end, type));

            Assert.That(ex.Message == "ƒата конца поиска не может быть раньше даты начала!");

        }
        [Test]
        public void TestInputInvalidInterval()
        {
            var start = new DateTime(2021, 1, 13);
            var end = new DateTime(2021, 3, 12);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            var ex = Assert.Throws<InvalidOperationException>(() => BankrotClient.GetMessages(start, end, type));
            Assert.That(ex.Message == "ћаксимальна€ длина интервала - 30 дней!");
        }
        [Test]
        public void TestExportMessagesToExcel()
        {
            var date = new DateTime(2021, 1, 13);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            var mess = BankrotClient.GetMessagesWithBirthDates(date, type);

            var excel = BankrotClient.ExportMessagesToExcel(mess);
            
        }

    }
}