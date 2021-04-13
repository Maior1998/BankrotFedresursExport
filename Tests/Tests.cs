using System;
using System.IO;
using System.Linq;

using BankruptFedresursClient;

using BankruptFedresursModel;
using OfficeOpenXml;
using NUnit.Framework;
using System.Xml;

namespace Tests
{
    [TestFixture]
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Test]
        public void SimpleDateFilterTest()
        {
            Random rand = new();
            DateTime from = new(
                rand.Next(2020, 2022),
                rand.Next(1, 13),
                rand.Next(1, 25)
                );
            DateTime to = from.AddDays(2);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            DebtorMessage[] messages = BankrotClient.GetMessages(from, to, type);
            if (messages.Length != 0)
            {
                Assert.GreaterOrEqual(messages.Select(x => x.DatePublished.Date).Min(), from);
                Assert.LessOrEqual(messages.Select(x => x.DatePublished.Date).Max(), to);
            }

        }
        [TestCase(2021, 3, 1,ExpectedResult = 727)]
        [TestCase(2020, 12, 10, ExpectedResult = 696)]
        [TestCase(2021, 2, 23, ExpectedResult = 163)]
        public int TestGetMessages(int year, int month, int day)
        {
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            
            DebtorMessage[] messages = BankrotClient.GetMessages(new DateTime(year, month, day), type);
            return messages.Length;
        }
        [Test]
        public void TestInputInvalidDate()
        {
            DateTime start = new(2021, 3, 13);
            DateTime end = new(2021, 3, 12);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            Exception ex = Assert.Throws<Exception>(() => BankrotClient.GetMessages(start, end, type));

            Assert.That(ex.Message == "���� ����� ������ �� ����� ���� ������ ���� ������!");

        }
        [Test]
        public void TestInputInvalidInterval()
        {
            DateTime start = new(2021, 1, 13);
            DateTime end = new(2021, 3, 12);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => BankrotClient.GetMessages(start, end, type));
            Assert.That(ex.Message == "������������ ����� ��������� - 30 ����!");
        }

        [TestCase(2, ExpectedResult = "09.09.1958")]
        public object TestExportMessagesToExcel(int indexRow)
        {
            DateTime date = new(2021, 1, 1);
            DebtorMessageType[] type = { BankrotClient.SupportedMessageTypes.First(x => x.Id == 19) };
            DebtorMessage[] mess = BankrotClient.GetMessagesWithBirthDates(date, type);

            MemoryStream memoryStreamExcel = BankrotClient.ExportMessagesToExcel(mess);
            ExcelPackage excelFile = new(memoryStreamExcel);
            string buf = (excelFile.Workbook.Worksheets.First().Cells[indexRow, 4].Text);
            return buf;
        }
        [Test]
        public void Test()
        {
            ExcelPackage excelFile = new(new FileInfo("C:\\Users\\aidan\\Desktop\\test.xlsx"));
            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets["����1"];
            ExcelCellAddress start = worksheet.Dimension.Start;
            ExcelCellAddress end = worksheet.Dimension.End;
            int count = 0;
            for (int i = start.Row; i <= end.Row; i++)
            {
                object cellValue = worksheet.Cells[i, 2].Text;
                if (cellValue.ToString() == "��������� � �������� ���� (������������)")
                    count++;
            }
            Assert.AreEqual("��������� � �������� ����", worksheet.Cells["B2"]);
            Assert.AreEqual(3, count);
        }
    }
}