using System;
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

        ////[Test]
        //public void SimpleDateFilterTest()
        //{
        //    Random rand = new Random();
        //    DateTime from = new DateTime
        //        (
        //        rand.Next(2020, 2022),
        //        rand.Next(1, 13),
        //        rand.Next(1, 25)
        //        );
        //    DateTime to = from.AddDays(2);
        //    DebtorMessageType type = BankrotClient.SupportedMessageTypes.First();
        //    var messages = BankrotClient.GetMessages(from, to, type);
        //    if (messages.Length != 0)
        //    {
        //        Assert.GreaterOrEqual(messages.Select(x => x.DatePublished.Date).Min(), from);
        //        Assert.LessOrEqual(messages.Select(x => x.DatePublished.Date).Max(), to);
        //    }

        //}
        [TestCase(2021, 3, 1,ExpectedResult = 727)]
        [TestCase(2020, 12, 10, ExpectedResult = 696)]
        [TestCase(2021, 2, 23, ExpectedResult = 163)]
        public int TestGetMessages(int y, int m, int d)
        {
            var type = new DebtorMessageType() { Id = 19, Name = "Реализация имущества должника-банкрота" };
            
            DebtorMessage[] messages = BankrotClient.GetMessages(new DateTime(y, m, d), type);
            return messages.Length;
        }
        [Test]
        public void TestInputInvalidDate()
        {
            var start = new DateTime(2021, 3, 13);
            var end = new DateTime(2021, 3, 12);
            var type = new DebtorMessageType() { Id = 19, Name = "Реализация имущества должника-банкрота" };
            var ex = Assert.Throws<Exception>(() => BankrotClient.GetMessages(start, end, type));

            Assert.That(ex.Message == "Дата конца поиска не может быть раньше даты начала!");

        }
        [Test]
        public void TestInputInvalidInterval()
        {
            var start = new DateTime(2021, 1, 13);
            var end = new DateTime(2021, 3, 12);
            var type = new DebtorMessageType() { Id = 19, Name = "Реализация имущества должника-банкрота" };
            var ex = Assert.Throws<InvalidOperationException>(() => BankrotClient.GetMessages(start, end, type));

            Assert.That(ex.Message == "Максимальная длина интервала - 30 дней!");
        }
        [Test]
        public void TestExportMessagesToExcel()
        {

        }

    }
}