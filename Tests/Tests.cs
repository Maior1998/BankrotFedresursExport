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
            DebtorMessageType type = BankrotClient.SupportedMessageTypes.First();
            var messages = BankrotClient.GetMessages(from, to, type);
            if (messages.Length != 0)
            {
                Assert.GreaterOrEqual(messages.Select(x => x.DatePublished.Date).Min(), from);
                Assert.LessOrEqual(messages.Select(x => x.DatePublished.Date).Max(), to);
            }
        }
        [Test]
        public void TestGetMessages()
        {
            var start = new DateTime(2021, 3, 1);
            var end = new DateTime(2021, 3, 1);
            var type = new DebtorMessageType(){Id = 19, Name = "Реализация имущества должника-банкрота"};
            var leng = 727;

            DebtorMessage[] messages = BankrotClient.GetMessages(start, end, type);
            Assert.AreEqual(leng, messages.Length);
        }
        [Test]
        public void TestExportMessagesToExcel()
        {

        }

    }
}