using System;
using System.Linq;

using BankruptFedresursClient;

using BankruptFedresursModel;

using NUnit.Framework;

namespace Tests
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
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
            DebtorMessageType type = BankrotClient.SupportedMessageTypes.First();
            DebtorMessage[] messages = BankrotClient.GetMessages(from, to, new[] { type });
            if (messages.Length != 0)
            {
                Assert.GreaterOrEqual(messages.Select(x => x.DatePublished.Date).Min(), from);
                Assert.LessOrEqual(messages.Select(x => x.DatePublished.Date).Max(), to);
            }

        }
    }
}