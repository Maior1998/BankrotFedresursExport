using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
            DateTime startDateTime = DateTime.Today.AddDays(-7);
            DateTime endDateTime = DateTime.Today.AddDays(1);
            string source = "https://bankrot.fedresurs.ru/Messages.aspx";

            ChromeOptions chromeOptions = new();

            //chromeOptions.AddArgument("--headless");
            chromeOptions.AddArgument("--user-agent=Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25");


            IWebDriver driver = new ChromeDriver(chromeOptions);
            driver.Url = source;
            driver.Navigate();
            Thread.Sleep(1500);
            //клик по типу сообщения

            driver.Manage().Timeouts().ImplicitWait = new TimeSpan(0, 0, 10);
            IWebElement input = driver.FindElement(By.Id("ctl00_cphBody_mdsMessageType_tbSelectedText"));
            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            js.ExecuteScript("document.getElementById(\"ctl00_cphBody_mdsMessageType_hfSelectedType\").value = \"\"");
            js.ExecuteScript("document.getElementById(\"ctl00_cphBody_mdsMessageType_hfSelectedValue\").value = \"ArbitralDecree\"");
            js.ExecuteScript("document.getElementById(\"ctl00_cphBody_mdsMessageType_tbSelectedText\").value = \"Сообщение о судебном акте\"");
            //вызвать обновление списка
            js.ExecuteScript("ToggleCourtDesicionType()");
            js.ExecuteScript($"document.getElementById('ctl00_cphBody_ddlCourtDecisionType').value={ИндексРеструктуризацииДолгов}");

            string startDateString = startDateTime.ToString("dd.MM.yyyy");
            string endDateString = endDateTime.ToString("dd.MM.yyyy");

            js.ExecuteScript($"document.getElementById(\"ctl00_cphBody_cldrBeginDate_tbSelectedDate\").value = \'{startDateString}\'");
            js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrBeginDate('{startDateString}')");

            js.ExecuteScript($"document.getElementById(\"ctl00_cphBody_cldrEndDate_tbSelectedDate\").value = \'{endDateString}\'");
            js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrEndDate('{endDateString}')");


            Console.WriteLine();
            input = driver.FindElement(By.Id("ctl00_cphBody_ibMessagesSearch"));
            input.Click();
            //waitForAjax(driver);
            Thread.Sleep(1500);

            bool canSearch = !driver.PageSource.Contains(
                "По заданным критериям не найдено ни одной записи. Уточните критерии поиска");
            if (!canSearch)
            {
                //Если поиск не дал результатов
                return;
            }

            IWebElement table = driver.FindElement(By.Id("ctl00_cphBody_gvMessages"));
            IWebElement pagingLabel = driver.FindElement(By.Id("ctl00_cphBody_PaggingAdvInfo1_tdPaggingAdvInfo"));
            string rawPagingInfo = pagingLabel.GetAttribute("innerText");
            Match pagingInfoMatch = pagingInfoRegex.Match(rawPagingInfo);
            int from = int.Parse(pagingInfoMatch.Groups["start"].Value);
            int to = int.Parse(pagingInfoMatch.Groups["end"].Value);
            int all = int.Parse(pagingInfoMatch.Groups["all"].Value);

            int pageLength = to - from + 1;

            int pagesCount = (int)Math.Ceiling((float)all / pageLength);
            List<DebtorMessage> debtors = new List<DebtorMessage>();
            for (int curPage = 1; curPage <= pagesCount; curPage++)
            {
                if (pagesCount != 1)
                {
                    Console.WriteLine();
                    js.ExecuteScript($"javascript:__doPostBack('ctl00$cphBody$gvMessages','Page${curPage}')");
                    Thread.Sleep(2000);
                }
                table = driver.FindElement(By.Id("ctl00_cphBody_gvMessages"));
                ReadOnlyCollection<IWebElement> rows = table.FindElements(By.TagName("tr"));
                foreach (IWebElement row in rows)
                {
                    ReadOnlyCollection<IWebElement> columns = row.FindElements(By.TagName("td"));
                    
                    if (columns.Count == 0) continue;
                    DebtorMessage buffer = new();
                    Match dateTimeRegexMatch = datetimeRegex.Match(columns[0].Text);

                    buffer.DatePublished = new DateTime
                        (
                        int.Parse(dateTimeRegexMatch.Groups["year"].Value),
                        int.Parse(dateTimeRegexMatch.Groups["month"].Value),
                        int.Parse(dateTimeRegexMatch.Groups["day"].Value),
                        int.Parse(dateTimeRegexMatch.Groups["hours"].Value),
                        int.Parse(dateTimeRegexMatch.Groups["minutes"].Value),
                        int.Parse(dateTimeRegexMatch.Groups["seconds"].Value)
                        );

                    buffer.Type = new DebtorMessageType() { Name = columns[1].Text };
                    string messagePageLink = columns[1].FindElement(By.TagName("a")).GetAttribute("href");
                    buffer.Guid = messagePageHrefRegex.Match(messagePageLink).Groups[1].Value;
                    buffer.Debtor = new Debtor() { FullName = columns[2].Text };
                    
                    buffer.Address = columns[3].Text;
                    buffer.Owner = new ArbitrManager() { FullName = columns[4].Text };
                    debtors.Add(buffer);
                }
                Console.WriteLine();
            }



            Console.WriteLine();

        }
        private static readonly Regex messagePageHrefRegex = new Regex(@"/MessageWindow\.aspx\?ID=([A-Z0-9]+)");
        private static readonly Regex datetimeRegex = new Regex(@"(?<day>\d+)\.(?<month>\d+)\.(?<year>\d{4}) (?<hours>\d+):(?<minutes>\d+):(?<seconds>\d+)");
        private static readonly Regex pagingInfoRegex = new Regex(@"(?<start>\d+) по (?<end>\d+) \(Всего: (?<all>\d+)");
        private static void waitForAjax(IWebDriver driver)
        {
            bool isAjaxStillLoading = true;
            while (isAjaxStillLoading)
            {
                isAjaxStillLoading = !(bool)((IJavaScriptExecutor)driver).ExecuteScript("return jQuery.active == 0");
                Thread.Sleep(150);
            }
        }
    }

}
