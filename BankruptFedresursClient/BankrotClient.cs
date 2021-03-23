using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using System.Threading;
using BankruptFedresursModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace BankruptFedresursClient
{
    public static class BankrotClient
    {

        public static DebtorMessage[] GetMessages(DateTime start, DateTime end, DebtorMessageType type)
        {
            return GetMessages(start, end, type.Id);
        }

        public static DebtorMessage[] GetMessages(DateTime start, DateTime end, int messageTypeId)
        {
            if ((end - start).Days > 30)
            {
                throw new InvalidOperationException("Максимальная длина интервала - 30 дней!");
            }
            string source = "https://bankrot.fedresurs.ru/Messages.aspx";

            ChromeOptions chromeOptions = new();

            //chromeOptions.AddArgument("--headless");
            chromeOptions.AddArgument("--user-agent=Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25");


            IWebDriver driver = new ChromeDriver(chromeOptions);
            driver.Url = source;
            driver.Navigate();
            WebDriverWait wait = new(driver, new TimeSpan(0, 0, 10));
            var element = wait.Until(driver => driver.FindElement(By.Id("ctl00_cphBody_mdsMessageType_hfSelectedType")));
            Thread.Sleep(rand.Next(500, 2000));




            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            js.ExecuteScript("ctl00_cphBody_mdsMessageType_hfSelectedType.value = \"\"");
            js.ExecuteScript("ctl00_cphBody_mdsMessageType_hfSelectedValue.value = \"ArbitralDecree\"");
            js.ExecuteScript("ctl00_cphBody_mdsMessageType_tbSelectedText.value = \"Сообщение о судебном акте\"");
            //вызвать обновление списка
            js.ExecuteScript("ToggleCourtDesicionType()");
            js.ExecuteScript($"ctl00_cphBody_ddlCourtDecisionType.value={messageTypeId}");

            string startDateString = start.ToString("dd.MM.yyyy");
            string endDateString = end.ToString("dd.MM.yyyy");

            js.ExecuteScript($"ctl00_cphBody_cldrBeginDate_tbSelectedDate.value = \'{startDateString}\'");
            js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrBeginDate('{startDateString}')");

            js.ExecuteScript($"ctl00_cphBody_cldrEndDate_tbSelectedDate.value = \'{endDateString}\'");
            js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrEndDate('{endDateString}')");

            Thread.Sleep(rand.Next(500, 2000));
            IWebElement input = driver.FindElement(By.Id("ctl00_cphBody_ibMessagesSearch"));
            input.Click();
            Thread.Sleep(rand.Next(500, 2000));

            

            

            

            
            List<DebtorMessage> messages = new List<DebtorMessage>();

            int curPage = 1;
            do
            {
                
                    Console.WriteLine();
                    js.ExecuteScript($"theForm.__EVENTTARGET.value = 'ctl00$cphBody$gvMessages';");
                    js.ExecuteScript($"theForm.__EVENTARGUMENT.value = 'Page${curPage}';");
                    js.ExecuteScript($"theForm.submit();");
                    WaitForAjax(driver);
                    Thread.Sleep(rand.Next(300, 1001));
                bool isSearchSuccess = !driver.PageSource.Contains(
                "По заданным критериям не найдено ни одной записи. Уточните критерии поиска");
                if (!isSearchSuccess)
                {
                    Thread.Sleep(rand.Next(500, 5000));
                    //Если поиск не дал результатов
                    break;
                }
                
                IWebElement table = driver.FindElement(By.Id("ctl00_cphBody_gvMessages"));
                ReadOnlyCollection<IWebElement> rows = table.FindElements(By.TagName("tr"));
                foreach (IWebElement row in rows)
                {
                    ReadOnlyCollection<IWebElement> columns = row.FindElements(By.TagName("td"));

                    if (columns.Count != 5) continue;
                    DebtorMessage buffer = new();
                    string dateCellValue = columns[0].Text;
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
                    messages.Add(buffer);
                }
                curPage++;
            } while (true);

            DebtorMessage[] resultArray = messages.ToArray();
            FillBirthDate(driver, resultArray);
            return resultArray;
        }

        private static void WaitForAjax(IWebDriver driver)
        {
            while (true) // Handle timeout somewhere
            {
                var ajaxIsComplete = (bool)(driver as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
                if (ajaxIsComplete)
                    break;
                Thread.Sleep(100);
            }
        }

        public static readonly DebtorMessageType[] SupportedMessageTypes = new[]
        {
            new DebtorMessageType()
            {
                Id=19,
                Name="Реализация имущества должника-банкрота"
            },

            new DebtorMessageType()
            {
                Id=18,
                Name="Реструктуризация долгов должника-банкрота"
            },
        };
        private static Random rand = new Random();
        private const string baseMessageViewUrl = @"https://bankrot.fedresurs.ru/MessageWindow.aspx?ID=";
        private static void FillBirthDate(IWebDriver driver, IEnumerable<DebtorMessage> messages)
        {
            foreach (var message in messages)
            {
                string url = $"{baseMessageViewUrl}{message.Guid}";
                driver.Url = url;
                WebDriverWait wait = new(driver, new TimeSpan(0, 0, 10));
                driver.Navigate();
                var element = wait.Until(driver => driver.FindElement(By.ClassName("even")));
                Thread.Sleep(rand.Next(1500, 15000));

                ReadOnlyCollection<IWebElement> rows = driver.FindElements(By.TagName("tr"));

                foreach (IWebElement row in rows)
                {
                    if (!string.IsNullOrWhiteSpace(row.GetAttribute("class")))
                    {
                        ReadOnlyCollection<IWebElement> columns = row.FindElements(By.TagName("td"));
                        if (columns.First().Text != "Дата рождения") continue;
                        Match dateTimeMatch = dateRegex.Match(columns.Last().Text);
                        DateTime date = new DateTime
                            (
                            int.Parse(dateTimeMatch.Groups["year"].Value),
                            int.Parse(dateTimeMatch.Groups["month"].Value),
                            int.Parse(dateTimeMatch.Groups["day"].Value)
                            );
                        message.Debtor.BirthDate = date;
                        break;
                    }
                }
                if (message.Debtor.BirthDate == default)
                {
                    throw new Exception("Дата рождения не найдена!");
                }
            }
            Console.WriteLine();

        }

        public static MemoryStream ExportMessagesToExcel(IEnumerable<DebtorMessage> messages)
        {
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("Выгрузка сообщений");

            for (int i = 0; i < columns.Length; i++)
            {
                sheet.Column(i + 1).Style.Numberformat.Format = columns[i].ColumnFormat;
                sheet.Cells[1, i + 1].Value = columns[i].Header;
            }
            int currentRowIndex = 0;
            foreach (DebtorMessage message in messages)
            {
                for (int currentColumnIndex = 0; currentColumnIndex < columns.Length; currentColumnIndex++)
                {
                    sheet.Cells[currentRowIndex + 2, currentColumnIndex + 1].Value = columns[currentColumnIndex].GetCellValueFromMessage(message);
                }
                currentRowIndex++;
            }
            MemoryStream stream = new MemoryStream();
            excelPackage.SaveAs(stream);
            return stream;
        }

        private static readonly DebtorMessageExcelExportColumn[] columns = new DebtorMessageExcelExportColumn[]
        {
            new DebtorMessageExcelExportColumn(
                "Дата публикации",
                message => message.DatePublished,
                "dd.mm.yyyy"),

            new DebtorMessageExcelExportColumn(
                "Тип сообщения",
                message => message.Type.Name),

            new DebtorMessageExcelExportColumn(
                "Должник",
                message => message.Debtor.FullName),

            new DebtorMessageExcelExportColumn(
                "Адрес",
                message => message.Address),

            new DebtorMessageExcelExportColumn(
                "Кем опубликовано",
                message => message.Owner.FullName),

            new DebtorMessageExcelExportColumn(
                "Дата рождения",
                message => message.Debtor.BirthDate,
                "dd.mm.yyyy"),
        };

        private class DebtorMessageExcelExportColumn
        {
            /// <summary>
            /// Создает новый столбец 
            /// </summary>
            /// <param name="expression"></param>
            public DebtorMessageExcelExportColumn(
                string header,
                Expression<Func<DebtorMessage, object>> expression,
                string columnFormat = "@")
            {
                cellValueFromMessageExpression = expression;
                Header = header;
                ColumnFormat = columnFormat;
            }
            /// <summary>
            /// Формат столбца
            /// </summary>
            public string ColumnFormat { get; set; }
            /// <summary>
            /// Заголовок столбца.
            /// </summary>
            public string Header { get; private set; }
            private Func<DebtorMessage, object> getCellValueFromMessage;
            /// <summary>
            /// Функция, которая вытаскивает значение данного столцба из сообщения по должнику.
            /// </summary>
            public Func<DebtorMessage, object> GetCellValueFromMessage
            => getCellValueFromMessage ??= cellValueFromMessageExpression.Compile();

            /// <summary>
            /// Выражение, которые вытащит из сообщения по должнику нужное поле.
            /// </summary>
            private Expression<Func<DebtorMessage, object>> cellValueFromMessageExpression;

        }

        static BankrotClient()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private static Regex dateRegex = new(@"(?<day>\d+)\.(?<month>\d+)\.(?<year>\d{4})");
        private static readonly Regex messagePageHrefRegex = new Regex(@"/MessageWindow\.aspx\?ID=([A-Z0-9]+)");
        private static readonly Regex datetimeRegex = new Regex(@"(?<day>\d+)\.(?<month>\d+)\.(?<year>\d{4}) (?<hours>\d+):(?<minutes>\d+):(?<seconds>\d+)");
        private static readonly Regex pagingInfoRegex = new Regex(@"(?<start>\d+) по (?<end>\d+) \(Всего: (?<all>\d+)");

    }

}
