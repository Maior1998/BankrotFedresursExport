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
            string source = "https://bankrot.fedresurs.ru/Messages.aspx";

            ChromeOptions chromeOptions = new();

            chromeOptions.AddArgument("--headless");
            chromeOptions.AddArgument("--user-agent=Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25");


            IWebDriver driver = new ChromeDriver(chromeOptions);
            driver.Url = source;
            driver.Navigate();
            Thread.Sleep(1500);



            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            js.ExecuteScript("document.getElementById(\"ctl00_cphBody_mdsMessageType_hfSelectedType\").value = \"\"");
            js.ExecuteScript("document.getElementById(\"ctl00_cphBody_mdsMessageType_hfSelectedValue\").value = \"ArbitralDecree\"");
            js.ExecuteScript("document.getElementById(\"ctl00_cphBody_mdsMessageType_tbSelectedText\").value = \"Сообщение о судебном акте\"");
            //вызвать обновление списка
            js.ExecuteScript("ToggleCourtDesicionType()");
            js.ExecuteScript($"document.getElementById('ctl00_cphBody_ddlCourtDecisionType').value={messageTypeId}");

            string startDateString = start.ToString("dd.MM.yyyy");
            string endDateString = end.ToString("dd.MM.yyyy");

            js.ExecuteScript($"document.getElementById(\"ctl00_cphBody_cldrBeginDate_tbSelectedDate\").value = \'{startDateString}\'");
            js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrBeginDate('{startDateString}')");

            js.ExecuteScript($"document.getElementById(\"ctl00_cphBody_cldrEndDate_tbSelectedDate\").value = \'{endDateString}\'");
            js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrEndDate('{endDateString}')");


            Console.WriteLine();
            IWebElement input = driver.FindElement(By.Id("ctl00_cphBody_ibMessagesSearch"));
            input.Click();
            //waitForAjax(driver);
            Thread.Sleep(1500);



            bool canSearch = !driver.PageSource.Contains(
                "По заданным критериям не найдено ни одной записи. Уточните критерии поиска");
            if (!canSearch)
            {
                //Если поиск не дал результатов
                return new DebtorMessage[0];
            }

            IWebElement pagingLabel = driver.FindElement(By.Id("ctl00_cphBody_PaggingAdvInfo1_tdPaggingAdvInfo"));
            string rawPagingInfo = pagingLabel.GetAttribute("innerText");
            Match pagingInfoMatch = pagingInfoRegex.Match(rawPagingInfo);
            int from = int.Parse(pagingInfoMatch.Groups["start"].Value);
            int to = int.Parse(pagingInfoMatch.Groups["end"].Value);
            int all = int.Parse(pagingInfoMatch.Groups["all"].Value);

            int pageLength = to - from + 1;

            int pagesCount = (int)Math.Ceiling((float)all / pageLength);
            DebtorMessage[] messages = new DebtorMessage[all];
            int curDebtorIndex = 0;
            for (int curPage = 1; curPage <= pagesCount; curPage++)
            {
                if (pagesCount != 1)
                {
                    Console.WriteLine();
                    js.ExecuteScript($"theForm.__EVENTTARGET.value = 'ctl00$cphBody$gvMessages';");
                    js.ExecuteScript($"theForm.__EVENTARGUMENT.value = 'Page${curPage}';");
                    js.ExecuteScript($"theForm.submit();");
                    Thread.Sleep(2000);
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
                    messages[curDebtorIndex] = buffer;
                    curDebtorIndex++;
                }

            }


            FillBirthDate(driver, messages);
            return messages;
        }
        private const string baseMessageViewUrl = @"https://bankrot.fedresurs.ru/MessageWindow.aspx?ID=";
        private static void FillBirthDate(IWebDriver driver, DebtorMessage[] messages)
        {
            foreach (var message in messages)
            {
                string url = $"{baseMessageViewUrl}{message.Guid}";
                driver.Url = url;
                driver.Navigate();
                Thread.Sleep(500);

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
            //Дата публикации
            //Тип сообщения
            //Должник
            //Адрес
            //Кем опубликовано
            //Дата рождения должника

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

        static BankrotClient()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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



        private static Regex dateRegex = new(@"(?<day>\d+)\.(?<month>\d+)\.(?<year>\d{4})");
        private static readonly Regex messagePageHrefRegex = new Regex(@"/MessageWindow\.aspx\?ID=([A-Z0-9]+)");
        private static readonly Regex datetimeRegex = new Regex(@"(?<day>\d+)\.(?<month>\d+)\.(?<year>\d{4}) (?<hours>\d+):(?<minutes>\d+):(?<seconds>\d+)");
        private static readonly Regex pagingInfoRegex = new Regex(@"(?<start>\d+) по (?<end>\d+) \(Всего: (?<all>\d+)");

    }

}
