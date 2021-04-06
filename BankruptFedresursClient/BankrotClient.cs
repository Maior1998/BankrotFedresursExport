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
using OfficeOpenXml.Table;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace BankruptFedresursClient
{
    public static class BankrotClient
    {
        private static CancellationToken cancellationToken = new();
        private static IWebDriver driver;
        private static bool isLoading;

        public static event Action<ExportStage> ProgressChanged;
        public static void SetCancellationToken(CancellationToken token)
        {
            cancellationToken = token;
        }

        public static DebtorMessage[] GetMessagesWithBirthDates(DateTime date, DebtorMessageType[] type)
        {
            return GetMessagesWithBirthDates(date, date, type);
        }

        public static DebtorMessage[] GetMessagesWithBirthDates(DateTime startDate, DateTime endDate,
            DebtorMessageType[] type)
        {
            return GetMessagesWithBirthDates(startDate, endDate, type.Select(x => x.Id).ToArray());
        }

        public static DebtorMessage[] GetMessagesWithBirthDates(DateTime startDate, DateTime endDate, int[] type)
        {
            DebtorMessage[] messages = GetMessages(startDate, endDate, type);
            FillBirthDate(messages);
            return messages;
        }

        /// <summary>
        /// Получает массив сообщений по должникам при помощи фильтров:
        /// день публикации сообщения,
        /// а также фильтр по типу опубликуемого арбитражным управляющим сообщения.
        /// </summary>
        /// <param name="date">День публикации сообщения.</param>
        /// <param name="type">Объект типа сообщения суда, которое необходимо вытащить.</param>
        /// <returns>Массив сообщений по должникам с учетом фильтров по дате публикации и типу сообщения.</returns>
        public static DebtorMessage[] GetMessages(DateTime date, DebtorMessageType[] type)
        {
            return GetMessages(date, type.Select(x => x.Id).ToArray());
        }

        /// <summary>
        /// Получает массив сообщений по должникам при помощи фильтров:
        /// день публикации сообщения,
        /// а также фильтр по типу опубликуемого арбитражным управляющим сообщения.
        /// </summary>
        /// <param name="date">День публикации сообщения.</param>
        /// <param name="type">Индекс типа сообщения суда, которое необходимо вытащить.</param>
        /// <returns>Массив сообщений по должникам с учетом фильтров по дате публикации и типу сообщения.</returns>
        public static DebtorMessage[] GetMessages(DateTime date, int[] messageTypeId)
        {
            return GetMessages(date, date, messageTypeId);
        }

        /// <summary>
        /// Получает массив сообщений по должникам при помощи фильтров:
        /// интервал публикации сообщения (дата-начало, дата-конец),
        /// а также фильтр по типу опубликуемого арбитражным управляющим сообщения.
        /// </summary>
        /// <param name="start">Дата-начало фильтра по дате публикации сообщения.</param>
        /// <param name="end">Дата-конец фильтра по дате публикации сообщения.</param>
        /// <param name="type">Объект типа сообщения суда, которое необходимо вытащить.</param>
        /// <returns>Массив сообщений по должникам с учетом фильтров по дате публикации и типу сообщения.</returns>
        public static DebtorMessage[] GetMessages(DateTime start, DateTime end, DebtorMessageType[] type)
        {
            return GetMessages(start, end, type.Select(x => x.Id).ToArray());
        }

        private static IWebDriver GetDriver()
        {
            ChromeDriverService chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;
            ChromeOptions chromeOptions = new();
            //chromeOptions.AddArgument("--headless");

            chromeOptions.AddArgument(
                $"--user-agent={ClientSettings.Settings.UserAgent}");
            return new ChromeDriver(chromeDriverService, chromeOptions);
        }


        /// <summary>
        /// Получает массив сообщений по должникам при помощи фильтров:
        /// интервал публикации сообщения (дата-начало, дата-конец),
        /// а также фильтр по типу опубликуемого арбитражным управляющим сообщения.
        /// </summary>
        /// <param name="start">Дата-начало фильтра по дате публикации сообщения.</param>
        /// <param name="end">Дата-конец фильтра по дате публикации сообщения.</param>
        /// <param name="messageTypeIds">Номера типов сообщения суда, которые необходимо вытащить.</param>
        /// <returns>Массив сообщений по должникам с учетом фильтров по дате публикации и типу сообщения.</returns>
        public static DebtorMessage[] GetMessages(DateTime start, DateTime end, int[] messageTypeIds)
        {
            if (isLoading)
                throw new InvalidOperationException("Already Loading");
            ProgressChanged?.Invoke(new ExportStage()
            {
                Name = "Подготовка"
            });
            isLoading = true;
            if (end < start)
            {
                throw new Exception("Дата конца поиска не может быть раньше даты начала!");
            }
            if ((end - start).Days > 30)
            {
                throw new InvalidOperationException("Максимальная длина интервала - 30 дней!");
            }
            string source = "https://bankrot.fedresurs.ru/Messages.aspx";
            List<DebtorMessage> resultList = new();
            foreach (int messageTypeId in messageTypeIds)
            {

                try
                {
                    driver = GetDriver();
                    driver.Url = source;
                    driver.Navigate();
                    WebDriverWait wait = new(driver, new TimeSpan(0, 0, 10));
                    wait.Until(driver => driver.FindElement(By.Id("ctl00_cphBody_mdsMessageType_hfSelectedType")));
                    Thread.Sleep(rand.Next(500, 2000));
                    //через js выбираем тип сообщения
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    js.ExecuteScript("ctl00_cphBody_mdsMessageType_hfSelectedType.value = \"\"");
                    js.ExecuteScript("ctl00_cphBody_mdsMessageType_hfSelectedValue.value = \"ArbitralDecree\"");
                    js.ExecuteScript(
                        "ctl00_cphBody_mdsMessageType_tbSelectedText.value = \"Сообщение о судебном акте\"");
                    //вызываем обновление типа решения суда, чтобы появился combobox
                    //(не уверен насколько это необходимо, но на всякий случай, так как это
                    //происходит при штатной работе пользователя с программой)
                    js.ExecuteScript("ToggleCourtDesicionType()");
                    //выбираем тип судебного акта тем, который нам дали на вход
                    js.ExecuteScript($"ctl00_cphBody_ddlCourtDecisionType.value={messageTypeId}");

                    //переводим даты в вид, который принимает форма заполнения дат на странице.
                    string startDateString = start.ToString("dd.MM.yyyy");
                    string endDateString = end.ToString("dd.MM.yyyy");

                    //выбираем дату начала как в самом "выборщике" дат, так и в поле позадни него
                    js.ExecuteScript($"ctl00_cphBody_cldrBeginDate_tbSelectedDate.value = \'{startDateString}\'");
                    js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrBeginDate('{startDateString}')");
                    //выбираем дату конца как в самом "выборщике" дат, так и в поле позадни него
                    js.ExecuteScript($"ctl00_cphBody_cldrEndDate_tbSelectedDate.value = \'{endDateString}\'");
                    js.ExecuteScript($"SetHiddenField_ctl00_cphBody_cldrEndDate('{endDateString}')");
                    //ждем произвольный интервал времени, чтобы изменения применились
                    Thread.Sleep(rand.Next(500, 2000));
                    //находим кнопку поиска и кликаем по ней
                    IWebElement input = driver.FindElement(By.Id("ctl00_cphBody_ibMessagesSearch"));
                    input.Click();
                    //ожидаем случайный промежуток времени, чтобы данные успели прогрузиться
                    Thread.Sleep(rand.Next(500, 2000));

                    List<DebtorMessage> messages = new();
                    //теперь нам нужено пробежать по всем страницам результатов поиска
                    int curPage = 1;
                    do
                    {
                        ProgressChanged?.Invoke(new ExportStage()
                        {
                            Name = $"Обход страниц с сообщениями (Прочитано страниц с сообщениями: {curPage - 1})",
                            Done = curPage - 1
                        });
                        cancellationToken.ThrowIfCancellationRequested();
                        //вызываем показ выбранной страницы (curPage)
                        js.ExecuteScript("theForm.__EVENTTARGET.value = 'ctl00$cphBody$gvMessages';");
                        js.ExecuteScript($"theForm.__EVENTARGUMENT.value = 'Page${curPage}';");
                        js.ExecuteScript("theForm.submit();");
                        //ожидаем завершения всех Ajax запросов
                        WaitForAjax(driver);
                        //ожидаем еще дополнительно немного времени,
                        //чтобы все отобразилось
                        Thread.Sleep(rand.Next(300, 1001));
                        //если замечаем, что на странице есть надпись о том, что результатов нет,
                        //то значит что мы дошли до конца
                        bool isSearchFailed = driver.PageSource.Contains(
                            "По заданным критериям не найдено ни одной записи. Уточните критерии поиска");
                        if (isSearchFailed)
                        {
                            //если дошли до конца, То просто ждем определенный
                            //промежуток времени и заканчиваем цикл обхода
                            Thread.Sleep(rand.Next(500, 5000));
                            //Если поиск не дал результатов
                            break;
                        }

                        //Если же результаты есть, то ищем таблицу, в которой будут результаты поиска
                        IWebElement table = driver.FindElement(By.Id("ctl00_cphBody_gvMessages"));
                        //и получаем массив ее строчек
                        ReadOnlyCollection<IWebElement> rows = table.FindElements(By.TagName("tr"));
                        //теперь наша задача пробежать по всем строкам и заполнить данные со страницы
                        foreach (IWebElement row in rows)
                        {
                            //для каждой строки для сначала
                            //получаем коллекцию всех ее столбцов
                            ReadOnlyCollection<IWebElement> columns = row.FindElements(By.TagName("td"));
                            //Если число столбцов не равно 5, то эта строка не наша,
                            //обычно такое происходит в конце таблицы, так как там
                            //данные, не входящие в таблицу, тоже в нее решили включить.
                            if (columns.Count != 5) continue;

                            //временный костыль потому что пока не знаем
                            //что делать с аннулированными сообщениями
                            if (columns[1].Text.ToLower().Contains("аннулировано"))
                                continue;

                            //сначала обрабатываем дату публикации сообщения
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

                            //затем берем тип акта суда
                            buffer.Type = new DebtorMessageType() { Name = columns[1].Text };

                            //также отсюда вытаскиваем GUID сообщения
                            string messagePageLink = columns[1].FindElement(By.TagName("a")).GetAttribute("href");
                            buffer.Guid = messagePageHrefRegex.Match(messagePageLink).Groups[1].Value;

                            //далее считываем ФИО должника.
                            buffer.Debtor = new Debtor() { FullName = columns[2].Text };

                            //затем заполняем адрес (должника или суда?)
                            buffer.Address = columns[3].Text;
                            //в конце будет идти ФИО арбитражного управляющего.
                            buffer.Owner = new ArbitrManager() { FullName = columns[4].Text };
                            messages.Add(buffer);
                        }

                        curPage++;

                    } while (true);

                    //переводим получившийся список в массив
                    resultList.AddRange(messages.ToArray());

                }
                finally
                {
                    CloseBrowser();
                    isLoading = false;
                }
            }


            //возвращаем результат
            return resultList.ToArray();
        }

        private static void CloseBrowser()
        {
            driver.Close();
            driver.Dispose();
        }

        /// <summary>
        /// Забирает управление на время ожидания
        /// выполнения всех Ajax запросов указанным драйвером браузера.
        /// </summary>
        /// <param name="driver">Драйвер, Ajax запросы которого необходимо подождать.</param>
        private static void WaitForAjax(IWebDriver driver)
        {
            IJavaScriptExecutor js = (driver as IJavaScriptExecutor);
            while (true)
            {
                bool ajaxIsComplete = (bool)js.ExecuteScript("return jQuery.active == 0");
                if (ajaxIsComplete)
                    break;
                Thread.Sleep(100);
            }
        }

        /// <summary>
        /// Массив (только для чтения) поддерживаемых типов вытаскиваемых сообщений.
        /// </summary>
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
        private static Random rand = new();
        private const string baseMessageViewUrl = @"https://bankrot.fedresurs.ru/MessageWindow.aspx?ID=";
        /// <summary>
        /// Производит заполнение даты рождения по коллекции должников при помощи указанного драйвера браузера.
        /// </summary>
        /// <param name="driver">Драйвер браузера, который будет использоваться при заполнении коллекции датами рождения.</param>
        /// <param name="messages">Коллекция сообщений по должникам, в которой необходимо заполнить дату рождения.</param>
        private static void FillBirthDate(
            IEnumerable<DebtorMessage> messages,
            IWebDriver driver = null)
        {
            if (driver == null)
                driver = GetDriver();
            int length = messages.Count();
            int doneCount = 0;
            foreach (DebtorMessage message in messages)
            {
                ProgressChanged?.Invoke(new ExportStage()
                {
                    Name = $"Актуализация дат рождения ({doneCount} готово из {length})",
                    Done = doneCount,
                    AllCount = length
                });
                cancellationToken.ThrowIfCancellationRequested();
                string url = $"{baseMessageViewUrl}{message.Guid}";
                driver.Url = url;
                WebDriverWait wait = new(driver, new TimeSpan(0, 0, 10));
                driver.Navigate();
                wait.Until(driver => driver.FindElement(By.ClassName("even")));
                Thread.Sleep(rand.Next(ClientSettings.Settings.MinRequestDelayInMsec, ClientSettings.Settings.MaxRequestDelayInMsec));

                ReadOnlyCollection<IWebElement> rows = driver.FindElements(By.TagName("tr"));

                foreach (IWebElement row in rows)
                {
                    if (!string.IsNullOrWhiteSpace(row.GetAttribute("class")))
                    {
                        ReadOnlyCollection<IWebElement> columns = row.FindElements(By.TagName("td"));
                        if (columns.First().Text != "Дата рождения") continue;
                        Match dateTimeMatch = dateRegex.Match(columns.Last().Text);
                        DateTime date = new(
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

                doneCount++;
            }
            Console.WriteLine();

        }
        /// <summary>
        /// Производит экспорт коллекции сообщений по должникам в поток памяти, содержащий Excel файл.
        /// </summary>
        /// <param name="messages">Коллекция сообщений по должникам, которую необходимо сохранить в виде Excel файла.</param>
        /// <returns>Поток памяти, содержащий в себе Excel файл.</returns>
        public static MemoryStream ExportMessagesToExcel(IEnumerable<DebtorMessage> messages)
        {
            ProgressChanged?.Invoke(new ExportStage()
            {
                Name = "Выгрзука в Excel файл"
            });
            ExcelPackage excelPackage = new();
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("Выгрузка сообщений");

            for (int i = 0; i < columns.Length; i++)
            {
                sheet.Column(i + 1).Style.Numberformat.Format = columns[i].ColumnFormat;
                sheet.Cells[1, i + 1].Value = columns[i].Header;
                sheet.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thick);
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
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

            ExcelRange range = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
            ExcelTable tab = sheet.Tables.Add(range, "Table1");
            tab.TableStyle = TableStyles.Light16;

            MemoryStream stream = new();
            excelPackage.SaveAs(stream);
            return stream;
        }
        /// <summary>
        /// Список столбцов (только для чтения),
        /// которые будет обрабатывать алгоритм обработки сообщений.
        /// </summary>
        private static readonly DebtorMessageExcelExportColumn[] columns = new DebtorMessageExcelExportColumn[]
        {
            new(
                "Дата публикации (МСК)",
                message => message.DatePublished,
                "dd.mm.yyyy"),

            new(
                "Тип сообщения",
                message => message.Type.Name),

            new(
                "Должник",
                message => message.Debtor.FullName),

            new(
                "Дата рождения",
                message => message.Debtor.BirthDate,
                "dd.mm.yyyy"),

            new(
                "Адрес",
                message => message.Address),

            new(
                "Кем опубликовано",
                message => message.Owner.FullName),
        };
        /// <summary>
        /// Определяет объект столбца, который будет обработан алгоритмом генерации Excel файла.
        /// </summary>
        private class DebtorMessageExcelExportColumn
        {
            /// <summary>
            /// Создает новый столбец 
            /// </summary>
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
            //Установка некоммерческой лицензии использования EPPLUS,
            //так как мы не собираемся продавать программу.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private static Regex dateRegex = new(@"(?<day>\d+)\.(?<month>\d+)\.(?<year>\d{4})");
        private static readonly Regex messagePageHrefRegex = new(@"/MessageWindow\.aspx\?ID=([A-Z0-9]+)");
        private static readonly Regex datetimeRegex = new(@"(?<day>\d+)\.(?<month>\d+)\.(?<year>\d{4}) (?<hours>\d+):(?<minutes>\d+):(?<seconds>\d+)");

    }

}
