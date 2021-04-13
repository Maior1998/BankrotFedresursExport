using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;

using BankruptFedresursModel;

using HtmlAgilityPack;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace BankruptFedresursClient
{
    public static class BankrotClient
    {
        private const string NotFoundString =
            "По заданным критериям не найдено ни одной записи. Уточните критерии поиска";
        private static CancellationToken cancellationToken = new();
        private static Cookie sessionCookie;
        private static bool isLoading;
        private const string UserAgent =
            "Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25";
        public static event Action<ExportStage> ProgressChanged;
        public static void SetCancellationToken(CancellationToken token)
        {
            cancellationToken = token;
        }

        public static DebtorMessage[] GetMessagesWithBirthDates(DateTime date, DebtorMessageType[] type)
        {
            return GetMessagesWithBirthDates(date, date, type);
        }

        public static DebtorMessage[] GetMessagesWithBirthDates(DateTime startDate, DateTime endDate, DebtorMessageType[] type)
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
        /// <param name="types">МассовОбъект типа сообщения суда, которое необходимо вытащить.</param>
        /// <returns>Массив сообщений по должникам с учетом фильтров по дате публикации и типу сообщения.</returns>
        public static DebtorMessage[] GetMessages(DateTime date, DebtorMessageType[] types)
        {
            return GetMessages(date, types);
        }

        private static Cookie GetCookie()
        {
            const string url = @"https://bankrot.fedresurs.ru";

            HttpWebRequest request = WebRequest.CreateHttp(url);
            request.UserAgent = UserAgent;
            request.Method = "POST";
            request.CookieContainer = new CookieContainer(10);
            WebResponse resp = request.GetResponse();
            new StreamReader(resp.GetResponseStream()).ReadToEnd();
            Cookie sessionCookie = request.CookieContainer.GetCookies(new Uri(url))[0];
            return sessionCookie;
        }


        /// <summary>
        /// Получает массив сообщений по должникам при помощи фильтров:
        /// интервал публикации сообщения (дата-начало, дата-конец),
        /// а также фильтр по типу опубликуемого арбитражным управляющим сообщения.
        /// </summary>
        /// <param name="start">Дата-начало фильтра по дате публикации сообщения.</param>
        /// <param name="end">Дата-конец фильтра по дате публикации сообщения.</param>
        /// <param name="messageTypes">Объект типа сообщения суда, которое необходимо вытащить.</param>
        /// <returns>Массив сообщений по должникам с учетом фильтров по дате публикации и типу сообщения.</returns>
        public static DebtorMessage[] GetMessages(DateTime start, DateTime end, DebtorMessageType[] messageTypes)
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

            StartDate = start;
            EndDate = end;
            List<DebtorMessage> resultList = new();
            sessionCookie = GetCookie();
            foreach (DebtorMessageType curMessageType in messageTypes)
            {
                ProgressChanged?.Invoke(new ExportStage()
                {
                    Name = $"Обрабатывается сообщение \"{curMessageType.Name}\" (номер: {curMessageType.Id})"
                });
                MessageId = curMessageType.Id;
                List<DebtorMessage> messages = new();
                //теперь нам нужно пробежать по всем страницам результатов поиска
                ushort curPage = 0;
                do
                {
                    PageId = curPage;
                    cancellationToken.ThrowIfCancellationRequested();
                    //вызываем показ выбранной страницы (curPage)
                    string pageContent = GetMessagesPage();
                    Thread.Sleep(rand.Next(300, 1001));
                    //если замечаем, что на странице есть надпись о том, что результатов нет,
                    //то значит что мы дошли до конца
                    bool isSearchFailed = pageContent.Contains(
                        NotFoundString);
                    if (isSearchFailed)
                    {
                        //если дошли до конца, То просто ждем случайно
                        //определенный промежуток времени и заканчиваем
                        //цикл обхода
                        Thread.Sleep(rand.Next(500, 5000));
                        //Если поиск не дал результатов
                        break;
                    }

                    HtmlDocument doc = new();
                    doc.LoadHtml(pageContent);
                    HtmlNode messagesTable = doc.DocumentNode.SelectSingleNode("//*[@id=\"ctl00_cphBody_gvMessages\"]");
                    HtmlNodeCollection rows = messagesTable.SelectNodes("tr");
                    foreach (HtmlNode row in rows)
                    {
                        HtmlNodeCollection columns = row.SelectNodes("td");
                        if (columns == null || columns.Count != 5) continue;
                        string datePublished = columns[0].InnerText.Trim('\r', '\n', '\t');
                        string messageType = columns[1].InnerText.Trim('\r', '\n', '\t');
                        string messageLink = columns[1].SelectSingleNode("a").GetAttributeValue("href", string.Empty);
                        string messageGuid = messagePageHrefRegex.Match(messageLink).Groups[1].Value;
                        string debtorFullName = columns[2].InnerText.Trim('\r', '\n', '\t');
                        string address = columns[3].InnerText.Trim('\r', '\n', '\t');
                        string authorFullName = columns[4].InnerText.Trim('\r', '\n', '\t');

                        if (messageType != "Сообщение о судебном акте")
                        {
                            continue;
                        }

                        DebtorMessage buffer = new()
                        {
                            Address = address,
                            DatePublished = DateTime.ParseExact(datePublished, "dd.MM.yyyy HH:mm:ss", CultureInfo.CurrentCulture),
                            Debtor = new Debtor() { FullName = debtorFullName },
                            Owner = new ArbitrManager() { FullName = authorFullName },
                            Type = curMessageType,
                            Guid = messageGuid
                        };
                        messages.Add(buffer);

                    }
                    ProgressChanged?.Invoke(new ExportStage()
                    {
                        Name = $"Обход страниц с сообщениями (Прочитано страниц с сообщениями: {curPage + 1})",
                        Done = curPage
                    });
                    curPage++;

                } while (true);

                //переводим получившийся список в массив
                resultList.AddRange(messages.ToArray());

            }

            isLoading = false;
            //возвращаем результат
            return resultList.ToArray();
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
            IEnumerable<DebtorMessage> messages)
        {
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
                Thread.Sleep(ClientSettings.Settings.RequestDelay);
                message.Debtor.BirthDate = GetDebtorBirthDate(message.Guid);
                doneCount++;
            }
            Console.WriteLine();

        }


        public static DateTime StartDate;
        public static DateTime EndDate;
        public static ushort MessageId;
        public static ushort PageId;
        private const string cookieStringTemplate =
            "MessageNumber=" +
            "&MessageType=ArbitralDecree" +
            "&MessageTypeText=%d0%a1%d0%be%d0%be%d0%b1%d1%89%d0%b5%d0%bd%d0%b8%d0%b5+%d0%be+%d1%81%d1%83%d0%b4%d0%b5%d0%b1%d0%bd%d0%be%d0%bc+%d0%b0%d0%ba%d1%82%d0%b5" +
            "&DateEndValue={0}+0%3a00%3a00" +
            "&DateBeginValue={1}+0%3a00%3a00" +
            "&PageNumber={3}" +
            "&DebtorText=" +
            "&DebtorId=" +
            "&DebtorType=" +
            "&PublisherType=" +
            "&PublisherId=" +
            "&PublisherText=" +
            "&IdRegion=" +
            "&IdCourtDecisionType={2}" +
            "&WithAu=False" +
            "&WithViolation=False";

        public static string CookieString => string.Format(
            cookieStringTemplate,
            StartDate.ToString("dd.MM.yyyy"),
            EndDate.ToString("dd.MM.yyyy"),
            MessageId,
            PageId);
        private const string MessagesUrl = "https://bankrot.fedresurs.ru/Messages.aspx?attempt=1";
        public static string GetMessagesPage()
        {
            HttpWebRequest request = WebRequest.CreateHttp(MessagesUrl);
            request.UserAgent = UserAgent;
            request.Method = "POST";
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(sessionCookie);
            request.CookieContainer.Add(new Cookie("Messages", CookieString, sessionCookie.Path, sessionCookie.Domain));
            return new StreamReader(request.GetResponse().GetResponseStream()).ReadToEnd();
        }

        public static DateTime GetDebtorBirthDate(string messageGuid)
        {
            HttpWebRequest request = WebRequest.CreateHttp($"{baseMessageViewUrl}{messageGuid}");
            request.UserAgent = UserAgent;
            request.Method = "POST";
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(sessionCookie);
            string response = new StreamReader(request.GetResponse().GetResponseStream()).ReadToEnd();
            HtmlDocument doc = new();
            doc.LoadHtml(response);
            HtmlNode birthdateFirstColumn = doc.DocumentNode.SelectSingleNode("//td[contains(text(),\"Дата рождения\")]");
            HtmlNode row = birthdateFirstColumn.ParentNode;
            HtmlNode birthdateLastColumn = row.LastChild;
            return DateTime.Parse(birthdateLastColumn.InnerText);
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
                Name = "Выгрузка в Excel файл"
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

            new ("Тип решения суда",
                message=>message.Type.Name),

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
            private readonly Expression<Func<DebtorMessage, object>> cellValueFromMessageExpression;

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
