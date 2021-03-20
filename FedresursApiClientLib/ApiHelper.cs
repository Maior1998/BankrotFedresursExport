using System;
using System.Net;
using System.Text;
using FedresursAPI;
using System.ServiceModel;

namespace FedresursApiClientLib
{
    public static class ApiHelper
    {
        public static void Test()
        {
            StringBuilder result = new();
            MessageServiceClient client = new(
                new BasicHttpBinding()
                {
                    Security = new BasicHttpSecurity()
                    {
                        Transport = new HttpTransportSecurity()
                        {
                            ClientCredentialType = HttpClientCredentialType.Digest,
                            ProxyCredentialType = HttpProxyCredentialType.Digest
                        },
                        Mode = BasicHttpSecurityMode.Transport,

                    },

                },
                new EndpointAddress(@"https://services.fedresurs.ru/Bankruptcy/MessageService/WebService.svc"));
            client.ClientCredentials.HttpDigest.ClientCredential.UserName = "demowebuser";
            client.ClientCredentials.HttpDigest.ClientCredential.Password = "Ax!761BN";
            //client.ClientCredentials.HttpDigest.AllowedImpersonationLevel = TokenImpersonationLevel.Impersonation;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;


            


            // Получение идентификаторов сообщении за указанный период
            DateTime dateFrom = new(2021, 03, 17);
            DateTime dateTo = new(2021, 03, 21);
            var testResponce = client.GetDebtorsByLastPublicationPeriod(dateFrom, dateTo);
            Console.WriteLine();
            #region Output
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(
                $"Получение идентификаторов сообщении за период с {dateFrom:dd-MM-yyyy} по {dateTo:dd-MM-yyyy}:{Environment.NewLine}");
            Console.ResetColor();
            #endregion
            int[] messageIds = client.GetMessageIds(dateFrom, dateTo);
            #region Output
            if (messageIds != null && messageIds.Length > 0)
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                foreach (int messageId in messageIds)
                {
                    Console.WriteLine(messageId);
                }
                Console.ResetColor();

                Console.WriteLine();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("Сообщений не найдено");
                Console.ResetColor();
            }
            #endregion

            if (messageIds != null && messageIds.Length > 0)
            {

                for (int i = 0; i < messageIds.Length; i++)
                {
                    int curMessageId = messageIds[i];

                    // Получение контента сообщения по его идентификатору
                    #region Output
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(
                        $"Получение контента сообщения с идентификатором {curMessageId}:{Environment.NewLine}");
                    Console.ResetColor();
                    #endregion

                    Console.WriteLine();

                    string messageContent = client.GetMessageContent(curMessageId);
                    #region Output
                    if (messageContent != null)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        result.AppendLine(messageContent);
                        Console.WriteLine(messageContent);
                        Console.ResetColor();

                        Console.WriteLine();
                    }
                    #endregion
                }
            }
            Console.WriteLine();

        }
    }
}
