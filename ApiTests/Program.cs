using System;
using System.Net;
using System.Security.Principal;
using System.ServiceModel;
using ServiceReference1 = buffer_net_framework.ServiceReference1;
namespace ApiTests
{
    class Program
    {
        static void Main(string[] args)
        {
            ServiceReference1.MessageServiceClient client = new(
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
                new EndpointAddress(@"https://services.fedresurs.ru/Bankruptcy/MessageServiceDemo/WebService.svc"));
            client.ClientCredentials.HttpDigest.ClientCredential.UserName = "demowebuser";
            client.ClientCredentials.HttpDigest.ClientCredential.Password = "Ax!761BN";
            //client.ClientCredentials.HttpDigest.AllowedImpersonationLevel = TokenImpersonationLevel.Impersonation;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            // Получение идентификаторов сообщении за указанный период
            DateTime dateFrom = new DateTime(2015, 11, 15);
            DateTime dateTo = new DateTime(2015, 12, 01);
            #region Output
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(String.Format("Получение идентификаторов сообщении за период с {0} по {1}:{2}", dateFrom.ToString("dd-MM-yyyy"), dateTo.ToString("dd-MM-yyyy"), Environment.NewLine));
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
                // Получение контента сообщения по его идентификатору
                #region Output
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(String.Format("Получение контента сообщения с идентификатором {0}:{1}", messageIds[0], Environment.NewLine));
                Console.ResetColor();
                #endregion
                string messageContent = client.GetMessageContent(messageIds[0]);
                #region Output
                if (messageContent != null)
                {
                    Console.ForegroundColor = ConsoleColor.DarkCyan;
                    Console.WriteLine(messageContent);
                    Console.ResetColor();

                    Console.WriteLine();
                }
                #endregion

            }

        }
    }
}
