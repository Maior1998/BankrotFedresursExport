using System;
using System.IO;
using Newtonsoft.Json;
using OpenQA.Selenium;

namespace BankruptFedresursClient
{
    /// <summary>
    /// Представляет собой пакет настроек клиента Selenium при работе с сайтом федресурса.
    /// </summary>
    
    public class ClientSettings
    {
        public string UserAgent { get; set; }
        public ushort MinRequestDelayInMsec { get;  set; }
        public ushort MaxRequestDelayInMsec { get;  set; }
        private const string SettingsFilePath = "ClientSettings.json";
        public static readonly ClientSettings Settings;

        static ClientSettings()
        {
            Settings = new ClientSettings()
            {
                MinRequestDelayInMsec = 2500,
                MaxRequestDelayInMsec = 2500,
                UserAgent = "Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25"
            };

            try
            {
                string settingsContent = File.ReadAllText(SettingsFilePath);
                ClientSettings settings = JsonConvert.DeserializeObject<ClientSettings>(settingsContent);
                Settings = settings;
                Console.WriteLine("Successfully readed settings");
            }
            catch (Exception readingException)
            {
                Console.WriteLine($"Error while reading settings file: {readingException.Message}");
                Console.WriteLine("Trying to save default settings file.");
                try
                {
                    File.WriteAllText(SettingsFilePath, JsonConvert.SerializeObject(Settings, Formatting.Indented));
                }
                catch (Exception savingException)
                {
                    Console.WriteLine($"Error while saving a default settings file! Message: {savingException.Message}");
                }
            }
        }

    }
}
