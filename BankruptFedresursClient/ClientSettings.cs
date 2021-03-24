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
        
        public string DriverPath { get; set; }
        public ushort MinRequestDelayInMsec { get;  set; }
        public ushort MaxRequestDelayInMsec { get;  set; }
        private const string SettingsFilePath = "ClientSettings.json";
        public static readonly ClientSettings Settings;

        static ClientSettings()
        {
            Settings = new ClientSettings()
            {
                DriverPath = "./",
                MinRequestDelayInMsec = 2500,
                MaxRequestDelayInMsec = 5000
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
