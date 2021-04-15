using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BankruptFedresursClient;
using Newtonsoft.Json;

namespace BankruptFedresursScript
{
    /// <summary>
    /// Представляет собой пакет настроек клиента Selenium при работе с сайтом федресурса.
    /// </summary>

    public class ScriptSettings
    {
        public ushort DateStartOffset { get; set; }
        public ushort DateEndOffset { get; set; }
        public string ExcelExportFilePath { get; set; }

        private const string SettingsFilePath = "ScriptSettings.json";
        public static readonly ScriptSettings Settings;

        static ScriptSettings()
        {
            Settings = new ScriptSettings()
            {
                DateStartOffset = 1,
                DateEndOffset = 1,
                ExcelExportFilePath = ""
            };

            try
            {
                string settingsContent = File.ReadAllText(SettingsFilePath);
                ScriptSettings settings = JsonConvert.DeserializeObject<ScriptSettings>(settingsContent);
                Settings = settings;
                Console.WriteLine("Successfully readed script settings");
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
