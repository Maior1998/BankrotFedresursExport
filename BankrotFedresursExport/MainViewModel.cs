using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Mvvm;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;
using BankruptFedresursClient;
using BankruptFedresursModel;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace BankrotFedresursExport
{
    public class MainViewModel : ReactiveObject, IDataErrorInfo
    {
        public MainViewModel()
        {
            this.WhenAnyValue(x => x.DateFrom).Subscribe(_ => UpdateDateTo());
            //значение строки задается в переменную, но при запуске не отображается
            BankrotClient.ProgressChanged += UpdateProgress;
            backgroundWorker.DoWork += (_, args) =>
            {
                ThreadPool.QueueUserWorkItem(_ => SaveMessages(args.Argument?.ToString()));

            };

        }

        public DateTime FromDateDisplayEnd => DateTime.Now;
        [Reactive] public DateTime ToDateDisplayEnd { get; private set; }
        private void UpdateDateTo()
        {
            DateTo = DateFrom;
            DateTime calculatedFromDateFrom = DateFrom.AddDays(29).Date;
            DateTime todayDate = DateTime.Today;

            ToDateDisplayEnd =
                todayDate < calculatedFromDateFrom
                    ? todayDate
                    : calculatedFromDateFrom;
        }
        private void UpdateProgress(ExportStage stage)
        {
            CurrentStatus = stage.Name;
            IsIndeterminate = stage.AllCount == 0;
            if (!IsIndeterminate)
                Progress = stage.Done * 100f / stage.AllCount;
        }
        [Reactive] public bool IsIndeterminate { get; set; }
        [Reactive] public float Progress { get; set; }
        [Reactive] public string CurrentStatus { get; set; }

        private static CancellationTokenSource cancellationTokenSource = new();
        [Reactive] public DateTime DateFrom { get; set; } = DateTime.Today;
        [Reactive] public DateTime DateTo { get; set; } = DateTime.Today;
        public MessageTypeSelectItem[] MessageTypes { get; }
            = BankrotClient.SupportedMessageTypes
                .Select(x => new MessageTypeSelectItem(x)).ToArray();

        private DelegateCommand save;

        private readonly BackgroundWorker backgroundWorker = new() { WorkerSupportsCancellation = true };
        public DelegateCommand Save => save ??= new DelegateCommand(
            () =>
            {
                if (!IsLoading)
                {
                    SaveFileDialog saveFileDialog = new()
                    {
                        AddExtension = true,
                        DefaultExt = "xlsx",
                        Filter = "Таблица Microsoft Excel|*.xlsx|Все файлы|*.*",
                        FileName = $"Выгрузка сообщений от {DateTime.Now:dd.MM.yyyy HH mm}"
                    };

                    IsLoading = true;
                    if (!saveFileDialog.ShowDialog().Value)
                    {
                        IsLoading = false;
                        return;
                    }

                    string filePath = saveFileDialog.FileName;
                    backgroundWorker.RunWorkerAsync(filePath);

                }
                else
                {

                    cancellationTokenSource.Cancel();
                    backgroundWorker.CancelAsync();
                    IsLoading = false;
                }
            }, () =>
            {
                return MessageTypes != null
                       && MessageTypes.Any(x => x.IsSelected)
                       && DateFrom <= DateTo;
            });



        public class MessageTypeSelectItem : ReactiveObject
        {
            public MessageTypeSelectItem(DebtorMessageType type)
            {
                Type = type;
            }
            public DebtorMessageType Type { get; set; }
            [Reactive] public bool IsSelected { get; set; }

            public override string ToString()
            {
                return Type?.Name;
            }
        }
        private void SaveMessages(string filePath)
        {
            cancellationTokenSource = new CancellationTokenSource();
            CancellationToken cancellationToken = cancellationTokenSource.Token;
            BankrotClient.SetCancellationToken(cancellationToken);
            try
            {
                MemoryStream memoryStream =
                    BankrotClient.ExportMessagesToExcel(
                        BankrotClient.GetMessagesWithBirthDates(
                            DateFrom,
                            DateTo,
                            MessageTypes
                                .Where(x => x.IsSelected)
                                .Select(x => x.Type).ToArray()));
                File.WriteAllBytes(filePath, memoryStream.ToArray());
                Process.Start("explorer.exe", $"/select, \"{filePath}\"");
            }
            catch (OperationCanceledException)
            {

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
            finally
            {
                IsLoading = false;
            }
        }

        [Reactive] public bool IsLoading { get; set; }
        public string Error => null;

        public string this[string columnName]
        {
            
            get
            {
                string result = null;
                
                switch (columnName)
                {
                    case nameof(DateFrom):
                    case nameof(DateTo):
                        if (DateFrom > DateTo)
                            result = "Дата начала поиска не может быть позднее даты конца поиска!";
                        if ((DateTo - DateFrom).TotalDays >= 30)
                            result = "Интервал поиска не может превышать 29 дней!";
                        break;
                        default:
                            break;
                }
                return result;
            }

        }
    }
}