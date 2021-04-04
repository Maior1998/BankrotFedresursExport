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
    public class MainViewModel : ReactiveObject
    {
        public MainViewModel()
        {   //значение строки задается в переменную, но при запуске не отображается
            BankrotClient.ProgressChanged += UpdateProgress;
            SelectedMessageType = MessageTypes.First();
            backgroundWorker.DoWork += (_, args) =>
            {
                ThreadPool.QueueUserWorkItem(_ => SaveMessages(args.Argument?.ToString()));

            };

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
        public DebtorMessageType[] MessageTypes { get; set; } = BankrotClient.SupportedMessageTypes;

        [Reactive] public DebtorMessageType SelectedMessageType { get; set; }

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
                        FileName = $"Выгрузка сообщений от {DateTime.Now:dd.MM.yyyy HH:mm}"
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
                return SelectedMessageType != null;
            });


        private void SaveMessages(string filePath)
        {
            cancellationTokenSource = new CancellationTokenSource();
            CancellationToken cancellationToken = cancellationTokenSource.Token;
            BankrotClient.SetCancellationToken(cancellationToken);
            try
            {
                MemoryStream memoryStream =
                    BankrotClient.ExportMessagesToExcel(
                        BankrotClient.GetMessages(DateFrom, DateTo, SelectedMessageType));
                File.WriteAllBytes(filePath, memoryStream.ToArray());
                Process.Start("explorer.exe",$"/select, \"{filePath}\"");
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
    }
}