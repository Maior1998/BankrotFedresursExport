using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
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
		{	//значение строки задается в переменную, но при запуске не отображается
			SelectedMessageType = MessageTypes.First();
		}

		[Reactive] public DateTime DateFrom { get; set; } = DateTime.Today;
		[Reactive] public DateTime DateTo { get; set; } = DateTime.Today;

		public DebtorMessageType[] MessageTypes { get; set; } = BankrotClient.SupportedMessageTypes;

		[Reactive] public DebtorMessageType SelectedMessageType { get; set; }

		private DelegateCommand save;

		public DelegateCommand Save => save ??= new DelegateCommand(
			() =>
			{
				SaveFileDialog saveFileDialog = new()
				{
					AddExtension = true,
					DefaultExt = "xlsx",
					Filter = "Таблица Microsoft Excel|*.xlsx|Все файлы|*.*",
					FileName = $"({DateTime.Today.ToShortDateString()}) - {SelectedMessageType}"
				};

				IsLoading = true;
				if (!saveFileDialog.ShowDialog().Value)
				{
					IsLoading = false;
					return;
				}
				string filePath = saveFileDialog.FileName;
				
				Task.Run(() =>
				{
					try
					{

						MemoryStream memoryStream =
							BankrotClient.ExportMessagesToExcel(BankrotClient.GetMessages(DateFrom, DateTo, SelectedMessageType));
						File.WriteAllBytes(filePath, memoryStream.ToArray());
						MessageBox.Show("Файл записан.", "Успех");
					}
					catch (Exception e)
					{
						MessageBox.Show(e.Message);
						throw;
					}
					finally{ IsLoading = false; }
				});

			}, () =>
			{
				return SelectedMessageType != null;
			});

		[Reactive] public bool IsLoading { get; set; }
	}
}