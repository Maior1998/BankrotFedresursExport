using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using DevExpress.Mvvm;

using Microsoft.Win32;

using ReactiveUI;
using ReactiveUI.Fody.Helpers;
using BankruptFedresursClient;
using BankruptFedresursModel;

namespace BankrotFedresursExport
{
	public class MainViewModel : ReactiveObject
	{
		public MainViewModel()
		{	//значение строки задается в переменную, но при запуске не отображается
			SelectedMessageType = MessageTypes.First().ToString();
		}

		[Reactive] public DateTime DateFrom { get; set; } = DateTime.Today;
		[Reactive] public DateTime DateTo { get; set; } = DateTime.Today;

		public DebtorMessageType[] MessageTypes { get; set; } = BankrotClient.SupportedMessageTypes;

		[Reactive] public string SelectedMessageType { get; set; }

		private DelegateCommand save;

		public DelegateCommand Save => save ??= new DelegateCommand(
			() =>
			{
				SaveFileDialog saveFileDialog = new();
				IsLoading = true;
				if (!saveFileDialog.ShowDialog().Value)
				{
					IsLoading = false;
					return;
				}
				string filePath = saveFileDialog.FileName;
				IsLoading = false;
				Console.WriteLine();
			}, () =>
			{
				return SelectedMessageType != null;
			});

		[Reactive] public bool IsLoading { get; set; }
	}
}