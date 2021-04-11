using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BankrotFedresursExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ((MainViewModel)DataContext).OnFromDateChanged += MainWindow_OnFromDateChanged;
        }

        private void MainWindow_OnFromDateChanged(DateTime obj)
        {
            dpTo.BlackoutDates.Clear();
            dpTo.SelectedDate = null;
            dpTo.BlackoutDates.Add(new CalendarDateRange(DateTime.MinValue,obj.Date.AddDays(-1).Date));
            DateTime calulatedDate = obj.AddDays(30).Date;
            DateTime todayTime = DateTime.Today.AddDays(1).Date;
            dpTo.BlackoutDates.Add(new CalendarDateRange(calulatedDate < todayTime ? calulatedDate : todayTime, DateTime.MaxValue));
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            dpFrom.BlackoutDates.Add(new CalendarDateRange(DateTime.MinValue, DateTime.Today.AddYears(-2)));
            dpFrom.BlackoutDates.Add(new CalendarDateRange(DateTime.Today.AddDays(1).Date, DateTime.MaxValue));
        }
    }
}
