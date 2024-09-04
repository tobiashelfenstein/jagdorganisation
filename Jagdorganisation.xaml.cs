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

using System.Collections;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;

using System.Threading;
using System.ComponentModel;

namespace Jagdorganisation
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private HunterGroupPrinter _printer;
        public MainWindow()
        {
            InitializeComponent();
        }

        struct ProcessData
        {
            public string Filename;
            public List<string> Checkboxes;
            public bool? Separator;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            (sender as BackgroundWorker).ReportProgress(10, "Einteilungsdatei wird verarbeitet");
            _printer = new HunterGroupPrinter();
            _printer.CreateCardsFromSource(((ProcessData)e.Argument).Filename);

            int progress = (100-20)/((ProcessData)e.Argument).Checkboxes.Count;
            for (int i = 0; i < ((ProcessData)e.Argument).Checkboxes.Count; i++)
            {
                (sender as BackgroundWorker).ReportProgress(10 + (i + 1) * progress, ((ProcessData)e.Argument).Checkboxes[i] + " werden gedruckt");
                PrintGroups(((ProcessData)e.Argument).Checkboxes[i], ((ProcessData)e.Argument).Separator);
            }

            (sender as BackgroundWorker).ReportProgress(100, "Fertig");
        }

        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
            StatusInfoText.Content = e.UserState.ToString();
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Alle Gruppeneinteilungen wurden gedruckt!", "Jagdorganisation", MessageBoxButton.OK, MessageBoxImage.Information);
            ProgressBar.Value = 0;
            StatusInfoText.Content = "keine Einteilung geladen";
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            // see https://automationtesting.in/row-count-excel-using-c/

            // select only excel files
            var open_dialog = new Microsoft.Win32.OpenFileDialog();
            open_dialog.Title = "Jagdeinteilung laden";
            open_dialog.Filter = "Jagdeinteilung (.xlsx)|*.xlsx";

            // show open file dialog box
            bool? result = open_dialog.ShowDialog();

            // save file name
            string source_file = "";
            if (result == true)
            {
                source_file = open_dialog.FileName;
            }
            else
            {
                return;
            }

            // create array with checkboxes for printing groups
            ProcessData process_data;
            process_data.Filename = source_file;
            process_data.Checkboxes = new List<string>();
            process_data.Separator = SeparatorCheckBox.IsChecked;
            foreach(CheckBox box in new CheckBox[] {
                LeaderCheckBox,
                ShootersCheckBox,
                DogsCheckBox,
                ReservesCheckBox
            })
            {
                if (box.IsChecked == true)
                {
                    process_data.Checkboxes.Add(box.Content.ToString());
                }
                
            }

            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync(process_data);
        }

        private void PrintGroups(string group, bool? separator)
        {
            // print out if the box is checked
            _printer.PrintCards(group, separator);

            // timer before next print action
            Thread.Sleep(30 * 1000);
        }
    }
}
