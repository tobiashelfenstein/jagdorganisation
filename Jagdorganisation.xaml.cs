using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace Jagdorganisation
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private struct DivisionData
        {
            public string Filename;
            public bool Separator;
            public List<string> Checkboxes;
        }

        private readonly CheckBox[] _checkboxes;
        private readonly BackgroundWorker _worker;
        private PrintManager _printer;
        private HunterGroupPrinter _creator;

        public MainWindow()
        {
            InitializeComponent();

            // initialize CheckBox array
            // separator CheckBox is intentionally not included
            _checkboxes = new CheckBox[] {
                LeaderCheckBox,
                ShootersCheckBox,
                DogsCheckBox,
                ReservesCheckBox
            };

            // initialize BackgroundWorker
            _worker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            // assign methods to background worker
            _worker.DoWork += Worker_DoWork;
            _worker.ProgressChanged += Worker_ProgressChanged;
            _worker.RunWorkerCompleted += Worker_RunWorkerCompleted;

            // start print manager
            _printer = new PrintManager();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            worker.ReportProgress(10, "Einteilungsdatei wird verarbeitet");

            _creator = new HunterGroupPrinter();
            _creator.CreateCardsFromSource(((DivisionData)e.Argument).Filename);

            int progress = (100 - 20) / ((DivisionData)e.Argument).Checkboxes.Count;
            for (int i = 0; i < ((DivisionData)e.Argument).Checkboxes.Count; i++)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                worker.ReportProgress(10 + ((i + 1) * progress), ((DivisionData)e.Argument).Checkboxes[i] + " werden gedruckt");
                _creator.PrintCards(((DivisionData)e.Argument).Checkboxes[i], ((DivisionData)e.Argument).Separator);

                Thread.Sleep(30 * 1000);
            }

            worker.ReportProgress(100, "Fertig");
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
            StatusInfoText.Content = e.UserState.ToString();
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show
                (
                    this,
                    "Druck wurde abgebrochen! " +
                    "Erfolgreich erstelle Karten befinden sich im Drucker.",
                    "Jagdorganisation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
            else
            {
                MessageBox.Show
                (
                    this,
                    "Alle Gruppeneinteilungen wurden gedruckt!",
                    "Jagdorganisation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information
                );
            }

            _printer.ResetDefaultPrinter();
            ResetInterface();
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            // see https://automationtesting.in/row-count-excel-using-c/

            // check, if at least one CheckBox is checked
            if (Array.TrueForAll(_checkboxes, IsCheckBoxSelected))
            {
                MessageBox.Show
                (
                    this,
                    "Keine Gruppe ausgewählt!",
                    "Jagdorganisation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );

                return;
            }

            PrinterSelection printer_dialog = new PrinterSelection();
            if (printer_dialog.ShowDialog() != true)
            {
                return;
            }

            _printer.SetSessionPrinter(printer_dialog.SelectedPrinter);
            _printer.SetPrinterSettings(PrinterHelper.ColorMode.DMCOLOR_MONOCHROME, PrinterHelper.PageDuplex.DMDUP_SIMPLEX);

            var devMode = PrinterHelper.GetPrinterDevMode(null);
            string s = String.Format("{0} ist Color: {1}", devMode.dmDeviceName, devMode.dmColor);
            Console.WriteLine(s);



            Microsoft.Win32.OpenFileDialog open_dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Jagdeinteilung laden",
                Filter = "Jagdeinteilung (.xlsx)|*.xlsx"
            };

            if (open_dialog.ShowDialog() != true)
            {
                return;
            }

            // lock user interface, only abort is enabled
            LockUserInterface(true);

            DivisionData div_data = new DivisionData
            {
                Filename = open_dialog.FileName,
                Separator = SeparatorCheckBox.IsChecked ?? false,
                Checkboxes = new List<string>()
            };

            // create list with checkbox discription for printing groups
            foreach (CheckBox box in _checkboxes)
            {
                if (box.IsChecked == true)
                {
                    div_data.Checkboxes.Add(box.Content.ToString());
                }
            }

            _worker.RunWorkerAsync(div_data);
        }

        private void LockUserInterface(bool locking)
        {
            // when locking is true, disable all gui functions
            PrintButton.IsEnabled = !locking;
            SettingsButton.IsEnabled = !locking;
            CloseButton.IsEnabled = !locking;

            // but enable the cancel button
            AbortButton.IsEnabled = locking;

            // also disbale all checkboxes, when locking is true
            foreach (CheckBox box in _checkboxes)
            {
                box.IsEnabled = !locking;
            }

            // disbale separator CheckBox manually
            // because its not included in _checkbox
            SeparatorCheckBox.IsEnabled = !locking;
        }

        private void AbortButton_Click(object sender, RoutedEventArgs e)
        {
            if (_worker.IsBusy)
            {
                _worker.CancelAsync();
                _worker.Dispose();
            }
        }

        private bool IsCheckBoxSelected(CheckBox box)
        {
            // reverse operation: if box is checked, return false
            // only unchecked boxes return true
            // also null returns true
            return !box.IsChecked ?? true;
        }

        private void ResetInterface()
        {
            ProgressBar.Value = 0;
            StatusInfoText.Content = "keine Einteilung geladen";

            // unlock user interface, only abort is disabled
            LockUserInterface(false);
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (_worker.IsBusy)
            {
                e.Cancel = true;
                MessageBox.Show
                (
                    this,
                    "Die Anwendung kann nicht geschlossen werden, " +
                    "da der Druckprozess gestartet wurde!",
                    "Jagdorganisation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SettingsWindow settings_dlg = new SettingsWindow();
            settings_dlg.Show();
        }
    }
}
