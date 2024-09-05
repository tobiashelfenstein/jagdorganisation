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
        private readonly CheckBox[] _checkboxes;

        private readonly BackgroundWorker _worker;
        private HunterGroupPrinter _printer;
        public MainWindow()
        {
            InitializeComponent();

            // initialize CheckBox array
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
        }

        struct DivisionData
        {
            public string Filename;
            public List<string> Checkboxes;
            public bool? Separator;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            (sender as BackgroundWorker).ReportProgress(10, "Einteilungsdatei wird verarbeitet");
            _printer = new HunterGroupPrinter();
            _printer.CreateCardsFromSource(((DivisionData)e.Argument).Filename);

            int progress = (100 - 20) / ((DivisionData)e.Argument).Checkboxes.Count;
            for (int i = 0; i < ((DivisionData)e.Argument).Checkboxes.Count; i++)
            {
                if (!(sender as BackgroundWorker).CancellationPending)
                {
                    (sender as BackgroundWorker).ReportProgress(10 + ((i + 1) * progress), ((DivisionData)e.Argument).Checkboxes[i] + " werden gedruckt");
                    PrintGroups(((DivisionData)e.Argument).Checkboxes[i], ((DivisionData)e.Argument).Separator);
                }
            }

            (sender as BackgroundWorker).ReportProgress(100, "Fertig");
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
                // TODO: RaceCondition -> beste Möglichkeit?
                MessageBox.Show(this, "Druck unterbrochen!", "Jagdorganisation", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                MessageBox.Show(this, "Alle Gruppeneinteilungen wurden gedruckt!", "Jagdorganisation", MessageBoxButton.OK, MessageBoxImage.Information);
                ProgressBar.Value = 0;
                StatusInfoText.Content = "keine Einteilung geladen";

                // unlock user interface, only abort is disabled
                LockUserInterface(false);
            }
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            // see https://automationtesting.in/row-count-excel-using-c/

            // check, if at least one CheckBox is checked
            if (Array.TrueForAll(_checkboxes, IsCheckBoxSelected))
            {
                MessageBox.Show(this, "Keine Gruppe ausgewählt!", "Jagdorganisation", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DivisionData division;

            // select only excel files
            Microsoft.Win32.OpenFileDialog open_dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Jagdeinteilung laden",
                Filter = "Jagdeinteilung (.xlsx)|*.xlsx"
            };

            // show open file dialog box
            // if user has canceld file selection, exit print action
            if (open_dialog.ShowDialog() != true)
            {
                return;
            }

            // lock user interface, only abort is enabled
            LockUserInterface(true);

            // define selected file as source file
            division.Filename = open_dialog.FileName;

            // create array with checkboxes for printing groups
            division.Checkboxes = new List<string>();
            division.Separator = SeparatorCheckBox.IsChecked;
            foreach (CheckBox box in _checkboxes)
            {
                if (box.IsChecked == true)
                {
                    division.Checkboxes.Add(box.Content.ToString());
                }
            }

            // process all data
            _worker.RunWorkerAsync(division);
        }

        private void PrintGroups(string group, bool? separator)
        {
            // print out if the box is checked
            _printer.PrintCards(group, separator);

            // timer before next print action
            Thread.Sleep(30 * 1000);
        }

        private void LockUserInterface(bool locking)
        {
            // when locking is true, disable all gui functions
            PrintButton.IsEnabled = !locking;
            SettingsButton.IsEnabled = !locking;
            CloseButton.IsEnabled = !locking;

            // but enable the cancel button
            AbortButton.IsEnabled = locking;
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
    }
}
