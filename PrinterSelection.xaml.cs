using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing.Printing;
using System.Windows;

namespace Jagdorganisation
{
    /// <summary>
    /// Interaktionslogik für PrinterSelection.xaml
    /// </summary>
    public partial class PrinterSelection : Window
    {
        // Key-Value pairs for comboboxes
        public ObservableCollection<KeyValuePair<string, PrinterHelper.ColorMode>> Color { get; set; }
        public ObservableCollection<KeyValuePair<string, PrinterHelper.PageDuplex>> Duplex { get; set; }

        // contain user selected values
        public string Printer { get; set; }
        public PrinterHelper.ColorMode ColorMode { get; set; }
        public PrinterHelper.PageDuplex PageDuplex { get; set; }

        public PrinterSelection()
        {
            InitializeComponent();

            // data context to find the Key-Value pairs
            DataContext = this;

            // define Key-Value pairs for color combobox
            Color = new ObservableCollection<KeyValuePair<string, PrinterHelper.ColorMode>>()
            {
                new KeyValuePair<string, PrinterHelper.ColorMode>(
                    "Schwarzweiß", PrinterHelper.ColorMode.DMCOLOR_MONOCHROME),
                new KeyValuePair<string, PrinterHelper.ColorMode>(
                    "Farbe",PrinterHelper.ColorMode.DMCOLOR_COLOR)
            };

            // define Key-Value pairs for duplex combobox
            Duplex = new ObservableCollection<KeyValuePair<string, PrinterHelper.PageDuplex>>()
            {
                new KeyValuePair<string, PrinterHelper.PageDuplex>(
                    "Kein",PrinterHelper.PageDuplex.DMDUP_SIMPLEX),
                new KeyValuePair<string, PrinterHelper.PageDuplex>(
                    "Lange Seite", PrinterHelper.PageDuplex.DMDUP_VERTICAL),
                new KeyValuePair<string, PrinterHelper.PageDuplex>(
                    "Kurze Seite", PrinterHelper.PageDuplex.DMDUP_HORIZONTAL)
            };

            InitPrinterList();
        }

        private void InitPrinterList()
        {
            // add all printers to the selection list
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                PrinterList.Items.Add(printer);
            }
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            // first check, if color mode and duplex mode is selected
            if (ColorSelection.SelectedIndex < 0 || DuplexSelection.SelectedIndex < 0)
            {
                MessageBox.Show
                (
                    this,
                    "Farb- und Duplex-Modus nicht gewählt! " +
                    "Bitte wählen Sie sowohl den Farbmodus als auch den Duplexmodus aus.",
                    "Jagdorganisation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );

                return;
            }

            // get all user selected values
            Printer = PrinterList.SelectedItem.ToString();
            ColorMode = (PrinterHelper.ColorMode)ColorSelection.SelectedValue;
            PageDuplex = (PrinterHelper.PageDuplex)DuplexSelection.SelectedValue;

            DialogResult = true;
            Close();
        }
    }
}
