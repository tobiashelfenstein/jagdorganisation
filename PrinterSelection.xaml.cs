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
        // Key-Value pair for comboboxes
        public ObservableCollection<KeyValuePair<string, PrinterHelper.ColorMode>> Color { get; set; }
        public ObservableCollection<KeyValuePair<string, PrinterHelper.PageDuplex>> Duplex { get; set; }

        private PrintSettings _printsettings;
        public PrinterSelection()
        {

            _printsettings = new PrintSettings();

            Color = new ObservableCollection<KeyValuePair<string, PrinterHelper.ColorMode>>()
            {
                new KeyValuePair<string, PrinterHelper.ColorMode>("Schwarzweiß", PrinterHelper.ColorMode.DMCOLOR_MONOCHROME),
                new KeyValuePair<string, PrinterHelper.ColorMode>("Farbe", PrinterHelper.ColorMode.DMCOLOR_COLOR)
            };

            Duplex = new ObservableCollection<KeyValuePair<string, PrinterHelper.PageDuplex>>()
            {
                new KeyValuePair<string, PrinterHelper.PageDuplex>("Kein", PrinterHelper.PageDuplex.DMDUP_SIMPLEX),
                new KeyValuePair<string, PrinterHelper.PageDuplex>("Lange Seite", PrinterHelper.PageDuplex.DMDUP_VERTICAL),
                new KeyValuePair<string, PrinterHelper.PageDuplex>("Kurze Seite", PrinterHelper.PageDuplex.DMDUP_HORIZONTAL)
            };

            InitializeComponent();

            DataContext = this;

            InitPrinterList();
        }

        public PrintSettings GetPrintSettings()
        {
            // returns a copy no reference of the print settings
            return _printsettings;
        }

        private void InitPrinterList()
        {
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                PrinterList.Items.Add(printer);
            }
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            _printsettings.SetSessionPrinter(PrinterList.SelectedItem.ToString());
            _printsettings.SetPrinterSettings(
                (PrinterHelper.ColorMode)ColorSelection.SelectedValue,
                (PrinterHelper.PageDuplex)DuplexSelection.SelectedValue
            );

            DialogResult = true;
            Close();
        }
    }
}
