using System.Windows;

using System.Drawing.Printing;

namespace Jagdorganisation
{
    /// <summary>
    /// Interaktionslogik für PrinterSelection.xaml
    /// </summary>
    public partial class PrinterSelection : Window
    {
        public string SelectedPrinter { get; set; }
        public PrinterSelection()
        {
            InitializeComponent();
            InitPrinterList();
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
            SelectedPrinter = PrinterList.SelectedItem.ToString();

            DialogResult = true;
            Close();
        }
    }
}
