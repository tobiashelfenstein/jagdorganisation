using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jagdorganisation
{
    public class PrintSettings
    {
        private string _session_printer; // default printer for this session
        private readonly string _default_printer; // user default printer

        private PrinterHelper.ColorMode _session_color;
        private readonly PrinterHelper.ColorMode _default_color;

        private PrinterHelper.PageDuplex _session_duplex;
        private readonly PrinterHelper.PageDuplex _default_duplex;

        public PrintSettings()
        {
            // save user default printer
            _default_printer = PrinterHelper.GetDefaultPrinterName();
        }

        ~PrintSettings()
        {
        }

        public void SetSessionPrinter(string printer)
        {
            _session_printer = printer;
            PrinterHelper.SetDefaultPrinter(printer);
        }

        public void ResetDefaultPrinter()
        {
            PrinterHelper.SetDefaultPrinter(_default_printer);
        }

        public void SetPrinterSettings(PrinterHelper.ColorMode color, PrinterHelper.PageDuplex duplex)
        {
            PrinterHelper.PrinterSettingsInfo settings = new PrinterHelper.PrinterSettingsInfo
            {
                Duplex = duplex,
                Color = color
            };

            PrinterHelper.ModifyPrinterSettings(_session_printer, ref settings);
        }
    }
}
