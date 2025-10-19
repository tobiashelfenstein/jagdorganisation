using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jagdorganisation
{
    public class PrintSettings
    {
        public string SessionPrinter { get; set; } // default printer for this session
        private readonly string _default_printer; // user default printer

        public PrinterHelper.ColorMode SessionColor { get; set; }
        private PrinterHelper.ColorMode _default_color;

        public PrinterHelper.PageDuplex SessionDuplex { get; set; }
        private PrinterHelper.PageDuplex _default_duplex;

        public PrintSettings()
        {
            // save user default printer
            // color and duplex are printer specific
            _default_printer = PrinterHelper.GetDefaultPrinterName();
        }

        ~PrintSettings()
        {
        }

        public void ResetDefaultPrinter()
        {
            PrinterHelper.SetDefaultPrinter(_default_printer);
        }

        public void ActivateSessionPrinter()
        {
            // first set user selected printer as new default printer
            // the save the printer settings
            PrinterHelper.SetDefaultPrinter(SessionPrinter);
            SavePrinterSettings();

            // modify the printer with new settings for duplex and color
            SetPrinterSettings(SessionDuplex, SessionColor);
        }

        public void SavePrinterSettings()
        {
            var devMode = PrinterHelper.GetPrinterDevMode(null);

            _default_duplex = (PrinterHelper.PageDuplex)devMode.dmDuplex;
            _default_color = (PrinterHelper.ColorMode)devMode.dmColor;
        }

        public void RestorePrinterSettings()
        {
            SetPrinterSettings(_default_duplex, _default_color);
        }

        private void SetPrinterSettings(PrinterHelper.PageDuplex duplex, PrinterHelper.ColorMode color)
        {
            PrinterHelper.PrinterSettingsInfo settings = new PrinterHelper.PrinterSettingsInfo
            {
                Duplex = duplex,
                Color = color
            };

            PrinterHelper.ModifyPrinterSettings(SessionPrinter, ref settings);
        }
    }
}
