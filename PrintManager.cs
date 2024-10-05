﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jagdorganisation
{
    class PrintManager
    {
        private string _session_printer; // default printer for this session
        private readonly string _default_printer; // user default printer

        public PrintManager()
        {
            // save user default printer
            _default_printer = PrinterHelper.GetDefaultPrinterName();
        }

        ~PrintManager()
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
