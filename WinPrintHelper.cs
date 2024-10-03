using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Jagdorganisation
{
    // see https://gist.github.com/huanlin/5671168
    public static class WinPrintHelper
    {
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string printer);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool GetDefaultPrinter(StringBuilder printer, ref int size);

        [DllImport("winspool.drv", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        internal static extern int DocumentProperties(
            IntPtr hwnd,
            IntPtr hPrinter,
            [MarshalAs(UnmanagedType.LPStr)] string pDeviceName,
            IntPtr pDevModeOutput,
            IntPtr pDevModeInput,
            int fMode
            );
    }

    // duplex
    // https://learn.microsoft.com/en-us/windows/win32/printdocs/documentproperties

    // set default printer
    // reset default printer
    // https://stackoverflow.com/questions/971604/how-do-i-set-the-windows-default-printer-in-c?rq=4
}
