using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Security;
using System.ComponentModel;

namespace Jagdorganisation
{
    /// <summary>
    /// Origin: http://blog.csdn.net/csui2008/article/details/5718461
    /// Modified and a little tested by Huan-Lin Tsai. May-29-2013.
    /// </summary>
    public static class PrinterHelper
    {
        #region "Private Variables"
        private static int lastError;
        private static int nRet;   //long 
        private static int intError;
        private static System.Int32 nJunk;

        #endregion

        #region "API Define"

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string printerName);


        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true,
             ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        //        [DllImport("winspool.Drv", EntryPoint="DocumentPropertiesA", SetLastError=true, 
        //             ExactSpelling=true, CallingConvention=CallingConvention.StdCall)]
        //        private static extern int DocumentProperties (IntPtr hwnd, IntPtr hPrinter, 
        //            [MarshalAs(UnmanagedType.LPStr)] string pDeviceNameg, 
        //            IntPtr pDevModeOutput, ref IntPtr pDevModeInput, int fMode);
        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesA", SetLastError = true,
             ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        internal static extern int DocumentProperties(
            IntPtr hwnd,
            IntPtr hPrinter,
            [MarshalAs(UnmanagedType.LPStr)] string pDeviceName,
            IntPtr pDevModeOutput,
            IntPtr pDevModeInput,
            int fMode
            );

        [DllImport("winspool.Drv", EntryPoint = "GetPrinterA", SetLastError = true,
             CharSet = CharSet.Ansi, ExactSpelling = true,
             CallingConvention = CallingConvention.StdCall)]
        private static extern bool GetPrinter(IntPtr hPrinter, Int32 dwLevel,
            IntPtr pPrinter, Int32 dwBuf, out Int32 dwNeeded);

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA",
             SetLastError = true, CharSet = CharSet.Ansi,
             ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool
            OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter,
            out IntPtr hPrinter, IntPtr pDefault); //ref PRINTER_DEFAULTS pd);

        [DllImport("winspool.drv", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern bool SetPrinter(IntPtr hPrinter, int Level, IntPtr
            pPrinter, int Command);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int size);

        [DllImport("GDI32.dll", EntryPoint = "CreateDC", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false,
             CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern IntPtr CreateDC([MarshalAs(UnmanagedType.LPTStr)]
            string pDrive,
            [MarshalAs(UnmanagedType.LPTStr)] string pName,
            [MarshalAs(UnmanagedType.LPTStr)] string pOutput,
            ref DEVMODE pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "ResetDC", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false,
             CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern IntPtr ResetDC(
            IntPtr hDC,
            ref DEVMODE
            pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "DeleteDC", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false,
             CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool DeleteDC(IntPtr hDC);

        [DllImport("winspool.drv", EntryPoint = "DeviceCapabilitiesA", SetLastError = true)]
        internal static extern Int32 DeviceCapabilities(
                               [MarshalAs(UnmanagedType.LPStr)] String device,
                               [MarshalAs(UnmanagedType.LPStr)] String port,
                               Int16 capability,
                               IntPtr outputBuffer,
                               IntPtr deviceMode);

        [DllImport("winspool.drv", SetLastError = true)]
        internal static extern bool EnumPrintersW(Int32 flags,
            [MarshalAs(UnmanagedType.LPTStr)] string printerName,
            Int32 level, IntPtr buffer, Int32 bufferSize, out Int32
            requiredBufferSize, out Int32 numPrintersReturned);

        [DllImport("kernel32.dll", EntryPoint = "GetLastError", SetLastError = false,
             ExactSpelling = true, CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern Int32 GetLastError();

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr windowHandle,
            uint Msg,
            IntPtr wParam,
            IntPtr lParam,
            SendMessageTimeoutFlags flags,
            uint timeout,
            out IntPtr result
            );

        #endregion

        #region "Data structure"

        /// <summary>
        ///  紙張存取權限等信息
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct PRINTER_DEFAULTS
        {
            public int pDatatype;
            public int pDevMode;
            public int DesiredAccess;//對印表機的存取權限
        }


        //紙張方向
        public enum PageOrientation
        {
            DMORIENT_PORTRAIT = 1,//直向
            DMORIENT_LANDSCAPE = 2,//橫向
        }

        /// <summary>
        /// 紙張類型
        /// </summary>
        public enum PaperSize
        {
            DMPAPER_LETTER = 1, // Letter 8 1/2 x 11 in
            DMPAPER_LETTERSMALL = 2, // Letter Small 8 1/2 x 11 in
            DMPAPER_TABLOID = 3, // Tabloid 11 x 17 in
            DMPAPER_LEDGER = 4, // Ledger 17 x 11 in
            DMPAPER_LEGAL = 5, // Legal 8 1/2 x 14 in
            DMPAPER_STATEMENT = 6, // Statement 5 1/2 x 8 1/2 in
            DMPAPER_EXECUTIVE = 7, // Executive 7 1/4 x 10 1/2 in
            DMPAPER_A3 = 8, // A3 297 x 420 mm
            DMPAPER_A4 = 9, // A4 210 x 297 mm
            DMPAPER_A4SMALL = 10, // A4 Small 210 x 297 mm
            DMPAPER_A5 = 11, // A5 148 x 210 mm
            DMPAPER_B4 = 12, // B4 250 x 354
            DMPAPER_B5 = 13, // B5 182 x 257 mm
            DMPAPER_FOLIO = 14, // Folio 8 1/2 x 13 in
            DMPAPER_QUARTO = 15, // Quarto 215 x 275 mm
            DMPAPER_10X14 = 16, // 10x14 in
            DMPAPER_11X17 = 17, // 11x17 in
            DMPAPER_NOTE = 18, // Note 8 1/2 x 11 in
            DMPAPER_ENV_9 = 19, // Envelope #9 3 7/8 x 8 7/8
            DMPAPER_ENV_10 = 20, // Envelope #10 4 1/8 x 9 1/2
            DMPAPER_ENV_11 = 21, // Envelope #11 4 1/2 x 10 3/8
            DMPAPER_ENV_12 = 22, // Envelope #12 4 /276 x 11
            DMPAPER_ENV_14 = 23, // Envelope #14 5 x 11 1/2
            DMPAPER_CSHEET = 24, // C size sheet
            DMPAPER_DSHEET = 25, // D size sheet
            DMPAPER_ESHEET = 26, // E size sheet
            DMPAPER_ENV_DL = 27, // Envelope DL 110 x 220mm
            DMPAPER_ENV_C5 = 28, // Envelope C5 162 x 229 mm
            DMPAPER_ENV_C3 = 29, // Envelope C3 324 x 458 mm
            DMPAPER_ENV_C4 = 30, // Envelope C4 229 x 324 mm
            DMPAPER_ENV_C6 = 31, // Envelope C6 114 x 162 mm
            DMPAPER_ENV_C65 = 32, // Envelope C65 114 x 229 mm
            DMPAPER_ENV_B4 = 33, // Envelope B4 250 x 353 mm
            DMPAPER_ENV_B5 = 34, // Envelope B5 176 x 250 mm
            DMPAPER_ENV_B6 = 35, // Envelope B6 176 x 125 mm
            DMPAPER_ENV_ITALY = 36, // Envelope 110 x 230 mm
            DMPAPER_ENV_MONARCH = 37, // Envelope Monarch 3.875 x 7.5 in
            DMPAPER_ENV_PERSONAL = 38, // 6 3/4 Envelope 3 5/8 x 6 1/2 in
            DMPAPER_FANFOLD_US = 39, // US Std Fanfold 14 7/8 x 11 in
            DMPAPER_FANFOLD_STD_GERMAN = 40, // German Std Fanfold 8 1/2 x 12 in
            DMPAPER_FANFOLD_LGL_GERMAN = 41, // German Legal Fanfold 8 1/2 x 13 in
            DMPAPER_USER = 256,// user defined
            DMPAPER_FIRST = DMPAPER_LETTER,
            DMPAPER_LAST = DMPAPER_USER,
        }


        /// <summary>
        /// 紙張來源
        /// </summary>
        public enum PaperSource
        {
            DMBIN_UPPER = 1,
            DMBIN_LOWER = 2,
            DMBIN_MIDDLE = 3,
            DMBIN_MANUAL = 4,
            DMBIN_ENVELOPE = 5,
            DMBIN_ENVMANUAL = 6,
            DMBIN_AUTO = 7,
            DMBIN_TRACTOR = 8,
            DMBIN_SMALLFMT = 9,
            DMBIN_LARGEFMT = 10,
            DMBIN_LARGECAPACITY = 11,
            DMBIN_CASSETTE = 14,
            DMBIN_FORMSOURCE = 15,
            DMRES_DRAFT = -1,
            DMRES_LOW = -2,
            DMRES_MEDIUM = -3,
            DMRES_HIGH = -4
        }


        /// <summary>
        /// 是否要雙面列印
        /// </summary>
        public enum PageDuplex
        {
            DMDUP_HORIZONTAL = 3,
            DMDUP_SIMPLEX = 1,
            DMDUP_VERTICAL = 2
        }


        /// <summary>
        /// 需要變更的列印參數
        /// </summary>
        public struct PrinterSettingsInfo
        {
            public PageOrientation Orientation; //列印方向
            public PaperSize Size;              //列印紙張類型（用數字表示 256為用戶自定紙張）
            public PaperSource source;          //紙張來源
            public PageDuplex Duplex;           //是否雙面列印等信息
            public int pLength;                 //紙張的高
            public int pWidth;                  //紙張的寬
            public int pmFields;                //需改變的信息進行"|"運算後的和
            public string pFormName;            //紙張的名字
        }

        //PRINTER_INFO_2 - 印表機信息結構包含 1..9 個等級，詳細信息請參考API
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct PRINTER_INFO_2
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pServerName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pPrinterName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pShareName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pPortName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDriverName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pComment;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pLocation;
            public IntPtr pDevMode;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pSepFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pPrintProcessor;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDatatype;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pParameters;
            public IntPtr pSecurityDescriptor;
            public Int32 Attributes;
            public Int32 Priority;
            public Int32 DefaultPriority;
            public Int32 StartTime;
            public Int32 UntilTime;
            public Int32 Status;
            public Int32 cJobs;
            public Int32 AveragePPM;
        }


        //PRINTER_INFO_5 - 印表機信息結構包含 1..9 個等級，詳細信息請參考API
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct PRINTER_INFO_5
        {
            [MarshalAs(UnmanagedType.LPTStr)]
            public String PrinterName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public String PortName;
            [MarshalAs(UnmanagedType.U4)]
            public Int32 Attributes;
            [MarshalAs(UnmanagedType.U4)]
            public Int32 DeviceNotSelectedTimeout;
            [MarshalAs(UnmanagedType.U4)]
            public Int32 TransmissionRetryTimeout;
        }


        //PRINTER_INFO_9
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct PRINTER_INFO_9
        {
            public IntPtr pDevMode;
        }

        /// <summary>
        /// The DEVMODE data structure contains information about the initialization and environment of a printer or a display device
        ///DEVMODE結構包含了印表機（或顯示設置)的初始化和當前狀態信息,詳細信息請參考API
        /// </summary>
        private const short CCDEVICENAME = 32;
        private const short CCFORMNAME = 32;
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public struct DEVMODE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCDEVICENAME)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public short dmOrientation;
            public short dmPaperSize;
            public short dmPaperLength;
            public short dmPaperWidth;
            public short dmScale;
            public short dmCopies;
            public short dmDefaultSource;
            public short dmPrintQuality;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCFORMNAME)]
            public string dmFormName;
            public short dmUnusedPadding;
            public short dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
        }

        //SendMessageTimeout Flags
        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0000,
            SMTO_BLOCK = 0x0001,
            SMTO_ABORTIFHUNG = 0x0002,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x0008
        }
        #endregion

        #region "const Variables"

        //DEVMODE.dmFields
        const int DM_FORMNAME = 0x10000;//改變紙張名稱時需在dmFields設置此常數
        const int DM_PAPERSIZE = 0x0002;//改變紙張類型時需在dmFields設置此常數
        const int DM_PAPERLENGTH = 0x0004;//改變紙張長度時需在dmFields設置此常數
        const int DM_PAPERWIDTH = 0x0008;//改變紙張寬度時需在dmFields設置此常數
        const int DM_DUPLEX = 0x1000;//改變紙張是否雙面列印時需在dmFields設置此常數
        const int DM_ORIENTATION = 0x0001;//改變紙張方向時需在dmFields設置此常數

        //用於改變DocumentProperties的參數，詳細信息請參考API
        const int DM_IN_BUFFER = 8;
        const int DM_OUT_BUFFER = 2;

        //用於設置對印表機的存取權限
        const int PRINTER_ACCESS_ADMINISTER = 0x4;
        const int PRINTER_ACCESS_USE = 0x8;
        const int STANDARD_RIGHTS_REQUIRED = 0xF0000;
        const int PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED | PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE);

        //得到指定列印的所有紙張
        const int PRINTER_ENUM_LOCAL = 2;
        const int PRINTER_ENUM_CONNECTIONS = 4;
        const int DC_PAPERNAMES = 16;
        const int DC_PAPERS = 2;
        const int DC_PAPERSIZE = 3;

        //sendMessageTimeOut
        const int WM_SETTINGCHANGE = 0x001A;
        const int HWND_BROADCAST = 0xffff;
        #endregion

        #region printer method

        public static bool OpenPrinterEx(string szPrinter, out IntPtr hPrinter, ref PRINTER_DEFAULTS pd)
        {
            bool bRet = OpenPrinter(szPrinter, out hPrinter, IntPtr.Zero);
            return bRet;
        }

        /// <summary>
        /// 獲取指定印表機的設置信息。
        /// </summary>
        /// <param name="PrinterName">印表機名稱</param>
        /// <returns>指定印表機的設置信息</returns>
        public static DEVMODE GetPrinterDevMode(string PrinterName)
        {
            if (PrinterName == string.Empty || PrinterName == null)
            {
                PrinterName = GetDefaultPrinterName();
            }

            PRINTER_DEFAULTS pd = new PRINTER_DEFAULTS();
            pd.pDatatype = 0;
            pd.pDevMode = 0;
            pd.DesiredAccess = PRINTER_ALL_ACCESS;
            // Michael: some printers (e.g. network printer) do not allow PRINTER_ALL_ACCESS and will cause Access Is Denied error.
            // When this happen, try PRINTER_ACCESS_USE.

            IntPtr hPrinter = new System.IntPtr();
            if (!OpenPrinterEx(PrinterName, out hPrinter, ref pd))
            {
                lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            int nBytesNeeded = 0;
            GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out nBytesNeeded);
            if (nBytesNeeded <= 0)
            {
                throw new System.Exception("Unable to allocate memory");
            }


            DEVMODE dm;

            // Allocate enough space for PRINTER_INFO_2... {ptrPrinterIn fo = Marshal.AllocCoTaskMem(nBytesNeeded)};
            IntPtr ptrPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded);

            // The second GetPrinter fills in all the current settings, so all you 
            // need to do is modify what you're interested in...
            nRet = Convert.ToInt32(GetPrinter(hPrinter, 2, ptrPrinterInfo, nBytesNeeded, out nJunk));
            if (nRet == 0)
            {
                lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            PRINTER_INFO_2 pinfo = new PRINTER_INFO_2();
            pinfo = (PRINTER_INFO_2)Marshal.PtrToStructure(ptrPrinterInfo, typeof(PRINTER_INFO_2));
            IntPtr Temp = new IntPtr();
            if (pinfo.pDevMode == IntPtr.Zero)
            {
                // If GetPrinter didn't fill in the DEVMODE, try to get it by calling
                // DocumentProperties...
                IntPtr ptrZero = IntPtr.Zero;
                //get the size of the devmode structure
                int sizeOfDevMode = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, IntPtr.Zero, IntPtr.Zero, 0);

                IntPtr ptrDM = Marshal.AllocCoTaskMem(sizeOfDevMode);
                int i;
                i = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, ptrDM, ptrZero, DM_OUT_BUFFER);
                if ((i < 0) || (ptrDM == IntPtr.Zero))
                {
                    //Cannot get the DEVMODE structure.
                    throw new System.Exception("Cannot get DEVMODE data");
                }
                pinfo.pDevMode = ptrDM;
            }
            intError = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, IntPtr.Zero, Temp, 0);

            //IntPtr yDevModeData = Marshal.AllocCoTaskMem(i1);
            IntPtr yDevModeData = Marshal.AllocHGlobal(intError);
            intError = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, yDevModeData, Temp, 2);
            dm = (DEVMODE)Marshal.PtrToStructure(yDevModeData, typeof(DEVMODE));//從記憶空間中取出印表機設備信息
            //nRet = DocumentProperties(IntPtr.Zero, hPrinter, sPrinterName, yDevModeData
            // , ref yDevModeData, (DM_IN_BUFFER | DM_OUT_BUFFER));
            if ((nRet == 0) || (hPrinter == IntPtr.Zero))
            {
                lastError = Marshal.GetLastWin32Error();
                //string myErrMsg = GetErrorMessage(lastError);
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            ClosePrinter(hPrinter);

            return dm;
        }


        /// <summary>
        /// 判斷目前預設印表機之特定紙張是否等於傳入之大小。
        /// </summary>
        /// <param name="FormName">紙張名稱。</param>
        /// <param name="width">寬。Unit: 1/10 of a millimeter.</param>
        /// <param name="length">高。Unit: 1/10 of a millimeter.</param>
        /// <returns>如果預設印表機的 DEVMODE 結構中的紙張大小與指定之 width 和 height 相同則傳回 true，否則傳回 false。</returns>
        public static bool IsPaperSize(string FormName, int width, int length)
        {
            DEVMODE dm = PrinterHelper.GetPrinterDevMode(null);
            if (FormName == dm.dmFormName && width == dm.dmPaperWidth && length == dm.dmPaperLength)
                return true;
            else
                return false;
        }

        /// <summary>
        /// 改變印表機的設定。
        /// </summary>
        /// <param name="printerName">印表機的名字,如果為空，自動得到預設印表機的名字</param>
        /// <param name="prnSettings">需改變信息</param>
        /// <returns>是否改變成功</returns>
        public static void ModifyPrinterSettings(string printerName, ref PrinterSettingsInfo prnSettings)
        {
            PRINTER_INFO_9 printerInfo;
            printerInfo.pDevMode = IntPtr.Zero;
            if (String.IsNullOrEmpty(printerName))
            {
                printerName = GetDefaultPrinterName();
            }

            IntPtr hPrinter = new System.IntPtr();

            PRINTER_DEFAULTS prnDefaults = new PRINTER_DEFAULTS();
            prnDefaults.pDatatype = 0;
            prnDefaults.pDevMode = 0;
            prnDefaults.DesiredAccess = PRINTER_ALL_ACCESS;

            if (!OpenPrinterEx(printerName, out hPrinter, ref prnDefaults))
            {
                return;
            }

            IntPtr ptrPrinterInfo = IntPtr.Zero;
            try
            {
                //得到結構體DEVMODE的大小
                int iDevModeSize = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);
                if (iDevModeSize < 0)
                    throw new ApplicationException("Cannot get the size of the DEVMODE structure.");

                //分配指向結構體DEVMODE的記憶空間緩沖區
                IntPtr hDevMode = Marshal.AllocCoTaskMem(iDevModeSize + 100);

                //得到一個指向 DEVMODE 結構的指標
                nRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, hDevMode, IntPtr.Zero, DM_OUT_BUFFER);
                if (nRet < 0)
                    throw new ApplicationException("Cannot get the size of the DEVMODE structure.");
                //給dm賦值
                DEVMODE dm = (DEVMODE)Marshal.PtrToStructure(hDevMode, typeof(DEVMODE));

                if ((((int)prnSettings.Duplex < 0) || ((int)prnSettings.Duplex > 3)))
                {
                    throw new ArgumentOutOfRangeException("nDuplexSetting", "nDuplexSetting is incorrect.");
                }
                else
                {
                    // 更改印表機設定
                    if ((int)prnSettings.Size != 0) //是否改變紙張類型
                    {
                        dm.dmPaperSize = (short)prnSettings.Size;
                        dm.dmFields |= DM_PAPERSIZE;
                    }
                    if (prnSettings.pWidth != 0)    //是否改變紙張寬度
                    {
                        dm.dmPaperWidth = (short)prnSettings.pWidth;
                        dm.dmFields |= DM_PAPERWIDTH;
                    }
                    if (prnSettings.pLength != 0)   //是否改變紙張高度
                    {
                        dm.dmPaperLength = (short)prnSettings.pLength;
                        dm.dmFields |= DM_PAPERLENGTH;
                    }
                    if (!String.IsNullOrEmpty(prnSettings.pFormName))    //是否改變紙張名稱
                    {
                        dm.dmFormName = prnSettings.pFormName;
                        dm.dmFields |= DM_FORMNAME;
                    }
                    if ((int)prnSettings.Orientation != 0)  //是否改變紙張方向
                    {
                        dm.dmOrientation = (short)prnSettings.Orientation;
                        dm.dmFields |= DM_ORIENTATION;
                    }
                    Marshal.StructureToPtr(dm, hDevMode, true);

                    //得到 printer info 的大小
                    nRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, printerInfo.pDevMode, printerInfo.pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);
                    if (nRet < 0)
                    {
                        throw new ApplicationException("Unable to set the PrintSetting for this printer");
                    }
                    int nBytesNeeded = 0;
                    GetPrinter(hPrinter, 9, IntPtr.Zero, 0, out nBytesNeeded);
                    if (nBytesNeeded == 0)
                        throw new ApplicationException("GetPrinter failed.Couldn't get the nBytesNeeded for shared PRINTER_INFO_9 structure");

                    //配置記憶體區塊
                    ptrPrinterInfo = Marshal.AllocCoTaskMem(nBytesNeeded);
                    bool bSuccess = GetPrinter(hPrinter, 9, ptrPrinterInfo, nBytesNeeded, out nJunk);
                    if (!bSuccess)
                        throw new ApplicationException("GetPrinter failed.Couldn't get the nBytesNeeded for shared PRINTER_INFO_9 structure");
                    //賦值給printerInfo
                    printerInfo = (PRINTER_INFO_9)Marshal.PtrToStructure(ptrPrinterInfo, printerInfo.GetType());
                    printerInfo.pDevMode = hDevMode;

                    //獲取一個指向 PRINTER_INFO_9 結構的指標
                    Marshal.StructureToPtr(printerInfo, ptrPrinterInfo, true);

                    //設置印表機
                    bSuccess = SetPrinter(hPrinter, 9, ptrPrinterInfo, 0);
                    if (!bSuccess)
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "SetPrinter() failed.Couldn't set the printer settings");

                    // 通知其它 app，印表機設定已經更改 -- Do NOT use because it causes app halt serveral seconds!!
                    /*
                    PrinterHelper.SendMessageTimeout(
                        new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, IntPtr.Zero,
                        PrinterHelper.SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out hDummy);
                     */
                }
            }
            finally
            {
                ClosePrinter(hPrinter);

                //釋放
                if (ptrPrinterInfo == IntPtr.Zero)
                    Marshal.FreeHGlobal(ptrPrinterInfo);
                if (hPrinter == IntPtr.Zero)
                    Marshal.FreeHGlobal(hPrinter);
            }
        }


        /// <summary>
        /// 改變印表機設定的另一個版本。測試過程中曾出現應用程式異常終止而且無任何錯誤訊息。請使用 ModifyPrinterSettings。
        /// </summary>
        /// <param name="printerName">印表機名稱。傳入 null 或空字串表示使用預設印表機。</param>
        /// <param name="PS">需改變信息</param>
        /// <returns>是否改變成功</returns>
        public static bool ModifyPrinterSettings_V2(string printerName, ref PrinterSettingsInfo PS)
        {
            PRINTER_DEFAULTS pd = new PRINTER_DEFAULTS();
            pd.pDatatype = 0;
            pd.pDevMode = 0;
            pd.DesiredAccess = PRINTER_ALL_ACCESS;
            if (String.IsNullOrEmpty(printerName))
            {
                printerName = GetDefaultPrinterName();
            }

            IntPtr hPrinter = new System.IntPtr();

            if (!OpenPrinterEx(printerName, out hPrinter, ref pd))
            {
                lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            //呼叫GetPrinter來獲取PRINTER_INFO_2在記憶空間的 bytes 數
            int nBytesNeeded = 0;
            GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out nBytesNeeded);
            if (nBytesNeeded <= 0)
            {
                ClosePrinter(hPrinter);
                return false;
            }
            //為PRINTER_INFO_2分配足夠的記憶空間
            IntPtr ptrPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded);
            if (ptrPrinterInfo == IntPtr.Zero)
            {
                ClosePrinter(hPrinter);
                return false;
            }

            //呼叫GetPrinter填充所的當前設定，也就是你所想改變的信息（ptrPrinterInfo中）
            if (!GetPrinter(hPrinter, 2, ptrPrinterInfo, nBytesNeeded, out nBytesNeeded))
            {
                Marshal.FreeHGlobal(ptrPrinterInfo);
                ClosePrinter(hPrinter);
                return false;
            }
            //把記憶區塊中指向 PRINTER_INFO_2 的指標轉化為 PRINTER_INFO_2 結構
            //如果 GetPrinter 沒有得到 DEVMODE 結構，將嘗試透過 DocumentProperties 來取得 DEVMODE 結構
            PRINTER_INFO_2 pinfo = new PRINTER_INFO_2();
            pinfo = (PRINTER_INFO_2)Marshal.PtrToStructure(ptrPrinterInfo, typeof(PRINTER_INFO_2));
            IntPtr Temp = new IntPtr();
            if (pinfo.pDevMode == IntPtr.Zero)
            {
                // If GetPrinter didn't fill in the DEVMODE, try to get it by calling
                // DocumentProperties...
                IntPtr ptrZero = IntPtr.Zero;
                //get the size of the devmode structure
                nBytesNeeded = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);
                if (nBytesNeeded <= 0)
                {
                    Marshal.FreeHGlobal(ptrPrinterInfo);
                    ClosePrinter(hPrinter);
                    return false;
                }
                IntPtr ptrDM = Marshal.AllocCoTaskMem(nBytesNeeded);
                int i;
                i = DocumentProperties(IntPtr.Zero, hPrinter, printerName, ptrDM, ptrZero, DM_OUT_BUFFER);
                if ((i < 0) || (ptrDM == IntPtr.Zero))
                {
                    //Cannot get the DEVMODE structure.
                    Marshal.FreeHGlobal(ptrDM);
                    ClosePrinter(ptrPrinterInfo);
                    return false;
                }
                pinfo.pDevMode = ptrDM;
            }
            DEVMODE dm = (DEVMODE)Marshal.PtrToStructure(pinfo.pDevMode, typeof(DEVMODE));

            //修改印表機的設定信息        
            if ((((int)PS.Duplex < 0) || ((int)PS.Duplex > 3)))
            {
                throw new ArgumentOutOfRangeException("nDuplexSetting", "nDuplexSetting is incorrect.");
            }
            else
            {
                if (String.IsNullOrEmpty(printerName))
                {
                    printerName = GetDefaultPrinterName();
                }
                if ((int)PS.Size != 0)//是否改變紙張類型
                {
                    dm.dmPaperSize = (short)PS.Size;
                    dm.dmFields |= DM_PAPERSIZE;
                }
                if (PS.pWidth != 0)//是否改變紙張寬度
                {
                    dm.dmPaperWidth = (short)PS.pWidth;
                    dm.dmFields |= DM_PAPERWIDTH;
                }
                if (PS.pLength != 0)//是否改變紙張高度
                {
                    dm.dmPaperLength = (short)PS.pLength;
                    dm.dmFields |= DM_PAPERLENGTH;
                }
                if (!String.IsNullOrEmpty(PS.pFormName))    //是否改變紙張名稱
                {
                    dm.dmFormName = PS.pFormName;
                    dm.dmFields |= DM_FORMNAME;
                }
                if ((int)PS.Orientation != 0)//是否改變紙張方向
                {
                    dm.dmOrientation = (short)PS.Orientation;
                    dm.dmFields |= DM_ORIENTATION;
                }
                Marshal.StructureToPtr(dm, pinfo.pDevMode, true);
                Marshal.StructureToPtr(pinfo, ptrPrinterInfo, true);
                pinfo.pSecurityDescriptor = IntPtr.Zero;
                //Make sure the driver_Dependent part of devmode is updated...
                nRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, pinfo.pDevMode, pinfo.pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);
                if (nRet <= 0)
                {
                    Marshal.FreeHGlobal(ptrPrinterInfo);
                    ClosePrinter(hPrinter);
                    return false;
                }

                //SetPrinter 更新印表機信息
                if (!SetPrinter(hPrinter, 2, ptrPrinterInfo, 0))
                {
                    Marshal.FreeHGlobal(ptrPrinterInfo);
                    ClosePrinter(hPrinter);
                    return false;
                }
                //通知其它應用程序，印表機信息已經更改
                IntPtr hDummy = IntPtr.Zero;
                PrinterHelper.SendMessageTimeout(
                    new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, IntPtr.Zero,
                    PrinterHelper.SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out hDummy);

                //釋放
                if (ptrPrinterInfo == IntPtr.Zero)
                    Marshal.FreeHGlobal(ptrPrinterInfo);
                if (hPrinter == IntPtr.Zero)
                    Marshal.FreeHGlobal(hPrinter);

                return true;

            }
        }


        /// <summary>
        /// 得到預設印表機的名字
        /// </summary>
        /// <returns>返回預設印表機的名字</returns>
        public static string GetDefaultPrinterName()
        {
            StringBuilder dp = new StringBuilder(256);
            int size = dp.Capacity;
            if (GetDefaultPrinter(dp, ref size))
            {
                return dp.ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 得到紙張的kind，如果為 0 則錯誤。
        /// </summary>
        /// <param name="printerName">印表機名稱。傳入 null 或空字串表示使用預設印表機。</param>
        /// <param name="paperName">紙張名稱，一定要填</param>
        /// <returns>kind</returns>
        public static short GetOnePaper(string printerName, string paperName)
        {

            short kind = 0;
            if (String.IsNullOrEmpty(printerName))
                printerName = GetDefaultPrinterName();
            PRINTER_INFO_5 info5;
            int requiredSize;
            int numPrinters;
            bool foundPrinter = EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                string.Empty, 5, IntPtr.Zero, 0, out requiredSize, out numPrinters);

            int info5Size = requiredSize;
            IntPtr info5Ptr = Marshal.AllocHGlobal(info5Size);
            IntPtr buffer = IntPtr.Zero;
            try
            {
                foundPrinter = EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                    string.Empty, 5, info5Ptr, info5Size, out requiredSize, out numPrinters);

                string port = null;
                for (int i = 0; i < numPrinters; i++)
                {
                    info5 = (PRINTER_INFO_5)Marshal.PtrToStructure(
                        (IntPtr)((i * Marshal.SizeOf(typeof(PRINTER_INFO_5))) + (int)info5Ptr),
                        typeof(PRINTER_INFO_5));
                    if (info5.PrinterName == printerName)
                    {
                        port = info5.PortName;
                    }
                }

                int numNames = DeviceCapabilities(printerName, port, DC_PAPERNAMES, IntPtr.Zero, IntPtr.Zero);
                if (numNames < 0)
                {
                    int errorCode = GetLastError();
                    Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                    return 0;
                }

                buffer = Marshal.AllocHGlobal(numNames * 64);
                numNames = DeviceCapabilities(printerName, port, DC_PAPERNAMES, buffer, IntPtr.Zero);
                if (numNames < 0)
                {
                    int errorCode = GetLastError();
                    Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                    return 0;
                }
                string[] names = new string[numNames];
                for (int i = 0; i < numNames; i++)
                {
                    names[i] = Marshal.PtrToStringAnsi((IntPtr)((i * 64) + (int)buffer));
                }
                Marshal.FreeHGlobal(buffer);
                buffer = IntPtr.Zero;

                int numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, IntPtr.Zero, IntPtr.Zero);
                if (numPapers < 0)
                {
                    Console.WriteLine("No papers");
                    return 0;
                }

                buffer = Marshal.AllocHGlobal(numPapers * 2);
                numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, buffer, IntPtr.Zero);
                if (numPapers < 0)
                {
                    Console.WriteLine("No papers");
                    return 0;
                }
                short[] kinds = new short[numPapers];
                for (int i = 0; i < numPapers; i++)
                {
                    kinds[i] = Marshal.ReadInt16(buffer, i * 2);
                }

                for (int i = 0; i < numPapers; i++)
                {
                    //                    Console.WriteLine("Paper {0} : {1}", kinds[i], names[i]);
                    if (names[i] == paperName)
                        kind = kinds[i];
                    break;
                }
            }
            finally
            {
                Marshal.FreeHGlobal(info5Ptr);
            }
            return kind;
        }


        /// <summary>
        /// 取得所有可用的紙張，並將紙張規格與名稱輸出至 console。
        /// </summary>
        /// <param name="printerName">印表機名稱。傳入 null 或空字串表示使用預設印表機。</param>
        public static void ShowPapers(string printerName)
        {
            if (String.IsNullOrEmpty(printerName))
            {
                printerName = GetDefaultPrinterName();
            }

            PRINTER_INFO_5 info5;
            int requiredSize;
            int numPrinters;
            bool foundPrinter = EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                string.Empty, 5, IntPtr.Zero, 0, out requiredSize, out numPrinters);

            int info5Size = requiredSize;
            IntPtr info5Ptr = Marshal.AllocHGlobal(info5Size);
            IntPtr buffer = IntPtr.Zero;
            try
            {
                foundPrinter = EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                    string.Empty, 5, info5Ptr, info5Size, out requiredSize, out numPrinters);

                string port = null;
                for (int i = 0; i < numPrinters; i++)
                {
                    info5 = (PRINTER_INFO_5)Marshal.PtrToStructure(
                        (IntPtr)((i * Marshal.SizeOf(typeof(PRINTER_INFO_5))) + (int)info5Ptr),
                        typeof(PRINTER_INFO_5));
                    if (info5.PrinterName == printerName)
                    {
                        port = info5.PortName;
                    }
                }

                int numNames = DeviceCapabilities(printerName, port, DC_PAPERNAMES, IntPtr.Zero, IntPtr.Zero);
                if (numNames < 0)
                {
                    int errorCode = GetLastError();
                    Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                    return;
                }

                buffer = Marshal.AllocHGlobal(numNames * 64);
                numNames = DeviceCapabilities(printerName, port, DC_PAPERNAMES, buffer, IntPtr.Zero);
                if (numNames < 0)
                {
                    int errorCode = GetLastError();
                    Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                    return;
                }
                string[] names = new string[numNames];
                for (int i = 0; i < numNames; i++)
                {
                    names[i] = Marshal.PtrToStringAnsi((IntPtr)((i * 64) + (int)buffer));
                }
                Marshal.FreeHGlobal(buffer);
                buffer = IntPtr.Zero;

                int numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, IntPtr.Zero, IntPtr.Zero);
                if (numPapers < 0)
                {
                    Console.WriteLine("No papers");
                    return;
                }

                buffer = Marshal.AllocHGlobal(numPapers * 2);
                numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, buffer, IntPtr.Zero);
                if (numPapers < 0)
                {
                    Console.WriteLine("No papers");
                    return;
                }
                short[] kinds = new short[numPapers];
                for (int i = 0; i < numPapers; i++)
                {
                    kinds[i] = Marshal.ReadInt16(buffer, i * 2);
                }

                for (int i = 0; i < numPapers; i++)
                {
                    Console.WriteLine("Paper {0} : {1}", kinds[i], names[i]);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(info5Ptr);
            }
        }

        #endregion
    }
}