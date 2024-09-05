using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;

namespace Jagdorganisation
{
    class HunterGroupPrinter
    {
        private xl.Application _xlApp;
        private xl.Workbooks _workbooks;
        private xl.Workbook _source_workbook;
        private xl.Workbook _temp_workbook;
        public HunterGroupPrinter()
        {
            // every new created workbook has only one sheet
            _xlApp = new xl.Application { SheetsInNewWorkbook = 1 };
            _workbooks = _xlApp.Workbooks;
        }

        ~HunterGroupPrinter()
        {
            // close excel connection and all workbooks
            _temp_workbook.Close(false);
            Marshal.FinalReleaseComObject(_temp_workbook);
            _temp_workbook = null;

            _source_workbook.Close(false);
            Marshal.FinalReleaseComObject(_source_workbook);
            _source_workbook = null;

            _workbooks.Close();
            Marshal.FinalReleaseComObject(_workbooks);
            _workbooks = null;

            _xlApp.Quit();
            Marshal.FinalReleaseComObject(_xlApp);
            _xlApp = null;
        }

        public void CreateCardsFromSource(string source_file)
        {
            // open source file with all data
            _source_workbook = _workbooks.Open(source_file);
            xl.Worksheet source_sheet = _source_workbook.Sheets["einteilung"];

            // create temporary workbook with only one sheet
            _temp_workbook = _workbooks.Add();

            // set up first sheet as separator
            _temp_workbook.Sheets[1].Name = "Trennblatt";
            _temp_workbook.Sheets[1].Cells[1, 1].Value = "Trennblatttext";
            _temp_workbook.Sheets[1].Cells[1, 1].Font.Name = "Arial";
            _temp_workbook.Sheets[1].Cells[1, 1].Font.Size = 72;
            _temp_workbook.Sheets[1].PageSetup.CenterHorizontally = true;
            _temp_workbook.Sheets[1].PageSetup.CenterVertically = true;
            _temp_workbook.Sheets[1].PageSetup.Orientation = xl.XlPageOrientation.xlLandscape;

            // copy hunter sheet for every group
            xl.Range source_range = source_sheet.Range["C6:D35"]; // JAEGERGRUPPEN
            foreach (xl.Range cell in source_range)
            {
                if (cell.Text != "")
                {
                    // copy sheet after the last sheet in workbook
                    xl.Worksheet last_sheet = _temp_workbook.Sheets[_temp_workbook.Sheets.Count];
                    _source_workbook.Sheets["standkarte"].Copy(Before: last_sheet);

                    // get the active sheet as copied sheet
                    xl.Worksheet copied_sheet = _temp_workbook.ActiveSheet;

                    // unlock sheet if locked
                    // new locking is not necessary
                    if (copied_sheet.ProtectContents) { copied_sheet.Unprotect("pljagdfa39"); }

                    // prepare values in sheet
                    copied_sheet.Name = cell.Text;
                    copied_sheet.Range["B15"].Value2 = source_sheet.Range["A" + cell.Row].Value2; // NUMMER
                    copied_sheet.Range["C15"].Value2 = cell.Text; // ANSTELLER
                }
            }
        }

        public void PrintCards(string group, bool? separator)
        {
            string title_text = "";
            string data_range = "";
            switch (group)
            {
                case "Ansteller":
                    title_text = "Ansteller";
                    data_range = "E6:E35";
                    break;

                case "Standschützen":
                    title_text = "Standkarte";
                    data_range = "G6:G35";
                    break;

                case "Hundestände":
                    title_text = "Hundestand";
                    data_range = "I6:I35";
                    break;

                case "Ersatzstände":
                    title_text = "Ersatzstand";
                    data_range = "J6:J35";
                    break;
            }

            // print out separator
            if (separator == true) { PrintSeparator(group); }

            // print sheets for specified group
            PrintSheets(data_range, title_text);
        }

        private void PrintSeparator(string separator_text)
        {
            // change separator text
            _temp_workbook.Sheets["trennblatt"].Cells[1, 1].Value = separator_text;

            // print out
            _temp_workbook.Sheets["trennblatt"].PrintOut();
        }

        private void PrintSheets(string data_range, string title_text)
        {
            xl.Range print_range = _source_workbook.Sheets["einteilung"].Range[data_range];
            foreach (xl.Range cell in print_range)
            {
                double? copies = cell.Value2;
                if (copies > 0)
                {
                    int row = cell.Row - print_range.Row + 1;
                    _temp_workbook.Sheets[row].Shapes.Item["TextField"].TextFrame.Characters.Text = title_text;
                    _temp_workbook.Sheets[row].PrintOut(Copies: copies);
                }
            }
        }
    }
}
