using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using xl = Microsoft.Office.Interop.Excel;

namespace Jagdorganisation
{
    public class HunterGroupPrinter
    {
        // define unwanted characters for excel sheet names
        // for C# and regex double backslash needed
        private const string SPECIAL_CHARS = "[/\\\\:=(){}\\[\\]*?\" <>|']";
        private struct PrintData
        {
            public string Title;
            public string Range;
        }

        private xl.Application _xlApp;
        private xl.Workbooks _workbooks;
        private xl.Workbook _src_wkb;
        private xl.Workbook _tmp_wkb;

        private string _data_sht_name;
        private string _tmpl_sht_name;
        public HunterGroupPrinter()
        {
            // every new created workbook has only one sheet
            _xlApp = new xl.Application { SheetsInNewWorkbook = 1 };
            _workbooks = _xlApp.Workbooks;

            // read from config
            _data_sht_name = Properties.Settings.Default.DataSheet;
            _tmpl_sht_name = Properties.Settings.Default.TemplateSheet;

            //_xlApp.ActivePrinter = printer;
            Console.WriteLine(_xlApp.ActivePrinter);
        }

        ~HunterGroupPrinter()
        {
            // close excel connection and all workbooks
            _tmp_wkb.Close(false);
            Marshal.FinalReleaseComObject(_tmp_wkb);
            _tmp_wkb = null;

            _src_wkb.Close(false);
            Marshal.FinalReleaseComObject(_src_wkb);
            _src_wkb = null;

            _workbooks.Close();
            Marshal.FinalReleaseComObject(_workbooks);
            _workbooks = null;

            _xlApp.Quit();
            Marshal.FinalReleaseComObject(_xlApp);
            _xlApp = null;
        }

        public void CreateCardsFromSource(string src_file)
        {
            // get necessary config data
            string grp_range_str = Properties.Settings.Default.HuntingGroups; // JAEGERGRUPPEN
            string sht_password = Properties.Settings.Default.SheetPassword;
            string number_clmn = Properties.Settings.Default.NumberColumn;
            string number_cell = Properties.Settings.Default.NumberCell; // GRUPPENNUMMER
            string leader_cell = Properties.Settings.Default.LeaderCell; // ANSTELLER

            _src_wkb = _workbooks.Open(src_file);
            xl.Worksheet src_sht = _src_wkb.Sheets[_data_sht_name];

            _tmp_wkb = CreateTmpWkbWithSeparator();

            // copy hunter sheet for every group
            xl.Range src_range = src_sht.Range[grp_range_str];
            foreach (xl.Range cell in src_range)
            {
                if (cell.Text == "")
                {
                    break;
                }

                xl.Worksheet last_sht = _tmp_wkb.Sheets[_tmp_wkb.Sheets.Count];
                _src_wkb.Sheets[_tmpl_sht_name].Copy(Before: last_sht);

                xl.Worksheet cp_sht = _tmp_wkb.ActiveSheet;

                // unlock sheet if locked
                // new locking is not necessary
                if (cp_sht.ProtectContents)
                {
                    cp_sht.Unprotect(sht_password);
                }

                cp_sht.Name = ModifyUnwantedNames(cell.Text);
                cp_sht.Range[number_cell].Value2 = src_sht.Range[number_clmn + cell.Row].Value2; // A + 3 = A3
                cp_sht.Range[leader_cell].Value2 = cell.Text;
            }
        }

        private xl.Workbook CreateTmpWkbWithSeparator()
        {
            // create temporary workbook with only one sheet
            xl.Workbook workbook = _workbooks.Add();

            workbook.Sheets[1].Name = "Trennblatt";
            workbook.Sheets[1].Cells[1, 1].Value = "Trennblatttext";
            workbook.Sheets[1].Cells[1, 1].Font.Name = "Arial";
            workbook.Sheets[1].Cells[1, 1].Font.Size = 72;
            workbook.Sheets[1].PageSetup.CenterHorizontally = true;
            workbook.Sheets[1].PageSetup.CenterVertically = true;
            workbook.Sheets[1].PageSetup.Orientation = xl.XlPageOrientation.xlLandscape;

            return workbook;
        }

        public void PrintCards(string group, bool separator)
        {
            // determine group title and data range for printing
            PrintData prt_data = DeterminePrintData(group);

            if (separator)
            {
                PrintSeparator(group);
            }

            // print sheets for specified group
            xl.Range prt_range = _src_wkb.Sheets[_data_sht_name].Range[prt_data.Range];
            foreach (xl.Range cell in prt_range)
            {
                // if cell.Value2 is null, set it to 0
                double copies = cell.Value2 ?? 0;
                if (copies > 0)
                {
                    int row = cell.Row - prt_range.Row + 1; // + 1 is for current row
                    _tmp_wkb.Sheets[row].Shapes.Item["TextField"].TextFrame.Characters.Text = prt_data.Title;
                    _tmp_wkb.Sheets[row].PrintOut(Copies: copies);
                }
            }
        }

        private PrintData DeterminePrintData(string group)
        {
            PrintData data = new PrintData();
            switch (group)
            {
                case "Ansteller":
                    data.Title = "Ansteller";
                    data.Range = Properties.Settings.Default.Leader;
                    break;

                case "Standschützen":
                    data.Title = "Standkarte";
                    data.Range = Properties.Settings.Default.Shooters;
                    break;

                case "Hundestände":
                    data.Title = "Hundestand";
                    data.Range = Properties.Settings.Default.Dogs;
                    break;

                case "Ersatzstände":
                    data.Title = "Ersatzstand";
                    data.Range = Properties.Settings.Default.Reserves;
                    break;
            }

            return data;
        }

        private void PrintSeparator(string separator_text)
        {
            _tmp_wkb.Sheets["trennblatt"].Cells[1, 1].Value = separator_text;
            _tmp_wkb.Sheets["trennblatt"].PrintOut();
        }

        private string ModifyUnwantedNames(string name)
        {
            // guid strings as well as regular names
            // may not longer than 30 characters
            name = name.Substring(0, Math.Min(name.Length, 30));
            name = Regex.Replace(name, SPECIAL_CHARS, "_");

            return name;
        }
    }
}
