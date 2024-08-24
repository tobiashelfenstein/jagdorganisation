using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Collections;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;

namespace Jagdorganisation
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            // see https://automationtesting.in/row-count-excel-using-c/

            // select only excel files
            var open_dialog = new Microsoft.Win32.OpenFileDialog();
            open_dialog.Title = "Jagdeinteilung laden";
            //open_dialog.DefaultExt = ".xlsx";
            open_dialog.Filter = "Jagdeinteilung (.xlsx)|*.xlsx";

            // show open file dialog box
            bool? result = open_dialog.ShowDialog();

            // save file name
            string source_file = "";
            if (result == true)
            {
                source_file = open_dialog.FileName;
                // datei überprüfen
                Console.WriteLine(source_file);
            }
            else
            {
                return;
            }

            HunterGroupPrinter printer = new HunterGroupPrinter();
            printer.CreateCardsFromSource(source_file);
            printer.PrintCards("ERSATZ");

            // start excel connection
            /*xl.Application xlApp = new xl.Application();
            xlApp.SheetsInNewWorkbook = 1;
            xl.Workbooks workbooks = xlApp.Workbooks;
            xl.Workbook workbook = workbooks.Open(source_file);

            xl.Worksheet source_sheet = workbook.Sheets["einteilung"];

            // calculate some numbers
            double num_groups = source_sheet.Range["C6:D35"].Rows.Count; // JAEGERGRUPPEN
            double num_leaders = xlApp.WorksheetFunction.CountA(source_sheet.Range["E6:E35"]);
            double num_shooters = xlApp.WorksheetFunction.CountA(source_sheet.Range["G6:G35"]);
            double num_dogs = xlApp.WorksheetFunction.CountA(source_sheet.Range["I6:I35"]);
            double num_reserves = xlApp.WorksheetFunction.CountA(source_sheet.Range["J6:J35"]);

            // create temporary workbook with only one sheet
            xl.Workbook temp_workbook = workbooks.Add();

            // set up first sheet as separator
            //temp_workbook.Worksheets.Add(After: temp_workbook.Sheets[temp_workbook.Sheets.Count]);
            temp_workbook.ActiveSheet.Name = "Trennblatt";
            temp_workbook.ActiveSheet.Cells[1, 1].Value = "Trennblatttext";
            temp_workbook.ActiveSheet.Cells[1, 1].Font.Name = "Arial";
            temp_workbook.ActiveSheet.Cells[1, 1].Font.Size = 72;
            temp_workbook.ActiveSheet.PageSetup.CenterHorizontally = true;
            temp_workbook.ActiveSheet.PageSetup.CenterVertically = true;
            temp_workbook.ActiveSheet.PageSetup.Orientation = xl.XlPageOrientation.xlLandscape;

            xl.Range source_range = source_sheet.Range["C6:D35"]; // JAEGERGRUPPEN
            foreach (xl.Range cell in source_range)
            {
                if (cell.Text != "")
                {
                    workbook.Sheets["standkarte"].Copy(Before: temp_workbook.Sheets[temp_workbook.Sheets.Count]);
                    temp_workbook.ActiveSheet.Name = cell.Text;
                    temp_workbook.ActiveSheet.Range["B15"].Value2 = source_sheet.Range["A" + cell.Row].Value2; // NUMMER
                    temp_workbook.ActiveSheet.Range["C15"].Value2 = cell.Text; // ANSTELLER
                }
            }

            // print cards out
            Console.WriteLine("Karten werden gedruckt");


            //temp_workbook.SaveAs("\\\\mmedia\\users\\tobias\\tmp\\testfile.xlsx");

            //Console.WriteLine(sheet.Cells[2, 1].Value.ToString());
            //Console.WriteLine(sheet.Range["C6:D35"].Rows.Count);

            // close excel connection
            temp_workbook.Close(false);
            Marshal.FinalReleaseComObject(temp_workbook);
            temp_workbook = null;

            workbook.Close(false, source_file, null);
            Marshal.FinalReleaseComObject(workbook);
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
*/
        }

        private void printCardsOut(string group)
        {

        }
    }
}
