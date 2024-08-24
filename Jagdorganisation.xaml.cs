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
        private HunterGroupPrinter _printer;
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
            open_dialog.Filter = "Jagdeinteilung (.xlsx)|*.xlsx";

            // show open file dialog box
            bool? result = open_dialog.ShowDialog();

            // save file name
            string source_file = "";
            if (result == true)
            {
                source_file = open_dialog.FileName;
            }
            else
            {
                return;
            }

            _printer = new HunterGroupPrinter();
            _printer.CreateCardsFromSource(source_file);

            // create array with checkboxes for printing groups
            CheckBox[] check_boxes = new CheckBox[] {
                LeaderCheckBox,
                ShootersCheckBox,
                DogsCheckBox,
                ReservesCheckBox
            };

            // define the print action
            // then execute the print action for each array element
            Action<CheckBox> print = new Action<CheckBox>(PrintGroups);
            Array.ForEach(check_boxes, print);
        }

        private void PrintGroups(CheckBox box)
        {
            // check, if the separator sheet should print
            bool? separator = SeparatorCheckBox.IsChecked;

            // print out if the box is checked
            if (box.IsChecked == true) { _printer.PrintCards(box.Content.ToString(), separator); }

            // timer before next print action

        }
    }
}
