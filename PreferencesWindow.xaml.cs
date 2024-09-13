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
using System.Windows.Shapes;
using System.ComponentModel;

namespace Jagdorganisation
{
    /// <summary>
    /// Interaktionslogik für Window1.xaml
    /// </summary>
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            AllowEdit(true);
        }

        private void AllowEdit(bool request)
        {
            foreach (TextBox tb in SheetGrid.Children.OfType<TextBox>())
            {
                tb.IsEnabled = request;
            }

            foreach (TextBox tb in DivisionGrid.Children.OfType<TextBox>())
            {
                tb.IsEnabled = request;
            }

            foreach (TextBox tb in TemplateGrid.Children.OfType<TextBox>())
            {
                tb.IsEnabled = request;
            }

            SaveButton.IsEnabled = request;
            AbortButton.IsEnabled = request;
        }

        private void AbortButton_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Reload();
            AllowEdit(false);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();
            AllowEdit(false);
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Reload();
            Close();
        }
    }
}
