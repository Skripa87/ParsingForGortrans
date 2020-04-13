using System;
using System.Collections.Generic;
using System.Windows;

namespace ParsingForGortrans
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private string _fileNameExcel;

        public MainWindow()
        {
            InitializeComponent();
            SelectedFile.Content = "";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_fileNameExcel))
                return;
            var reportManager = new ManagerReport(_fileNameExcel);
            reportManager.GetReport();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document";
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Text documents (.xlsx)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                _fileNameExcel = dlg.FileName;
                SelectedFile.Content = _fileNameExcel;
            }
        }
        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void CheckBox_OnChecked(object sender, RoutedEventArgs e)
        {
            
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Holidays_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}
