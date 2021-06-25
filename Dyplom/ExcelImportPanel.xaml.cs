using System;
using System.Windows;
using Microsoft.Win32;
using ExcelDataReader;
using System.Data;
using System.IO;

namespace Dyplom
{
    /// <summary>
    /// Логика взаимодействия для ExcelImportPanel.xaml
    /// </summary>
    public partial class ExcelImportPanel : Window
    {
        private string FileName = string.Empty;
        private DataTableCollection tableCollection = null;
        public ExcelImportPanel()
        {
            InitializeComponent();
        }

        private void OpenMenu()
        {
            Hide();
            MainPanel w1 = new MainPanel();
            w1.Owner = this;
            w1.Show();
        }

        private void AddListOfStudentBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL файлы(*.xltm;*.xlsx)|*.xltm;*.xlsx" + "|Все файлы (*.*)|*.* ";
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = false;
            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    FileName = openFileDialog.FileName;

                    OpenExcelFile(FileName);
                }

                else
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }



        }

        private void MainPanelBackBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenMenu();
        }

        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            tableCollection = db.Tables;

            ExcelImportListComboBox.Items.Clear();

            foreach(DataTable table in tableCollection)
            {
                ExcelImportListComboBox.Items.Add(table.TableName); 
            }
            ExcelImportListComboBox.SelectedIndex = 0;

        }

        private void ExcelImportListComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            DataTable dataTable = tableCollection[Convert.ToString(ExcelImportListComboBox.SelectedItem)];

           
        }
    }
}
