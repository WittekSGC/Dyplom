using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Dyplom.Models;
using System.Data.Entity;
using Microsoft.Win32;
using ExcelDataReader;
using System.Data;
using Microsoft.Office.Interop.Excel;

namespace Dyplom
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class MainPanel : System.Windows.Window
    {
        ModelContext db;
        private string FileName = string.Empty;
        private DataTableCollection tableCollection = null;
        public MainPanel()
        {
            InitializeComponent();

            db = new ModelContext();
            studentInfoGrid.Items.Clear();
            db.Students.Load();
            studentInfoGrid.ItemsSource = db.Students.Local.ToBindingList();

            db.Classes.Load();

            this.Closing += MainWindow_Closing;

        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            db.Dispose();
        }
        private void restartBtn_Click(object sender, RoutedEventArgs e)
        {
            db.Students.Load();
            studentInfoGrid.ItemsSource = db.Students.Local.ToBindingList();
        }
        private void updateBtn_Click(object sender, RoutedEventArgs e)
        {
            db.SaveChanges();
            MessageBox.Show("Обновление базы данных прошло успешно!", "Уведомление от системы");
            db.Students.Load();
        }

        private void deleteBtn_Click(object sender, RoutedEventArgs e)
        {
            if (studentInfoGrid.SelectedItems.Count > 0)
            {
                for (int i = 0; i < studentInfoGrid.SelectedItems.Count; i++)
                {
                    Students student = studentInfoGrid.SelectedItems[i] as Students;
                    if (student != null)
                    {
                        db.Students.Remove(student);
                        MessageBox.Show("Удаление записи из базы данных прошло успешно!", "Уведомление от системы");
                    }
                }
            }
            db.SaveChanges();
        }
        private void findBtn_Click(object sender, RoutedEventArgs e)
        {
            ModelContext db;

            db = new ModelContext();
            switch (findComboBox.SelectedIndex)
            {
                case (0):
                    {
                        db.Students.Where(s => s.studentName.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (1):
                    {
                        //db.Students.Where(s => s.Birthdate.Contains(findTextBox.Text)).Load();
                        //MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (2):
                    {
                        db.Students.Where(s => s.homeAdressReg.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (3):
                    {
                        db.Students.Where(s => s.homeAdressRel.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (4):
                    {
                        db.Students.Where(s => s.studentTel.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (5):
                    {
                        db.Students.Where(s => s.motherName.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (6):
                    {
                        db.Students.Where(s => s.motherPlaceOfWork.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (7):
                    {
                        db.Students.Where(s => s.motherWorkPhone.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (8):
                    {
                        db.Students.Where(s => s.motherMobPhone.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (9):
                    {
                        db.Students.Where(s => s.fatherName.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (10):
                    {
                        db.Students.Where(s => s.fatherPlaceOfWork.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (11):
                    {
                        db.Students.Where(s => s.fatherWorkPhone.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (12):
                    {
                        db.Students.Where(s => s.fatherMobPhone.Contains(findTextBox.Text)).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
            }
            switch (childInvalidComboBox.SelectedIndex)
            {
                case (0):
                    {
                        db.Students.Where(s => s.isChildInvalit == true).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        //db.StudentsInfo.Local.Where(s => s.isChildInvalit == true);
                        break;
                    }
                case (1):
                    {
                        db.Students.Where(s => s.isChildInvalit == false).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (2):
                    {
                        db.Students.Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
            }
            switch (childOPFRComboBox.SelectedIndex)
            {
                case (0):
                    {
                        db.Students.Where(s => s.isChildWithOPFR == true).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (1):
                    {
                        db.Students.Where(s => s.isChildWithOPFR == false).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (2):
                    {
                        db.Students.Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
            }
            switch (childincustodyComboBox.SelectedIndex)
            {
                case (0):
                    {
                        db.Students.Where(s => s.childInCustody == true).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (1):
                    {
                        db.Students.Where(s => s.childInCustody == false).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (2):
                    {
                        db.Students.Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
            }
            switch (childinFosterCareComboBox.SelectedIndex)
            {
                case (0):
                    {
                        db.Students.Where(s => s.isChildInFosterCare == true).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (1):
                    {
                        db.Students.Where(s => s.isChildInFosterCare == false).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (2):
                    {
                        db.Students.Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
            }
            switch (childRegisteredComboBox.SelectedIndex)
            {
                case (0):
                    {
                        db.Students.Where(s => s.isChildRegistered == true).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (1):
                    {
                        db.Students.Where(s => s.isChildRegistered == false).Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
                case (2):
                    {
                        db.Students.Load();
                        MessageBox.Show("Поиск записи в базе данных прошло успешно!", "Уведомление от системы");
                        break;
                    }
            }
            studentInfoGrid.ItemsSource = db.Students.Local.ToBindingList();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }



        private void importExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            //OpenExcelImportMenu();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL файлы(*.xltm;*.xlsx)|*.xltm;*.xlsx|Все файлы (*.*)|*.* ";
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
        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = false
                }
            });

            tableCollection = db.Tables;


            OpenData();
        }
        private void exportExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenExcelExportMenu();
        }

        private void OpenExcelExportMenu()
        {
            ExcelImportPanel excelImport = new ExcelImportPanel();
            excelImport.Owner = this;
            excelImport.Show();
        }

        private void classesShowBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OpenData()
        {
            System.Data.DataTable table = tableCollection[0];
            if (table == null) return;

            foreach (DataRow row in table.Rows)
            {
                Students s = new Students();
                try
                {
                    s.studentName = row.Field<string>(table.Columns[0].ColumnName);
                    s.birthdate = row.Field<DateTime>(table.Columns[1].ColumnName);
                    s.homeAdressReg = row.Field<string>(table.Columns[2].ColumnName);
                    s.homeAdressRel = row.Field<string>(table.Columns[3].ColumnName);
                    s.studentTel = row.Field<object>(table.Columns[4].ColumnName).ToString();
                    s.motherName = row.Field<string>(table.Columns[5].ColumnName);
                    s.motherPlaceOfWork = row.Field<string>(table.Columns[6].ColumnName);
                    s.motherWorkPhone = row.Field<string>(table.Columns[7].ColumnName);
                    s.motherMobPhone = row.Field<object>(table.Columns[8].ColumnName).ToString();
                    s.fatherName = row.Field<string>(table.Columns[9].ColumnName);
                    s.fatherPlaceOfWork = row.Field<string>(table.Columns[10].ColumnName);
                    s.fatherWorkPhone = row.Field<object>(table.Columns[11].ColumnName).ToString();
                    s.fatherMobPhone = row.Field<object>(table.Columns[12].ColumnName).ToString();
                    s.isChildInvalit = Convert.ToBoolean(row.Field<double>(table.Columns[13].ColumnName));
                    s.isChildWithOPFR = Convert.ToBoolean(row.Field<double>(table.Columns[14].ColumnName));
                    s.childInCustody = Convert.ToBoolean(row.Field<double>(table.Columns[15].ColumnName));
                    s.isChildInFosterCare = Convert.ToBoolean(row.Field<double>(table.Columns[16].ColumnName));
                    s.doesChildStudyAtHome = Convert.ToBoolean(row.Field<double>(table.Columns[17].ColumnName));
                    s.isChildRegistered = Convert.ToBoolean(row.Field<double>(table.Columns[18].ColumnName));
                    s.numberOfChildInFamilyUnder18 = Convert.ToInt32(row.Field<double>(table.Columns[19].ColumnName));
                    s.incompleteFamilyOneMother = Convert.ToBoolean(row.Field<double>(table.Columns[20].ColumnName));
                    s.incompleteFamilyOneFather = Convert.ToBoolean(row.Field<double>(table.Columns[21].ColumnName));
                    s.aSingleMother = Convert.ToBoolean(row.Field<double>(table.Columns[22].ColumnName));
                    s.motherEducation = row.Field<string>(table.Columns[23].ColumnName);
                    s.fatherEducation = row.Field<string>(table.Columns[24].ColumnName);
                    s.motherStatus = row.Field<string>(table.Columns[25].ColumnName);
                    s.fatherStatus = row.Field<string>(table.Columns[26].ColumnName);
                    s.classid = Convert.ToInt32(row.Field<double>(table.Columns[27].ColumnName));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    continue;
                }
                db.Students.Add(s);
            }
            
            studentInfoGrid.UpdateLayout();
        }

        private void exportExcelBtn_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            app.Workbooks.Add();
            Worksheet worksheet = app.ActiveSheet;
            for (int i = 0; i < studentInfoGrid.Columns.Count() - 1; i++)
            {
                worksheet.Cells[1, i + 1] = studentInfoGrid.Columns[i].Header.ToString();
            }
            for (int i = 0; i < db.Students.Count() ; i++)
            {
                for (int j = 0; j < studentInfoGrid.Columns.Count() - 1; j++)
                {
                    var ci = new DataGridCellInfo(studentInfoGrid.Items[i], studentInfoGrid.Columns[j]);
                    var content = (ci.Column.GetCellContent(ci.Item) as TextBlock).Text;
                    worksheet.Cells[i + 2, j + 1] = content;
                }
            }
            worksheet.Columns.AutoFit();


            app.Visible = true;
        }

        private void showReportBtn_Click(object sender, RoutedEventArgs e)
        {
            new ReportPanel().Show();
        }

        private void showExportClassInfoBtn_Click(object sender, RoutedEventArgs e)
        {
            new ClassInfoExportPanel().Show();
        }

        private void showExportOneParentInfoBtn_Click(object sender, RoutedEventArgs e)
        {
            new OneParentsExportPanel().Show();
        }

        /*private void OpenExcelImportMenu()
        {
            ExcelExportPanel excelExport = new ExcelExportPanel();
            excelExport.Owner = this;
            excelExport.Show();
        }*/
    }
}
