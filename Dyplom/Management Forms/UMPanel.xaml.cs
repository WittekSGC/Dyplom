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
using System.Data.Entity;
using Dyplom.Models;

namespace Dyplom
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>
    public partial class UMPanel : Window
    {
        ModelContext db;
        public UMPanel()
        {
            InitializeComponent();

            db = new ModelContext();
            leadTeacherInfoGrid.Items.Clear();
            db.LeadTeachers.Load();
            leadTeacherInfoGrid.ItemsSource = db.LeadTeachers.Local.ToBindingList();

            managementsInfoGrid.Items.Clear();
            db.Management.Load();
            managementsInfoGrid.ItemsSource = db.Management.Local.ToBindingList();

            this.Closing += MainWindow_Closing;


        }

        private void OpenManagementMenu()
        {
            Hide();
            ManagementMenuWindow w1 = new ManagementMenuWindow();
            w1.Owner = this;
            w1.Show();
        }
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            db.Dispose();
        }
        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void updateManagementBtn_Click(object sender, RoutedEventArgs e)
        {
            db.SaveChanges();
            MessageBox.Show("Обновление базы данных прошло успешно!", "Уведомление от системы");
            db.Management.Load();
        }

        private void deleteManagementBtn_Click(object sender, RoutedEventArgs e)
        {
            {
                /*if (managementsInfoGrid.SelectedItems.Count > 0)
                {
                    for (int i = 0; i < managementsInfoGrid.SelectedItems.Count; i++)
                    {
                        Management management = managementsInfoGrid.SelectedItems[i] as Management;
                        if (management != null)
                        {
                            db.LeadTeachers.Remove(management);
                            MessageBox.Show("Удаление записи из базы данных прошло успешно!", "Уведомление от системы");
                        }
                    }
                }
                db.SaveChanges();
            }*/
            }
        }
        private void updateBtnLeadTeachers_Click(object sender, RoutedEventArgs e)
        {
            db.SaveChanges();
            MessageBox.Show("Обновление базы данных прошло успешно!", "Уведомление от системы");
            db.LeadTeachers.Load();
        }

        private void deleteBtnLeadTeachers_Click(object sender, RoutedEventArgs e)
        {
            if (leadTeacherInfoGrid.SelectedItems.Count > 0)
            {
                for (int i = 0; i < leadTeacherInfoGrid.SelectedItems.Count; i++)
                {
                    LeadTeachers leadTeachers = leadTeacherInfoGrid.SelectedItems[i] as LeadTeachers;
                    if (leadTeachers != null)
                    {
                        db.LeadTeachers.Remove(leadTeachers);
                        MessageBox.Show("Удаление записи из базы данных прошло успешно!", "Уведомление от системы");
                    }
                }
            }
            db.SaveChanges();
        }

        private void ManagementPanelBackBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenManagementMenu();
        }
    }
}
