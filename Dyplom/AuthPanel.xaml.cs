using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Dyplom.Models;
using System.Data.Entity;

namespace Dyplom
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ModelContext db;
        public MainWindow()
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

        private void OpenManagementMenuPanel()
        {
            Hide();
            ManagementMenuWindow managementMenuWindow = new ManagementMenuWindow();
            managementMenuWindow.Owner = this;
            managementMenuWindow.Show();
        }

        private void Button_Login_Click(object sender, RoutedEventArgs e)
        {

            //TEST
            //OpenMenu();
            //return;

            db = new ModelContext();

            if (!String.IsNullOrWhiteSpace(loginTextBox.Text) && !String.IsNullOrWhiteSpace(passTextBox.Password))
            {
                if ((bool)cbLeadTeachers.IsChecked)
                {
                    LeadTeachers leadTeachers = db.LeadTeachers.Where(p => p.teacherlogin == loginTextBox.Text).SingleOrDefault();
                    if (leadTeachers != null)
                        if (leadTeachers.teacherpassword == passTextBox.Password.ToString())
                        {
                            OpenMenu();
                        }
                        else
                        {
                            MessageBox.Show("Неправильный пароль", "Ошибка авторизации");
                        }
                    else
                    {
                        MessageBox.Show("Неправильный логин", "Ошибка авторизации");
                    }
                    db.Dispose();
                }
                else
                {
                    Management management = db.Management.Where(p => p.ManagementLogin == loginTextBox.Text).SingleOrDefault();
                    if (management != null)
                        if (management.ManagementPassword == passTextBox.Password.ToString())
                        {
                            OpenManagementMenuPanel();
                        }
                        else
                        {
                            MessageBox.Show("Неправильный пароль", "Ошибка авторизации");
                        }
                    else
                    {
                        MessageBox.Show("Неправильный логин", "Ошибка авторизации");
                    }
                    db.Dispose();
                }
        }


            /*private void TbLogin_KeyDown(object sender, KeyEventArgs e)
            {
                if (Keyboard.IsKeyDown(Key.Enter))
                {
                    passTextBox.Focus();
                }
            }

            private void TbPassword_KeyDown(object sender, KeyEventArgs e)
            {
                if (Keyboard.IsKeyDown(Key.Enter) && enterButton.IsEnabled)
                {
                    Button_Login_Click(enterButton, new RoutedEventArgs());
                }
            }*/
        }

        private void cbLeadTeachers_Checked(object sender, RoutedEventArgs e)
        {

        }
    }

}