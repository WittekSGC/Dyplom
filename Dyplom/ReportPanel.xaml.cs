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

namespace Dyplom
{
    /// <summary>
    /// Логика взаимодействия для ReportPanel.xaml
    /// </summary>
    public partial class ReportPanel : Window
    {
        public ReportPanel()
        {
            InitializeComponent();

            //сюда надо вставить ссылку на опубликованный отчет 
            reportBrowser.Source = new Uri("https://app.powerbi.com/view?r=eyJrIjoiMzNiMTZmYjEtOTU1Mi00NzdmLTk5MGItYzUyYjc2ZDA2NzFkIiwidCI6IjUyYjkyMzYyLWNiMmYtNDdlYy1iOTBjLTkxNWQ1ZjBmMzcxNyIsImMiOjl9");
        }
    }
}
