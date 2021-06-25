using System;
using System.Collections.Generic;
using System.Data.Entity;
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

using Excel = Microsoft.Office.Interop.Excel;

namespace Dyplom
{
    /// <summary>
    /// Логика взаимодействия для OneParentsExportPanel.xaml
    /// </summary>
    public partial class OneParentsExportPanel : Window
    {
        private Excel.Application app;
        private Excel.Worksheet worksheet;
        private List<Models.Students> students;
        private int currentRow;

        public OneParentsExportPanel()
        {
            InitializeComponent();
        }

        private void oneFathersExportBTN_Click(object sender, RoutedEventArgs e)
        {
            PrepareExcel();
            CreateHeader("один отец");

            CollectOneFatherData();

            CreateBody();

            CreateFooter();
        }

        private void CollectOneMotherData()
        {
            using (Models.ModelContext db = new Models.ModelContext())
            {
                db.Students.Load();

                students =  db.Students.Where(w => w.incompleteFamilyOneMother == true).ToList();
            }
        }

        private void CollectOneFatherData()
        {
            using (Models.ModelContext db = new Models.ModelContext())
            {
                db.Students.Load();

                students =  db.Students.Where(w => w.incompleteFamilyOneFather == true).ToList();
            }
        }

        private void oneMothersExportBTN_Click(object sender, RoutedEventArgs e)
        {
            PrepareExcel();
            CreateHeader("одна мать");

            CollectOneMotherData();

            CreateBody();

            CreateFooter();
        }

        private void CreateFooter()
        {
            currentRow+=2;
            worksheet.Range[worksheet.Cells[currentRow, 1], worksheet.Cells[currentRow, 8]].Merge();
            worksheet.Range[worksheet.Cells[currentRow, 1], worksheet.Cells[currentRow, 8]].Value2 = "Директор";
            currentRow++;
            worksheet.Range[worksheet.Cells[currentRow, 1], worksheet.Cells[currentRow, 8]].Merge();
            worksheet.Range[worksheet.Cells[currentRow, 1], worksheet.Cells[currentRow, 8]].Value2 = "государственного учреждения образования";
            currentRow++;
            worksheet.Range[worksheet.Cells[currentRow, 1], worksheet.Cells[currentRow, 6]].Merge();
            worksheet.Range[worksheet.Cells[currentRow, 1], worksheet.Cells[currentRow, 6]].Value2 = @"""Средняя школа №21 г.Могилева""";
            worksheet.Range[worksheet.Cells[currentRow, 7], worksheet.Cells[currentRow, 8]].Merge();
            worksheet.Range[worksheet.Cells[currentRow, 7], worksheet.Cells[currentRow, 8]].Value2 = "Г.А. Викторова";

            worksheet.Columns.AutoFit();

            app.Visible = true;
        }

        private void CreateBody()
        {
            int counter = 0;
            currentRow = 7;

            foreach (Models.Students st in students)
            {
                counter++;
                currentRow++;

                worksheet.Cells[currentRow, 1].Value2 = counter;
                worksheet.Cells[currentRow, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[currentRow, 2].Value2 = counter;
                worksheet.Cells[currentRow, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[currentRow, 3].Value2 = st.motherName;
                worksheet.Cells[currentRow, 4].Value2 = st.studentName;

                using (Models.ModelContext db = new Models.ModelContext())
                {
                    db.Classes.Load();
                    int classNumber = db.Classes.Where(w => w.classid == st.classid).Select(s => s.StudyYear).FirstOrDefault();
                    string classLetter = db.Classes.Where(w => w.classid == st.classid).Select(s => s.GradeSymbol).FirstOrDefault();
                    worksheet.Cells[currentRow, 5].Value2 = classNumber.ToString()+classLetter;
                    worksheet.Cells[currentRow, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }

                worksheet.Cells[currentRow, 6].Value2 = st.birthdate.ToString("d");
                worksheet.Cells[currentRow, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[currentRow, 7].Value2 = st.numberOfChildInFamilyUnder18;
                worksheet.Cells[currentRow, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[currentRow, 8].Value2 = st.homeAdressReg;

                for (int i = 8; i <= currentRow; i++)
                {
                    for (int j = 1; j <= 8; j++)
                    {
                        worksheet.Cells[i, j].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                }
            }
        }

        private void PrepareExcel()
        {
            app = new Excel.Application();
            app.Workbooks.Add();
            worksheet = app.ActiveSheet;
        }

        private void CreateHeader(string mode)
        {
            List<Excel.Range> centerTextRanges = new List<Excel.Range>();
            List<Excel.Range> boldTextRanges = new List<Excel.Range>();

            //учебный год
            int nowDay = DateTime.Now.Day;
            int nowMonth = DateTime.Now.Month;
            int firstYear, lastYear, halfYear;
            //если месяц меньше или равен августу - то это конечная дата учебного года
            if (nowMonth <= 8)
            {
                lastYear = DateTime.Now.Year;
                firstYear = lastYear - 1;
                halfYear = 2;
            }
            else
            {
                firstYear = DateTime.Now.Year;
                lastYear = firstYear + 1;
                halfYear = 1;
            }

            Excel.Range spisok = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 8]];
            spisok.Merge();
            spisok.Value2 = "С П И С О К";
            centerTextRanges.Add(spisok);

            Excel.Range sub1 = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[2, 8]];
            sub1.Merge();
            sub1.Value2 = "учащихся государственного учреждения образования «Средняя школа №21 г. Могилева»";
            centerTextRanges.Add(sub1);

            Excel.Range sub2 = worksheet.Range[worksheet.Cells[3, 1], worksheet.Cells[3, 8]];
            sub2.Merge();
            sub2.Value2 = $"из семей, где воспитанием занимается {mode}";
            centerTextRanges.Add(sub2);

            Excel.Range sub3 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4, 8]];
            sub3.Merge();
            sub3.Value2 = $" на {nowDay}.{nowMonth} {firstYear}/{lastYear} учебного года";
            centerTextRanges.Add(sub3);

            Excel.Range PPfamiy = worksheet.Cells[6, 1];
            PPfamiy.Value2 = "№ п\\п семьи";
            centerTextRanges.Add(PPfamiy);
            boldTextRanges.Add(PPfamiy);

            Excel.Range PPNL = worksheet.Cells[6, 2];
            PPNL.Value2 = "№ п\\п н/л";
            centerTextRanges.Add(PPNL);
            boldTextRanges.Add(PPNL);

            Excel.Range fioParent = worksheet.Cells[6, 3];
            fioParent.Value2 = "Ф.И.О. родителя";
            centerTextRanges.Add(fioParent);
            boldTextRanges.Add(fioParent);

            Excel.Range fioLearn = worksheet.Cells[6, 4];
            fioLearn.Value2 = "Ф.И.О. учащегося";
            centerTextRanges.Add(fioLearn);
            boldTextRanges.Add(fioLearn);

            Excel.Range learnClass = worksheet.Cells[6, 5];
            learnClass.Value2 = "Класс";
            centerTextRanges.Add(learnClass);
            boldTextRanges.Add(learnClass);

            Excel.Range dateOfBirth = worksheet.Cells[6, 6];
            dateOfBirth.Value2 = "Дата рождения";
            centerTextRanges.Add(dateOfBirth);
            boldTextRanges.Add(dateOfBirth);

            Excel.Range childrenCount = worksheet.Cells[6, 7];
            childrenCount.Value2 = "Количество детей";
            centerTextRanges.Add(childrenCount);
            boldTextRanges.Add(childrenCount);

            Excel.Range homeAddress = worksheet.Cells[6, 8];
            homeAddress.Value2 = "Домашний адрес";
            centerTextRanges.Add(homeAddress);
            boldTextRanges.Add(homeAddress);

            Excel.Range n1 = worksheet.Cells[7, 1];
            n1.Value2 = "1";
            centerTextRanges.Add(n1);
            Excel.Range n2 = worksheet.Cells[7, 2];
            n2.Value2 = "2";
            centerTextRanges.Add(n2);
            Excel.Range n3 = worksheet.Cells[7, 3];
            n3.Value2 = "3";
            centerTextRanges.Add(n3);
            Excel.Range n4 = worksheet.Cells[7, 4];
            n4.Value2 = "4";
            centerTextRanges.Add(n4);
            Excel.Range n5 = worksheet.Cells[7, 5];
            n5.Value2 = "5";
            centerTextRanges.Add(n5);
            Excel.Range n6 = worksheet.Cells[7, 6];
            n6.Value2 = "6";
            centerTextRanges.Add(n6);
            Excel.Range n7 = worksheet.Cells[7, 7];
            n7.Value2 = "7";
            centerTextRanges.Add(n7);
            Excel.Range n8 = worksheet.Cells[7, 8];
            n8.Value2 = "8";
            centerTextRanges.Add(n8);



            foreach (Excel.Range vs in centerTextRanges)
            {
                vs.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                vs.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            foreach (Excel.Range vs in boldTextRanges)
            {
                vs.Font.Bold = true;
            }

            for (int i = 6; i <= 7; i++)
            {
                for (int j = 1; j <= 8; j++)
                {
                    worksheet.Cells[i,j].Borders.LineStyle = Excel.XlLineStyle.xlContinuous; 
                }
            }
        } 
    }
}
