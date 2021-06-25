using Dyplom.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    /// Логика взаимодействия для ClassInfoExportPanel.xaml
    /// </summary>
    public partial class ClassInfoExportPanel : Window
    {
        private ModelContext db;
        public List<int> classesNumbers;
        public List<string> classesLetters;

        public ClassInfoExportPanel()
        {
            InitializeComponent();

            db = new ModelContext();
            LoadClassesGradation();
        }

        private void LoadClassesGradation()
        {
            db.Classes.Load();

            classesNumbers = db.Classes.GroupBy(g=>g.StudyYear).Select(c => c.Key).ToList();
            int firstNumberLetters = classesNumbers[0];
            classesLetters = db.Classes.Where(c => c.StudyYear == firstNumberLetters).Select(c => c.GradeSymbol).ToList();

            foreach (int item in classesNumbers)
            {
                numberCB.Items.Add(item);
            }

            foreach (string item in classesLetters)
            {
                letterCB.Items.Add(item);
            }

            numberCB.SelectedIndex = 0;
            letterCB.SelectedIndex = 0;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int NumberLetters = Convert.ToInt32(numberCB.SelectedItem);
            classesLetters = db.Classes.Where(c => c.StudyYear == NumberLetters).Select(c => c.GradeSymbol).ToList();

            letterCB.Items.Clear();
            foreach (string item in classesLetters)
            {
                letterCB.Items.Add(item);
            }
            letterCB.SelectedIndex = 0;
        }

        private void exportBTN_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application app = new Excel.Application();
            app.Workbooks.Add();
            Excel.Worksheet worksheet = app.ActiveSheet;

            CreateFileHead(worksheet);

            db.Students.Load();
            int classNumber = Convert.ToInt32(numberCB.SelectedItem);
            string classLetter = letterCB.SelectedItem.ToString();

            int classId = db.Classes.Where(w => w.StudyYear == classNumber && w.GradeSymbol == classLetter).Select(s => s.classid).Single();
            var students = db.Students.Where(w => w.classid == classId);

            int count = 0;
            int row = 4;
            foreach (Students st in students)
            {

                row++;
                count++;
                string name = st.studentName;
                bool isInvalid = st.isChildInvalit;
                bool isOPFR = st.isChildWithOPFR;
                bool isFoster = st.isChildInFosterCare;
                bool isHomeStudy = st.doesChildStudyAtHome;
                int childrenCount = st.numberOfChildInFamilyUnder18;
                bool isOnlyMother = st.incompleteFamilyOneMother;
                bool isOnlyFather = st.incompleteFamilyOneFather;
                string fatherEduc = st.fatherEducation;
                string motherEduc = st.motherEducation;
                string fatherWork = st.fatherStatus;
                string motherWork = st.motherStatus;

                worksheet.Cells[row, 1].Value2 = count;
                worksheet.Cells[row, 2].Value2 = name;
                if (isInvalid)
                    worksheet.Cells[row, 3].Value2 = "*";
                if (isOPFR)
                    worksheet.Cells[row, 4].Value2 = "*";
                if (isFoster)
                    worksheet.Cells[row, 5].Value2 = "*";
                if (isHomeStudy)
                    worksheet.Cells[row, 6].Value2 = "*";
                worksheet.Cells[row, 9].Value2 = childrenCount;
                if (childrenCount>=3)
                    worksheet.Cells[row, 10].Value2 = "*";
                if (isOnlyMother)
                    worksheet.Cells[row, 11].Value2 = "*";
                if (isOnlyFather)
                    worksheet.Cells[row, 12].Value2 = "*";
                switch (fatherEduc.ToLower())
                {
                    case "высшее":
                        worksheet.Cells[row, 14].Value2 = "п";
                        break;
                    case "среднее специальное":
                        worksheet.Cells[row, 15].Value2 = "п";
                        break;
                    case "общее среднее":
                        worksheet.Cells[row, 16].Value2 = "п";
                        break;
                    case "профтехобразование":
                        worksheet.Cells[row, 17].Value2 = "п";
                        break;
                    default:
                        break;
                }
                switch (motherEduc.ToLower())
                {
                    case "высшее":
                        worksheet.Cells[row, 14].Value2 += "м";
                        break;
                    case "среднее специальное":
                        worksheet.Cells[row, 15].Value2 += "м";
                        break;
                    case "общее среднее":
                        worksheet.Cells[row, 16].Value2 += "м";
                        break;
                    case "профтехобразование":
                        worksheet.Cells[row, 17].Value2 += "м";
                        break;
                    default:
                        break;
                }
                switch (fatherWork.ToLower())
                {
                    case "индивидальный предприниматель":
                    case "ип":
                        worksheet.Cells[row, 18].Value2 = "п";
                        break;
                    case "работает":
                        worksheet.Cells[row, 19].Value2 = "п";
                        break;
                    case "служит":
                    case "в армии":
                        worksheet.Cells[row, 20].Value2 = "п";
                        break;
                    case "пенсионер":
                        worksheet.Cells[row, 21].Value2 = "п";
                        break;
                    default:
                        break;
                }
                switch (motherWork.ToLower())
                {
                    case "индивидальный предприниматель":
                    case "ип":
                        worksheet.Cells[row, 18].Value2 += "м";
                        break;
                    case "работает":
                        worksheet.Cells[row, 19].Value2 += "м";
                        break;
                    case "служит":
                    case "в армии":
                        worksheet.Cells[row, 20].Value2 += "м";
                        break;
                    case "пенсионер":
                        worksheet.Cells[row, 21].Value2 += "м";
                        break;
                    default:
                        break;
                }

            }

            Excel.Range final = worksheet.Range[worksheet.Cells[row + 1, 1], worksheet.Cells[row + 1, 2]];
            final.Merge();
            final.Value2 = $"ИТОГО:{count}";
            final.Font.Bold = true;

            /*Пытался сделать формулу - почему-то вылетает COM ошибка, именно с этой формулой*/
            //Excel.Range sub = worksheet.Cells[row + 1, 3];
            //string query = string.Format(@"=COUNTIF(C5:C{0});""*""", row);
            //sub.Formula = query;
            //sub.Calculate();



            worksheet.Rows.AutoFit();
            worksheet.Columns.AutoFit();
            
            
            app.Visible = true;
        }

        private void CreateFileHead(Excel.Worksheet worksheet)
        {

            for (int i = 1; i <= 4; i++)
            {
                for (int j = 1; j <= 20; j++)
                {
                    worksheet.Cells[i, j].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }


            //ячейки жирного шрифта
            List<Excel.Range> boldRanges = new List<Excel.Range>();

            //ячейки вертикального текста
            List<Excel.Range> verticalTextRanges = new List<Excel.Range>();

            //заголовок
            Excel.Range head = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 10]];
            head.Merge();//объединение ячеек

            //учебный год
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
            head.Value2 = $"Социальный паспорт {numberCB.SelectedItem} \"{letterCB.SelectedItem}\" \t {firstYear}/{lastYear} учебный год, {halfYear}-е полугодие \t Классный руководитель: Пока что не знаю";

            //категории
            Excel.Range fio = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[4, 2]];
            fio.Merge();
            fio.Value2 = "Ф.И.О. учащегося";
            fio.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; //вертикаль - по центру
            fio.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //горизонталь - по центру
            fio.Borders.LineStyle = Excel.XlLineStyle.xlContinuous; //бордер - одиночный

            Excel.Range character = worksheet.Range[worksheet.Cells[2, 3], worksheet.Cells[3, 8]];
            character.Merge();
            character.Value2 = "характеристика учащихся";

            Excel.Range familyStatus = worksheet.Range[worksheet.Cells[2, 10], worksheet.Cells[2, 12]];
            familyStatus.Merge();
            familyStatus.Value2 = "Статус семьи";

            Excel.Range education = worksheet.Range[worksheet.Cells[2, 13], worksheet.Cells[2, 16]];
            education.Merge();
            education.Value2 = "Образование родителей";

            Excel.Range state = worksheet.Range[worksheet.Cells[2, 17], worksheet.Cells[2, 20]];
            state.Merge();
            state.Value2 = "Статус родителей";

            boldRanges.Add(fio);
            boldRanges.Add(head);
            boldRanges.Add(character);
            boldRanges.Add(familyStatus);
            boldRanges.Add(education);
            boldRanges.Add(state);

            worksheet.Cells[4, 3].Value2 = "ребенок-инвалид";
            worksheet.Cells[4, 4].Value2 = "ребенок с ОПФР";
            worksheet.Cells[4, 5].Value2 = "находится на опеке";
            worksheet.Cells[4, 6].Value2 = "воспитывается в приёмной семье";
            worksheet.Cells[4, 7].Value2 = "обучается на дому";
            worksheet.Cells[4, 8].Value2 = "признан в СОП;  на учете в ИПР, на ВК";
            worksheet.Cells[4, 9].Value2 = "Количество детей в семье до 18 лет";
            worksheet.Cells[4, 10].Value2 = "Многодетная семья";
            worksheet.Cells[4, 11].Value2 = "Неполная  семья (одна мать)";
            worksheet.Cells[4, 12].Value2 = "Неполная  семья (один отец)";
            worksheet.Cells[4, 13].Value2 = "высшее";
            worksheet.Cells[4, 14].Value2 = "среднее специальное";
            worksheet.Cells[4, 15].Value2 = "общее среднее";
            worksheet.Cells[4, 16].Value2 = "профтехобразование";
            worksheet.Cells[4, 17].Value2 = "ИП";
            worksheet.Cells[4, 18].Value2 = "Работает";
            worksheet.Cells[4, 19].Value2 = "Служит и т.п.";
            worksheet.Cells[4, 20].Value2 = "пенсионер";

            for (int i = 3; i <= 20; i++)
            {
                verticalTextRanges.Add(worksheet.Cells[4, i]);
                boldRanges.Add(worksheet.Cells[4, i]);
            }

            foreach (Excel.Range item in boldRanges)
            {
                item.Font.Bold = true;
            }
            foreach (Excel.Range item in verticalTextRanges)
            {
                item.Orientation = 90;
            }
        }
    }
}
