using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using BDClassLib;
using PriemLib;

namespace Priem
{
    public partial class Columns : Form
    {
        private FormFilter own;
        private DBPriem _bdc;
        private string _fac;
        private SortedList<string,string> _columnList;

        //конструктор
        public Columns(FormFilter owner, string facId)
        {
            InitializeComponent();
            this.CenterToParent();
            own = owner;
            _fac = facId;
            _bdc = MainClass.Bdc;
            _columnList = FillColumns();

            foreach (DataGridViewColumn col in own.Dgv.Columns)
                if (col.Visible)
                    try
                    {
                        lbYes.Items.Add(_columnList[col.Name]);
                    }
                    catch { }


            foreach(string li in _columnList.Values)
            {
                try
                {
                    if (!lbYes.Items.Contains(li))
                        lbNo.Items.Add(li);
                }
                catch
                {
                }
            }
        }

        private SortedList<string, string> FillColumns()
        {
            SortedList<string, string> list = new SortedList<string, string>();
            list.Add("Рег_номер", "Рег. номер");
            list.Add("Ид_номер", "Идент. номер");
            list.Add("Фамилия", "Фамилия");
            list.Add("Имя", "Имя");
            list.Add("Отчество", "Отчество");
            list.Add("ФИО", "ФИО");
            list.Add("Дата_рождения", "Дата рождения");
            list.Add("Место_рождения", "Место рождения");
            list.Add("Тип_паспорта", "Тип паспорта");
            list.Add("Серия_паспорта", "Серия паспорта");
            list.Add("Номер_паспорта", "Номер паспорта");
            list.Add("Кем_выдан_паспорт", "Кем выдан паспорт"); 
            list.Add("Дата_выдачи_паспорта", "Дата выдачи паспорта");
            list.Add("Код_подразделения_паспорта", "Код подразделения (паспорт)");
            list.Add("Личный_код_паспорт", "Личный код (паспорт)");
            list.Add("Пол_мужской", "Пол");
            list.Add("Телефон", "Телефон");
            list.Add("Email", "E-mail");
            list.Add("Адрес_регистрации", "Адрес регистрации");
            list.Add("Адрес_проживания", "Адрес проживания");
            list.Add("Страна", "Страна");
            list.Add("Гражданство", "Гражданство");
            list.Add("Регион", "Регион");
            list.Add("Предоставлять_общежитие_поступление", "Предоставлять общежитие на время поступления");
            list.Add("Выдано_направление_на_поселение", "Выдано направление на поселение");
            list.Add("Факультет_выдавший_направление", "Факультет, выдавший направление");           
            list.Add("Ин_язык", "Ин. язык испытания");           
            list.Add("Страна_получ_пред_образ", "Страна получения пред.образования");            
            list.Add("Название_уч_заведения", "Название уч. заведения"); 
            list.Add("Предоставлять_общежитие_обучение", "Предоставлять общежитие на время обучения");        
            list.Add("Протокол_о_допуске", "Протокол о допуске");
            list.Add("Представление", "Представление к зачислению");
            list.Add("Коэффициент_полупрохода", "Рейтинговый коэффициент");
            list.Add("Сумма_баллов", "Сумма баллов");
            list.Add("Номер_зачетки", "Номер зачетной книжки");
            list.Add("Целевик_тип", "Тип целевика");
            list.Add("Средний_балл_сессии", "Средний балл сессии");
            list.Add("Статус_студента", "Статус студента");
            list.Add("Приоритет", "Приоритет");
            list.Add("Доп_контакты", "Доп. контакты");
            list.Add("Данные_о_родителях", "Данные о родителях");
            
            list.Add("Слушатель", "Слушатель");
            list.Add("Оплатил", "Оплатил");
            list.Add("Забрал_док", "Забрал док.");
            list.Add("Данные_проверены", "Данные проверены");
            list.Add("Дата_возврата_док", "Дата возврата док."); 
            list.Add("Дата_подачи_док", "Дата подачи док.");
            list.Add("Поданы_подлинник_атт", "Подан подлинник документа об образовании");            
            list.Add("Подал_подлинники", "Подал подлинники");
                        
            list.Add("Факультет", "Факультет");
            list.Add("Направление", "Направление");
            list.Add("Образ_программа", "Образовательная программа");
            list.Add("Код_направления", "Код направления");
            list.Add("Форма_обучения", "Форма обучения");
            list.Add("Основа_обучения", "Основа обучения");            
            list.Add("Тип_конкурса", "Тип конкурса");
            list.Add("Доп_тип_конкурса", "Доп. тип конкурса");
           
            list.Add("сирота", "сирота");
            list.Add("чернобылец", "чернобылец");
            list.Add("военнослужащий", "военнослужащий");
            list.Add("полусирота", "полусирота");
            list.Add("инвалид", "инвалид");
            list.Add("уч.боев.", "уч.боев.");
            list.Add("стажник", "стажник");
            list.Add("реб.-сирота", "реб.-сирота");
            list.Add("огр.возможности", "огр.возможности");
            list.Add("Англ_с_нуля", "Желает изучать англ с нуля");
            list.Add("Англ_оценка", "Итог. оценка по англ");

            list.Add("Номер_приказа_о_зачислении", "Номер приказа о зачислении");
            list.Add("Дата_приказа_о_зачислении", "Дата приказа о зачислении");
            list.Add("Номер_приказа_о_зачислении_иностр", "Номер приказа о зачислении (иностр)");
            list.Add("Дата_приказа_о_зачислении_иностр", "Дата приказа о зачислении (иностр)");
                                  
            string exemQuery = null;

            if (!string.IsNullOrEmpty(_fac) && _fac != "0")
            {
                exemQuery = string.Format("SELECT DISTINCT ed.extExamInEntry.ExamId, ed.extExamInEntry.ExamName AS Name FROM ed.extExamInEntry Where FacultyId={0}", _fac);
            }
            else
            {
                exemQuery = "SELECT DISTINCT ed.extExamInEntry.ExamId, ed.extExamInEntry.ExamName AS Name FROM ed.extExamInEntry";
            }
            
            DataSet dsExams = _bdc.GetDataSet(exemQuery);
            foreach (DataRow dr in dsExams.Tables[0].Rows)
                list.Add(dr["Name"].ToString(), "Экзамен "+dr["Name"].ToString());


            if (MainClass.dbType == PriemType.PriemAspirant)
            {
                list.Add("Медалист", "Красный диплом");
                list.Add("Профиль", "Профиль");
                list.Add("Направление_подготовки", "Базовое_направление(специальность)");
                list.Add("Год_выпуска", "Год выпуска");              

                list.Add("Номер_диплома", "Номер диплома");
                list.Add("Серия_диплома", "Серия диплома");

                list.Add("Квалификация", "Квалификация");
                list.Add("Место_предыдущего_образования_маг", "Место предыдущего образования (для маг)");
            }
            else
            {
                list.Add("Программы_для_ВО", "Программа для лиц с ВО");
                list.Add("Программы_сокр", "Сокращенная программа");
                list.Add("Программы_парал", "Параллельная программа");

                list.Add("Город_уч_заведения", "Город уч.заведения");
                list.Add("Тип_уч_заведения", "Тип уч.заведения");
                list.Add("Медалист", "Медалист");
                list.Add("Номер_школы", "Номер школы");
                list.Add("Серия_атт", "Серия аттестата");
                list.Add("Номер_атт", "Номер аттестата");
                list.Add("Регион_аттестата", "Регион аттестата");                

                list.Add("Серия_диплома", "Серия диплома");
                list.Add("Номер_диплома", "Номер диплома");

                list.Add("Год_выпуска", "Год выпуска");

                list.Add("Средний_балл_атт", "Средний балл аттестата");               
                list.Add("Статус_ФБС", "Статус ФБС");
                list.Add("Поданы_подлинники_ЕГЭ", "Поданы подлинники ЕГЭ");

                list.Add("Профиль", "Профиль");

                list.Add("Свидетельство_ЕГЭ_2011", "Свидетельство ЕГЭ 2011 года");
                list.Add("Свидетельство_ЕГЭ_2012", "Свидетельство ЕГЭ 2012 года");
                list.Add("Загружено_Свид-во_ЕГЭ_2012", "Загружено свид-во ЕГЭ 2012 года");

                list.Add("ЕГЭ_англ.яз", "ЕГЭ англ.яз");
                list.Add("ЕГЭ_русск.язык", "ЕГЭ русск.язык");
                list.Add("ЕГЭ_математика", "ЕГЭ математика");
                list.Add("ЕГЭ_физика", "ЕГЭ физика");
                list.Add("ЕГЭ_химия", "ЕГЭ химия");
                list.Add("ЕГЭ_биология", "ЕГЭ биология");
                list.Add("ЕГЭ_история", "ЕГЭ история");
                list.Add("ЕГЭ_география", "ЕГЭ география");
                list.Add("ЕГЭ_немец.язык", "ЕГЭ немец.язык");
                list.Add("ЕГЭ_франц.язык", "ЕГЭ франц.язык");
                list.Add("ЕГЭ_обществознание", "ЕГЭ обществознание");
                list.Add("ЕГЭ_литература", "ЕГЭ литература");
                list.Add("ЕГЭ_испан.язык", "ЕГЭ испан.язык");
                list.Add("ЕГЭ_информатика", "ЕГЭ информатика");

                list.Add("Аттестат_алгебра", "Aттестат Алгебра");
                list.Add("Аттестат_англ_язык", "Aттестат Англ. язык");
                list.Add("Аттестат_астрономия", "Aттестат Астрономия");
                list.Add("Аттестат_биология", "Aттестат Биология");
                list.Add("Аттестат_вселенная_чел", "Aттестат Вселенная человека");
                list.Add("Аттестат_вс_история", "Aттестат Всеобщая история");
                list.Add("Аттестат_география", "Aттестат География");
                list.Add("Аттестат_геометрия", "Aттестат Геометрия");
                list.Add("Аттестат_информатика", "Aттестат Информатика");
                list.Add("Аттестат_история_Спб", "Aттестат История и культура Санкт-Петербурга");
                list.Add("Аттестат_ист_России", "Aттестат История России");
                list.Add("Аттестат_литература", "Aттестат Литература");
                list.Add("Аттестат_мировая_худ_культура", "Aттестат Мировая художественная культура");
                list.Add("Аттестат_обществознание", "Aттестат Обществознание");
                list.Add("Аттестат_ОБЖ", "Aттестат ОБЖ");
                list.Add("Аттестат_русск_язык", "Aттестат Русский язык");
                list.Add("Аттестат_технология", "Aттестат Технология");
                list.Add("Аттестат_физика", "Aттестат Физика");
                list.Add("Аттестат_физ_культура", "Aттестат Физическая культура");
                list.Add("Аттестат_химия", "Aттестат Химия");
                list.Add("Аттестат_экология", "Aттестат Экология");
                list.Add("Аттестат_немецкий_язык", "Aттестат Немецкий язык");
                list.Add("Аттестат_испанский_язык", "Aттестат Испанский язык");
                list.Add("Аттестат_французский_язык", "Aттестат Французский язык");
                list.Add("Аттестат_итальянский_язык", "Aттестат Итальянский язык");                

                // олимпиады
                //list.Add("Всероссийкая", "Всероссийкая олимпиада");
                //list.Add("Международная", "Международная олимпиада");
                //list.Add("Региональная", "Региональная олимпиада");
                //list.Add("Межвузовская", "Межвузовская олимпиада");
                //list.Add("СПбГУ", "СПбГУ олимпиада");
                //list.Add("Школьников", "Олимпиада школьников");  

                list.Add("Степень_диплома", "Степень диплома"); 

            }

            return list;
        }
        
        //кнопка ОК
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (own is ListAbit)
            {                
                MainClass._config.ClearColumnListAbit();

                foreach (string li in lbYes.Items)
                {
                    MainClass._config.AddColumnNameAbit(_columnList.Keys[_columnList.IndexOfValue(li)]);
                }
            }
            else if (own is ListPersonFilter)
            {
                MainClass._config.ClearColumnListPerson();

                foreach (string li in lbYes.Items)
                {
                    MainClass._config.AddColumnNamePerson(_columnList.Keys[_columnList.IndexOfValue(li)]);
                }
            }

            own.UpdateDataGrid();
            this.Close();
        }

        //функции переноса строк
        private void btnLeft_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbNo, lbYes, false);
        }

        //
        private void btnRight_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbYes, lbNo, false);
        }

        //
        private void btnLeftAll_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbNo, lbYes, true);
        }

        //
        private void btnRightAll_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbYes, lbNo, true);
        }

        //стрелка вверх
        private void btnUp_Click(object sender, EventArgs e)
        {
            if (lbYes.Items.Count == 0)
                return;

            int i = lbYes.SelectedIndex;
            if (i == 0)
                return;

            object obj = lbYes.Items[i];
            lbYes.Items.RemoveAt(i);
            lbYes.Items.Insert(i - 1, obj);
            lbYes.SetSelected(i - 1, true);
        }

        //стрелка вниз
        private void btnDown_Click(object sender, EventArgs e)
        {
            if (lbYes.Items.Count == 0)
                return;

            int i = lbYes.SelectedIndex;
            if (i == lbYes.Items.Count - 1)
                return;

            object obj = lbYes.Items[i];
            lbYes.Items.RemoveAt(i);
            lbYes.Items.Insert(i + 1, obj);
            lbYes.SetSelected(i + 1, true);
        }
    }
}