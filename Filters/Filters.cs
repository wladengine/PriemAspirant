using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using PriemLib;

namespace Priem
{
    public partial class Filters : Form
    {
        private FilterType _currentType;
        private FilterItem _currentItem;
        private int _current;

        private DBPriem _bdc;
        private FormFilter _owner;

        private bool flag;

        //конструтор
        public Filters(FormFilter la, DBPriem bdc, List<iFilter> list)
        {
            InitializeComponent();

            _bdc = bdc;
            _owner = la;

            FillFilters();

            if (list != null)
            {
                foreach (iFilter obj in list)
                {
                    flag = true;
                    AddInList(obj);
                }
                lbFilters.SelectedIndex = 0;
            }
        }

        //заполняем комбобокс фильтрами
        private void FillFilters()
        {
            //абитуриент            
            AddItem(new FilterItem("Ид_номер", FilterType.FromTo, "ed.extPersonAspirant.PersonNum", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Рег_номер", FilterType.FromTo, "ed.qAbiturient.RegNum", "ed.qAbiturient"));
            AddItem(new FilterItem("Номер зачетки", FilterType.FromTo, "ed.qAbiturient.StudyNumber", "ed.qAbiturient"));
            AddItem(new FilterItem("Фамилия", FilterType.Text, "ed.extPersonAspirant.Surname", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Имя", FilterType.Text, "ed.extPersonAspirant.Name", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Отчество", FilterType.Text, "ed.extPersonAspirant.SecondName", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Дата Рождения", FilterType.DateFromTo, "ed.extPersonAspirant.BirthDate", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Номер паспорта", FilterType.FromTo, "ed.extPersonAspirant.PassportNumber", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Серия паспорта", FilterType.FromTo, "ed.extPersonAspirant.PassportSeries", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Дата выдачи паспорта", FilterType.DateFromTo, "ed.extPersonAspirant.PassportDate", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Пол мужской?", FilterType.Bool, "ed.extPersonAspirant.Sex", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Страна", FilterType.Multi, "ed.extPersonAspirant.CountryId", "ed.extPersonAspirant", " SELECT ed.Country.Id, ed.Country.Name FROM ed.Country ORDER BY Name "));
            AddItem(new FilterItem("Гражданство", FilterType.Multi, "ed.extPersonAspirant.NationalityId", "ed.extPersonAspirant", " SELECT ed.Country.Id, ed.Country.Name FROM ed.Country ORDER BY Name "));
            AddItem(new FilterItem("Регион", FilterType.Multi, "ed.extPersonAspirant.RegionId", "ed.extPersonAspirant", " SELECT Id, Name FROM ed.Region "));
            AddItem(new FilterItem("Телефон", FilterType.Text, "ed.extPersonAspirant.Phone", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Предоставлять общежитие на время поступления", FilterType.Bool, "ed.extPersonAspirant.HostelAbit", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Выдано направление на поселение", FilterType.Bool, "ed.extPersonAspirant.HasAssignToHostel", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Выдан экзаменационный пропуск", FilterType.Bool, "ed.extPersonAspirant.HasExamPass", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Факультет, выдавший направление", FilterType.Multi, "ed.extPersonAspirant.HostelFacultyId", "ed.extPersonAspirant", " SELECT Id, Name FROM ed.SP_Faculty "));
            AddItem(new FilterItem("Желает изучать англ с нуля", FilterType.Bool, "ed.extPersonAspirant.StartEnglish", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Сдает ЕГЭ в СПбГУ", FilterType.Bool, "ed.extPersonAspirant.EgeInSpbgu", "ed.extPersonAspirant"));                      

            AddItem(new FilterItem("Факультет", FilterType.Multi, "ed.qAbiturient.FacultyId", "ed.qAbiturient", " SELECT Id, Name FROM ed.qFaculty "));
            AddItem(new FilterItem("Форма обучения", FilterType.Multi, "ed.qAbiturient.StudyFormId", "ed.qAbiturient", "SELECT Id, Name FROM ed.StudyForm "));
            AddItem(new FilterItem("Основа обучения", FilterType.Multi, "ed.qAbiturient.StudyBasisId", "ed.qAbiturient", " SELECT Id, Name FROM ed.StudyBasis "));
            AddItem(new FilterItem("Предоставлять общежитие на время обучения", FilterType.Bool, "ed.qAbiturient.HostelEduc", "ed.qAbiturient"));
            AddItem(new FilterItem("Тип конкурса", FilterType.Multi, "ed.qAbiturient.CompetitionId", "ed.qAbiturient", "SELECT Id, Name FROM ed.Competition ORDER BY Name"));           
            AddItem(new FilterItem("Слушатель", FilterType.Bool, "ed.qAbiturient.IsListener", "ed.qAbiturient"));                      
            AddItem(new FilterItem("Оплатил", FilterType.Bool, "ed.qAbiturient.IsPaid", "ed.qAbiturient"));
            AddItem(new FilterItem("Забрал документы", FilterType.Bool, "ed.qAbiturient.BackDoc", "ed.qAbiturient"));
            AddItem(new FilterItem("Данные проверены", FilterType.Bool, "ed.qAbiturient.Checked", "ed.qAbiturient"));
            AddItem(new FilterItem("Не допущен", FilterType.Bool, "ed.qAbiturient.NotEnabled", "ed.qAbiturient"));
            AddItem(new FilterItem("Дата возврата документов", FilterType.DateFromTo, "ed.qAbiturient.BackDocDate", "ed.qAbiturient"));
            AddItem(new FilterItem("Дата подачи документов", FilterType.DateFromTo, "ed.qAbiturient.DocDate", "ed.qAbiturient"));
            AddItem(new FilterItem("Сумма баллов", FilterType.FromTo, "ed.extAbitMarksSum.TotalSum", "ed.extAbitMarksSum"));
            AddItem(new FilterItem("Ин. язык испытания", FilterType.Multi, "ed.qAbiturient.LanguageId", "ed.qAbiturient", "SELECT Id, Name FROM ed.Language ORDER BY Name"));
            AddItem(new FilterItem("Средний балл сессии", FilterType.FromTo, "ed.qAbiturient.SessionAVG", "ed.qAbiturient"));
            AddItem(new FilterItem("Статус студента", FilterType.FromTo, "ed.qAbiturient.StudentStatus", "ed.qAbiturient"));
            AddItem(new FilterItem("Договор об оплате", FilterType.Bool, " EXISTS (SELECT Top(1) ed.PaidData.Id FROM ed.PaidData WHERE ed.PaidData.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient"));

            //экзамены 
            DataSet dsExams = _bdc.GetDataSet("SELECT DISTINCT ed.extExamInEntry.ExamId AS Id, ed.extExamInEntry.ExamName AS Name FROM ed.extExamInEntry");

            foreach (DataRow dr in dsExams.Tables[0].Rows)
                AddItem(new FilterItem("Экзамен " + dr["Name"].ToString(), FilterType.FromTo, string.Format("(select Min(ed.qMark.Value) FROM ed.qMark INNER JOIN ed.extExamInEntry ON ed.qMark.ExamInEntryId =ed.extExamInEntry.Id WHERE AbiturientId = ed.qAbiturient.Id AND ed.extExamInEntry.ExamId={0})", dr["Id"]), "ed.qAbiturient"));
            
            AddItem(new FilterItem("Номер протокола о допуске", FilterType.FromTo, "ed.extEnableProtocol.Number", "ed.extEnableProtocol"));
            AddItem(new FilterItem("Номер представления к зачислению", FilterType.FromTo, "extEntryView.Number", "extEntryView"));
            AddItem(new FilterItem("Рейтинговый коэффициент", FilterType.FromTo, "ed.qAbiturient.Coefficient", "ed.qAbiturient"));
            AddItem(new FilterItem("Подал подлинники для зачисления", FilterType.Bool, "ed.qAbiturient.HasOriginals", "ed.qAbiturient"));
            AddItem(new FilterItem("Зачислен в СПбГУ (человек)", FilterType.Bool, " EXISTS (SELECT Top(1) ed.extEntryView.Id FROM ed.extEntryView INNER JOIN ed.extAbitAspirant ON ed.extAbitAspirant.id = ed.extentryview.abiturientid WHERE ed.extAbitAspirant.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Зачислен в СПбГУ (заявление)", FilterType.Bool, " EXISTS (SELECT Top(1) ed.extEntryView.Id FROM ed.extEntryView WHERE ed.extEntryView.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient"));
            AddItem(new FilterItem("Отчислен из СПбГУ", FilterType.Bool, " EXISTS (SELECT Top(1) ed.extProtocol.Id FROM ed.extProtocol WHERE ProtocolTypeId = 4 AND IsOld = 0 AND Excluded = 1 AND ed.extProtocol.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient"));
            AddItem(new FilterItem("Есть в представлении к отчислению", FilterType.Bool, " EXISTS (SELECT Top(1) ed.extDisEntryView.Id FROM ed.extDisEntryView WHERE ed.extDisEntryView.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient"));
            AddItem(new FilterItem("Есть в протоколе о допуске", FilterType.Bool, " EXISTS (SELECT Top(1) ed.extEnableProtocol.Id FROM ed.extEnableProtocol WHERE ed.extEnableProtocol.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient"));
            AddItem(new FilterItem("Есть в представлении к зачислению", FilterType.Bool, " EXISTS (SELECT Top(1) ed.extEntryView.Id FROM ed.extEntryView WHERE ed.extEntryView.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient"));

            //льготы
            AddItem(new FilterItem("Военнослужащий", FilterType.Bool, "ed.extPersonAspirant.Privileges & 4 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Сирота", FilterType.Bool, "ed.extPersonAspirant.Privileges & 1 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Чернобылец", FilterType.Bool, "ed.extPersonAspirant.Privileges & 2 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Полусирота", FilterType.Bool, "ed.extPersonAspirant.Privileges & 16 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Инвалид", FilterType.Bool, "ed.extPersonAspirant.Privileges & 32 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Уч. боев.", FilterType.Bool, "ed.extPersonAspirant.Privileges & 64 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Стажник", FilterType.Bool, "ed.extPersonAspirant.Privileges & 128 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Реб.-сирота", FilterType.Bool, "ed.extPersonAspirant.Privileges & 256 ", "ed.extPersonAspirant"));
            AddItem(new FilterItem("Огр. возможности", FilterType.Bool, "ed.extPersonAspirant.Privileges & 512 ", "ed.extPersonAspirant"));

            //Заявления
            AddItem(new FilterItem("Только одно заявление на университет", FilterType.Bool, " (SELECT Count(Id) FROM ed.qAbiturient WHERE PersonId = ed.qAbiturient.PersonId) = 1 ", "ed.qAbiturient"));
            AddItem(new FilterItem("Только одно заявление на ваш факультет", FilterType.Bool, " (SELECT Count(Id) FROM ed.qAbiturient AS ab WHERE ab.PersonId = ed.qAbiturient.PersonId AND ab.FacultyId = ed.qAbiturient.FacultyId) = 1 ", "ed.qAbiturient"));
            AddItem(new FilterItem("Только дневное на ваш факультет", FilterType.Bool, " ( NOT EXISTS (SELECT * FROM ed.qAbiturient ab WHERE ab.PersonId = ed.qAbiturient.PersonId and ab.StudyFormId <> 1)) ", "ed.qAbiturient"));
            AddItem(new FilterItem("Только вечернее на ваш факультет", FilterType.Bool, " ( NOT EXISTS (SELECT * FROM ed.qAbiturient ab WHERE ab.PersonId = ed.qAbiturient.PersonId and  ab.StudyFormId <> 2)) ", "ed.qAbiturient"));
            AddItem(new FilterItem("Только заочное на ваш факультет", FilterType.Bool, " ( NOT EXISTS (SELECT * FROM ed.qAbiturient ab WHERE ab.PersonId = ed.qAbiturient.PersonId and  ab.StudyFormId <> 3)) ", "ed.qAbiturient"));
            AddItem(new FilterItem("Только бюджет на ваш факультет", FilterType.Bool, " ( NOT EXISTS (SELECT * FROM ed.qAbiturient ab WHERE ab.PersonId = ed.qAbiturient.PersonId and  ab.StudyBasisId <> 1)) ", "ed.qAbiturient"));
            AddItem(new FilterItem("Только платно на ваш факультет", FilterType.Bool, " ( NOT EXISTS (SELECT * FROM ed.qAbiturient ab WHERE ab.PersonId = ed.qAbiturient.PersonId and  ab.StudyBasisId <> 2)) ", "ed.qAbiturient"));

            AddItem(new FilterItem("Направление", FilterType.Multi, "ed.qAbiturient.LicenseProgramId", "ed.qAbiturient", " SELECT DISTINCT ed.qLicenseProgram.Id, ed.qLicenseProgram.Code + ' ' + ed.qLicenseProgram.Name AS Name FROM ed.qLicenseProgram "));
            AddItem(new FilterItem("Образовательная программа", FilterType.Multi, "ed.qAbiturient.ObrazProgramId", "ed.qAbiturient", " SELECT DISTINCT ed.qObrazProgram.Id, ed.qObrazProgram.Name AS Name FROM ed.qObrazProgram "));

            if (MainClass.dbType != PriemType.PriemAspirant)
            {
                AddItem(new FilterItem("Программы для лиц с ВО", FilterType.Bool, "ed.qAbiturient.IsSecond", "ed.qAbiturient"));
                AddItem(new FilterItem("Сокращенные программы", FilterType.Bool, "ed.qAbiturient.IsReduced", "ed.qAbiturient"));
                AddItem(new FilterItem("параллельные программы", FilterType.Bool, "ed.qAbiturient.IsParallel", "ed.qAbiturient"));  

                AddItem(new FilterItem("Медалист", FilterType.Bool, "ed.extPersonAspirant.IsExcellent", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Подан подлинник аттестата", FilterType.Bool, "ed.qAbiturient.AttDocOrigin", "ed.qAbiturient"));

                AddItem(new FilterItem("Номер аттестата", FilterType.FromTo, "ed.extPersonAspirant.AttestatNum", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Серия аттестата", FilterType.FromTo, "ed.extPersonAspirant.AttestatSeries", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Регион аттестата", FilterType.FromTo, "ed.extPersonAspirant.AttestatRegion", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Средний балл аттестата", FilterType.FromTo, "ed.extPersonAspirant.SchoolAVG", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Введен средний балл аттестата", FilterType.Bool, "(NOT ed.extPersonAspirant.SchoolAVG IS NULL AND Len(ed.extPersonAspirant.SchoolAVG) > 0)", "ed.extPersonAspirant"));

                AddItem(new FilterItem("Город учебного заведения", FilterType.Text, "ed.extPersonAspirant.SchoolCity", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Тип учебного заведения", FilterType.Multi, "ed.extPersonAspirant.SchoolTypeId", "ed.extPersonAspirant", "SELECT Id, Name FROM SchoolType ORDER BY Name"));
                AddItem(new FilterItem("Название учебного заведения", FilterType.Text, "ed.extPersonAspirant.SchoolName", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Номер учебного заведения", FilterType.FromTo, "ed.extPersonAspirant.SchoolNum", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Год окончания учебного заведения", FilterType.FromTo, "ed.extPersonAspirant.SchoolExitYear", "ed.extPersonAspirant"));

                AddItem(new FilterItem("Профиль", FilterType.Multi, "ed.qAbiturient.ProfileId", "ed.qAbiturient", " SELECT DISTINCT ed.qProfile.Id, ed.qProfile.Name AS Name FROM ed.qProfile "));

                AddItem(new FilterItem("Поданы подлинники свидетельств ЕГЭ", FilterType.Bool, "ed.qAbiturient.EgeDocOrigin", "ed.qAbiturient"));
                AddItem(new FilterItem("Статус ФБС", FilterType.Multi, "(SELECT FBSStatusId FROM ed.extFBSStatus WHERE ed.extFBSStatus.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant", "SELECT Id, Name FROM ed.FBSStatus WHERE Id <> 3"));

                //олимпиады
                AddItem(new FilterItem("Международная олимпиада", FilterType.Bool, " EXISTS (SELECT * FROM ed.Olympiads WHERE ed.Olympiads.AbiturientId = ed.qAbiturient.Id AND ed.Olympiads.OlympTypeId=1 ) ", "ed.qAbiturient"));
                AddItem(new FilterItem("Всероссийская олимпиада", FilterType.Bool, " EXISTS (SELECT * FROM ed.Olympiads WHERE ed.Olympiads.AbiturientId = ed.qAbiturient.Id AND ed.Olympiads.OlympTypeId=2 ) ", "ed.qAbiturient"));
                AddItem(new FilterItem("Олимпиада СПбГУ", FilterType.Bool, " EXISTS (SELECT * FROM ed.Olympiads WHERE ed.Olympiads.AbiturientId = ed.qAbiturient.Id AND ed.Olympiads.OlympTypeId=3 ) ", "ed.qAbiturient"));
                AddItem(new FilterItem("Другие олимпиады школьников", FilterType.Bool, " EXISTS (SELECT * FROM ed.Olympiads WHERE ed.Olympiads.AbiturientId = ed.qAbiturient.Id AND ed.Olympiads.OlympTypeId=4 ) ", "ed.qAbiturient"));

                AddItem(new FilterItem("Степень диплома олимпиады", FilterType.Multi, "(SELECT MAX(OlympValueId) FROM ed.Olympiads WHERE ed.Olympiads.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient", " SELECT Id, Name FROM ed.OlympValue "));

                //ЕГЭ
                AddItem(new FilterItem("Номер свидетельства ЕГЭ 2011 года", FilterType.FromTo, " (SELECT Top(1) ed.EgeCertificate.Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id AND ed.EgeCertificate.Year = 2011)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Номер свидетельства ЕГЭ 2012 года", FilterType.FromTo, " (SELECT Top(1) ed.EgeCertificate.Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id AND ed.EgeCertificate.Year = 2012)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Загружено cвид-во ЕГЭ 2012 года", FilterType.Bool, "EXISTS (select Top(1) ed.EgeCertificate.IsImported FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id AND Year=2012 AND IsImported > 0)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ", FilterType.Bool, " EXISTS (SELECT ed.EgeCertificate.Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));

                AddItem(new FilterItem("Есть свидетельство ЕГЭ 2012 года", FilterType.Bool, " EXISTS (SELECT ed.EgeCertificate.Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id AND ed.EgeCertificate.Year = 2012)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Есть свидетельство ЕГЭ 2011 года", FilterType.Bool, " EXISTS (SELECT ed.EgeCertificate.Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id AND ed.EgeCertificate.Year = 2011)", "ed.extPersonAspirant"));


                AddItem(new FilterItem("Апелляция", FilterType.Bool, " EXISTS (SELECT * FROM ed.EgeMark LEFT JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeMark.IsAppeal>0 AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id) ", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Из олимпиад", FilterType.Bool, " EXISTS (SELECT * FROM ed.EgeMark LEFT JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeMark.IsFromOlymp>0 AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id) ", "ed.extPersonAspirant"));            
                AddItem(new FilterItem("ЕГЭ Английский язык", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Английский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Русский язык", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Русский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Математика", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Математика' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Физика", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Физика' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Химия", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Химия' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Биология", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Биология' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ История", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='История' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ География", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='География' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Немецкий язык", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Немецкий язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Французский язык", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Французский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Обществознание", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Обществознание' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Литература", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Литература' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Испанский язык", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Испанский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("ЕГЭ Информатика и ИКТ", FilterType.FromTo, " (SELECT Max(ed.EgeMark.Value) FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Информатика и ИКТ' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));

                AddItem(new FilterItem("Сдавал ЕГЭ Английский язык", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Английский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Русский язык", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Русский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Математика", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Математика' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Физика", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Физика' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Химия", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Химия' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Биология", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Биология' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ История", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='История' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ География", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='География' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Немецкий язык", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Немецкий язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Французский язык", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Французский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Обществознание", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Обществознание' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Литература", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Литература' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Испанский язык", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Испанский язык' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Сдавал ЕГЭ Информатика и ИКТ", FilterType.Bool, " EXISTS (SELECT * FROM (ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id)INNER JOIN ed.EgeExamName ON ed.EgeMark.EgeExamNameId = ed.EgeExamName.Id WHERE ed.EgeExamName.NAme='Информатика и ИКТ' AND ed.EgeCertificate.PersonId = ed.extPersonAspirant.Id)", "ed.extPersonAspirant"));

                AddItem(new FilterItem("Сдавал ЕГЭ в СПбГУ", FilterType.Bool, " EXISTS (SELECT * FROM ed.qMark INNER JOIN ed.extExamInEntry ON qMark.ExamInEntryId = ed.extExamInEntry.Id WHERE ed.qMark.IsFromEge = 0 AND ed.qMark.IsFromOlymp = 0 AND ed.extExamInEntry.IsAdditional = 0 AND ed.qMark.AbiturientId = ed.qAbiturient.Id)", "ed.qAbiturient"));

            }
            else
            {
                AddItem(new FilterItem("Номер диплома", FilterType.FromTo, "ed.extPersonAspirant.DiplomNum", "ed.extPersonAspirant"));

                AddItem(new FilterItem("Название учебного заведения", FilterType.Text, "ed.extPersonAspirant.HightEducation", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Год окончания учебного заведения", FilterType.FromTo, "ed.extPersonAspirant.HEExitYear", "ed.extPersonAspirant"));

                AddItem(new FilterItem("Профиль", FilterType.Multi, "ed.qAbiturient.ProfileId", "ed.qAbiturient", " SELECT DISTINCT ed.qProfile.Id, ed.qProfile.Name AS Name FROM ed.qProfile "));
                
                AddItem(new FilterItem("Красный диплом", FilterType.Bool, "ed.extPersonAspirant.IsExcellent", "ed.extPersonAspirant"));
                AddItem(new FilterItem("Подан подлинник диплома", FilterType.Bool, "ed.qAbiturient.AttDocOrigin", "ed.qAbiturient"));
            }

            cmbFilters.SelectedIndex = 0;
        }

        //добавление в комбик
        private void AddItem(FilterItem fi)
        {
            cmbFilters.Items.Add(fi);
        }

        private void ShowBool()
        {
            fBool.Visible = true;
            fFromTo.Visible = false;
            fMult.Visible = false;
            fDateFromTo.Visible = false;
            fText.Visible = false;

            fBool.Clear();
        }

        private void ShowText()
        {
            fBool.Visible = false;
            fFromTo.Visible = false;
            fMult.Visible = false;
            fDateFromTo.Visible = false;
            fText.Visible = true;

            fText.Text = string.Empty;
        }

        private void ShowFromTo()
        {
            fBool.Visible = false;
            fFromTo.Visible = true;
            fMult.Visible = false;
            fDateFromTo.Visible = false;
            fText.Visible = false;

            fFromTo.Clear();
            fFromTo.Location = new Point(43, 118);
        }

        private void ShowMult()
        {
            fBool.Visible = false;
            fFromTo.Visible = false;
            fMult.Visible = true;
            fDateFromTo.Visible = false;
            fText.Visible = false;

            fMult.ClearLists();
        }

        private void ShowMultFromTo()
        {
            fBool.Visible = false;
            fFromTo.Visible = true;
            fMult.Visible = true;
            fDateFromTo.Visible = false;
            fText.Visible = false;

            fMult.ClearLists();
            fFromTo.Clear();
            fFromTo.Location = new Point(43, 242);
        }

        private void ShowDateFromTo()
        {
            fBool.Visible = false;
            fFromTo.Visible = false;
            fMult.Visible = false;
            fDateFromTo.Visible = true;
            fText.Visible = false;

            fDateFromTo.Clear();
        }

        //реакция на смену в комбике
        private void cmbFilters_SelectedIndexChanged(object sender, EventArgs e)
        {
            FilterItem fi = cmbFilters.SelectedItem as FilterItem;
            FilterType ft = fi.Type;

            _current = cmbFilters.SelectedIndex;
            _currentItem = fi;
            _currentType = ft;

            switch (ft)
            {
                case FilterType.Bool:
                    ShowBool();
                    break;
                case FilterType.FromTo:
                    ShowFromTo();
                    break;
                case FilterType.Multi:
                    ShowMult();
                    FillMulti();
                    break;
                case FilterType.DateFromTo:
                    ShowDateFromTo();
                    break;
                case FilterType.Text:
                    ShowText();
                    break;
                case FilterType.MultiFromTo:
                    ShowMultFromTo();
                    FillMulti();
                    break;
            }
        }

        //заполняет мультифильтр-юзерконтрол
        private void FillMulti()
        {
            try
            {
                DataSet ds = _bdc.GetDataSet(_currentItem.Query);

                //если ничего нет - выходим
                if (ds.Tables[0].Rows.Count == 0)
                    return;

                List<ListItem> list = new List<ListItem>();

                foreach (DataRow dr in ds.Tables[0].Rows)
                    list.Add(new ListItem(dr["Id"].ToString(), dr["Name"].ToString()));

                fMult.ClearLists();
                fMult.FillList(list, true);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении фильтра:" + ex.Message);
            }
        }

        //стрелка вверх
        private void btnUp_Click(object sender, EventArgs e)
        {
            flag = true;
            if (lbFilters.Items.Count == 0)
                return;

            int i = lbFilters.SelectedIndex;
            if (i == 0)
                return;

            object obj = lbFilters.Items[i];
            lbFilters.Items.RemoveAt(i);
            lbFilters.Items.Insert(i - 1, obj);
            lbFilters.SetSelected(i - 1, true);
            flag = false;
        }

        //стрелка вниз
        private void btnDown_Click(object sender, EventArgs e)
        {
            flag = true;
            if (lbFilters.Items.Count == 0)
                return;

            int i = lbFilters.SelectedIndex;
            if (i == lbFilters.Items.Count - 1)
                return;

            object obj = lbFilters.Items[i];
            lbFilters.Items.RemoveAt(i);
            lbFilters.Items.Insert(i + 1, obj);
            lbFilters.SetSelected(i + 1, true);
        }

        //закрытие с передачей результата
        private void btnExit_Click(object sender, EventArgs e)
        {
            List<iFilter> list = new List<iFilter>();
            int cnt = CheckFilters(list);

            if (cnt < 0)
                return;
            else if (cnt == 0)
                _owner.FilterList = null;
            else
            {
                _owner.FilterList = list;
            }
            _owner.UpdateDataGrid();

            this.Close();
        }

        //проверка и подсчет
        private int CheckFilters(List<iFilter> list)
        {
            int filtCount = 0, leftBrackCount = 0, rightBrackCount = 0;

            foreach (iFilter obj in lbFilters.Items)
            {
                list.Add(obj);

                if (obj is Filter)
                    filtCount++;
                else if (obj is LeftBracket)
                    leftBrackCount++;
                else if (obj is RightBracket)
                    rightBrackCount++;
            }

            if (leftBrackCount != rightBrackCount)
            {
                WinFormsServ.Error("Неправильная расстановка скобок");
                return -1;
            }
            return filtCount;
        }

        //удаление
        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (lbFilters.Items.Count != 0)
            {
                flag = true;
                lbFilters.Items.RemoveAt(lbFilters.SelectedIndex);

                if (lbFilters.Items.Count == 0)
                    btnChange.Enabled = false;
                else
                    lbFilters.SelectedIndex = 0;

                cmbFilters.SelectedIndex = 0;
            }
        }

        //сохранение нового
        private void btnAdd_Click(object sender, EventArgs e)
        {
            Filter f = Save();
            if (f != null)
            {
                AddInList(f);
                btnChange.Enabled = true;
            }
        }

        //сохранение
        private Filter Save()
        {
            switch (_currentType)
            {
                case FilterType.Bool:
                    return AddBool();
                case FilterType.FromTo:
                    return AddFromTo();
                case FilterType.DateFromTo:
                    return AddDateFromTo();
                case FilterType.Multi:
                    return AddMult();
                case FilterType.Text:
                    return AddText();
                case FilterType.MultiFromTo:
                    return AddMultFromTo();
            }
            return null;
        }


        //добавление в список созданных фильтров
        private void AddInList(object filt)
        {
            int i = lbFilters.SelectedIndex;

            lbFilters.Items.Insert(i + 1, filt);
            lbFilters.SelectedItem = filt;
        }

        //добавление нового
        private BoolFilter AddBool()
        {
            bool val = fBool.Value;

            BoolFilter bf = new BoolFilter(_currentItem, val, _current);

            return bf;
        }

        private TextFilter AddText()
        {
            string s = fText.Text.Trim();
            if (s.Length == 0)
            {
                epError.SetError(fText, "Пустое значение");
                return null;
            }
            else
                epError.Clear();

            TextFilter tf = new TextFilter(_currentItem, s, _current);
            return tf;
        }

        //добавление нового
        private FromToFilter AddFromTo()
        {
            if (!fFromTo.CheckValues())
                return null;

            FromToFilter ftf = new FromToFilter(_currentItem, fFromTo.FromValue, fFromTo.ToValue, _current);

            return ftf;
        }

        //добавление нового
        private DateFromToFilter AddDateFromTo()
        {
            if (!fDateFromTo.CheckValues())
                return null;

            DateFromToFilter dftf = new DateFromToFilter(_currentItem, fDateFromTo.FromValue, fDateFromTo.ToValue, _current);

            return dftf;
        }

        private MultiFromToFilter AddMultFromTo()
        {
            if (_currentItem.Name == "Экзамены")
                return AddExams();

            return null;
        }

        //добавление нового
        private MultiSelectFilter AddMult()
        {
            if (fMult.YesCount == 0)
            {
                WinFormsServ.Error("Не выбраны значения!");
                return null;
            }

            MultiSelectFilter msf = new MultiSelectFilter(_currentItem, fMult.GetSelectedItems(), _current);

            return msf;
        }

        private MultiFromToFilter AddExams()
        {
            if (fMult.YesCount == 0)
            {
                WinFormsServ.Error("Не выбраны значения!");
                return null;
            }
            if (!fFromTo.CheckValues())
                return null;

            FromToFilter ftf = new FromToFilter(new FilterItem("Оценка", FilterType.FromTo, "temp.markvalue", null), fFromTo.FromValue, fFromTo.ToValue, _current);
            MultiSelectFilter msf = new MultiSelectFilter(new FilterItem("Экзамены", FilterType.Multi, "qMark.ExamInprogramId", null), fMult.GetSelectedItems(), _current);

            return new MultiFromToFilter(_currentItem, msf, ftf, _current);
        }

        //изменение текущего
        private void btnChange_Click(object sender, EventArgs e)
        {
            if (lbFilters.SelectedIndex < 0)
            {
                return;
            }

            flag = true;

            Filter filter = Save();
            if (filter != null)
                lbFilters.Items[lbFilters.SelectedIndex] = filter;
        }

        //реакция на тычок по созданному фильтру
        private void lbFilters_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbFilters.SelectedIndex < 0)
            {
                btnChange.Enabled = false;
                return;
            }
            else
                btnChange.Enabled = true;

            //если жмут по стрелкам, ничего не происходит
            if (flag)
            {
                flag = false;
                return;
            }

            object obj = lbFilters.SelectedItem;

            if (obj is Bracket)
                btnUp.Enabled = btnDown.Enabled = false;
            else
                btnUp.Enabled = btnDown.Enabled = true;

            if (!(obj is Filter))
                return;

            Filter filter = obj as Filter;

            //проставили в комбик нужный фильтр
            cmbFilters.SelectedIndex = filter.NumInList;

            if (filter is BoolFilter)
            {
                BoolFilter bf = filter as BoolFilter;
                fBool.Value = bf.Value;
            }
            else if (filter is DateFromToFilter)
            {
                DateFromToFilter dftf = filter as DateFromToFilter;

                fDateFromTo.FromValue = dftf.FromValue;
                fDateFromTo.ToValue = dftf.ToValue;
            }
            else if (filter is FromToFilter)
            {
                FromToFilter ftf = filter as FromToFilter;

                fFromTo.FromValue = ftf.FromValue;
                fFromTo.ToValue = ftf.ToValue;
            }
            else if (filter is MultiSelectFilter)
            {
                MultiSelectFilter msf = filter as MultiSelectFilter;

                FillMulti();
                fMult.FillList(msf.List, false);

                foreach (ListItem li in msf.List)
                    fMult.RemoveAtRight(li.Name);
            }
            else if (filter is TextFilter)
            {
                TextFilter tf = filter as TextFilter;

                fText.Text = tf.Value;
            }
            else if (filter is MultiFromToFilter)
            {
                MultiSelectFilter msf = (filter as MultiFromToFilter).MultiSelectFilter;

                FillMulti();
                fMult.FillList(msf.List, false);

                foreach (ListItem li in msf.List)
                    fMult.RemoveAtRight(li.Name);

                FromToFilter ftf = (filter as MultiFromToFilter).FromtoFilter;

                fFromTo.FromValue = ftf.FromValue;
                fFromTo.ToValue = ftf.ToValue;
            }
        }

        //вставка или
        private void btnOr_Click(object sender, EventArgs e)
        {
            AddInList(new Or());
        }

        //вставка скобок
        private void btnBrackets_Click(object sender, EventArgs e)
        {
            AddInList(new LeftBracket());
            AddInList(new RightBracket());
        }

        //очистка
        private void btnClear_Click(object sender, EventArgs e)
        {
            lbFilters.Items.Clear();
        }
    }
}