using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Linq;

using BDClassLib;
using EducServLib;
using PriemLib;

namespace Priem
{
    public partial class Migrator : Form
    {
        private DBPriem _bdc;
        private OleDbClass _odc;
        private string _emptyBase = string.Format(@"{0}\Templates\EMPTYAbiturientDB.mdb", Application.StartupPath);
        private string _metroBase = string.Format(@"{0}\Templates\MetroDB.mdb", Application.StartupPath);

        private long _NewId = 1000001;
        private SortedList<string, long> _slIds;
        private SortedList<string, string> slProf = null;
        private SortedList<string, string> slSpec = null;
        ArrayList _alQueries;
       
        //конструктор
        public Migrator()
        {
            InitializeComponent();
            InitItems();
        }

        //дополнительная инициализация
        private void InitItems()
        {
            this.CenterToParent();
            this.MdiParent = MainClass.mainform;

            _bdc = MainClass.Bdc;
            
            if(MainClass.IsPasha())
                ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.SP_Faculty"), false, true);               
            else
                ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.SP_Faculty"), false, false);           

            btnStart.Enabled = false;
            btnMetro.Enabled = false;
        }

        //старт
        private void btnStart_Click(object sender, EventArgs e)
        {
            string newfile = folderBrowser.SelectedPath + "/AbiturientDB.mdb";

            FileInfo fi = new FileInfo(_emptyBase);
            fi.CopyTo(newfile, true);

            _odc = new OleDbClass();
            _odc.OpenDatabase(newfile);
                        
            _alQueries = new ArrayList();
            PrepareRegion();
            MigrateProfSpez();
            MigrateOrders();
            _odc.ExecuteWithTrasaction(_alQueries);
            MigrateAbits();
            //_odc.ExecuteWithTrasaction(_alQueries);
            MessageBox.Show("Готово!");
            _odc.CloseDataBase();
        }

        public string FacultyId
        {
            get { return ComboServ.GetComboId(cbFaculty); }            
        }

        Dictionary<string, string> _dRegion;
        private void PrepareRegion()
        {
            _dRegion = new Dictionary<string, string>();

            DataSet ds = _bdc.GetDataSet("SELECT * FROM ed.Region");

            foreach (DataRow row in ds.Tables[0].Rows)
                _dRegion.Add(row["Name"].ToString(), row["Id"].ToString());
        }

        //путь сохранения
        private void btnFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbFolder.Text = folderBrowser.SelectedPath;
                btnStart.Enabled = true;
                if(MainClass.IsPasha())
                    btnMetro.Enabled = true;
            }
        }
        
        //миграция
        private void MigrateProfSpez()
        {
            slProf = new SortedList<string, string>();
            slSpec = new SortedList<string, string>();

            string query = @"SELECT DISTINCT LicenseProgramId, FacultyId, LicenseProgramCode, LicenseProgramName  
                       FROM ed.qEntry WHERE qEntry.StudyLevelGroupId = " + MainClass.studyLevelGroupId + ((string.IsNullOrEmpty(FacultyId) ? "" : " AND ed.qEntry.FacultyId = " + FacultyId));

            DataSet dsProf = _bdc.GetDataSet(query);

            int newProfId = 1;
            foreach (DataRow rowProf in dsProf.Tables[0].Rows)
            {
                string queryIns = string.Format("INSERT INTO Profession (Id, Name, Code, FacultyId) VALUES ({0}, '{1}', '{2}', {3})", newProfId, rowProf["LicenseProgramName"].ToString(), rowProf["LicenseProgramCode"].ToString(), rowProf["FacultyId"].ToString());
                _odc.ExecuteQuery(queryIns);

                string key = rowProf["FacultyId"].ToString() + "_" + rowProf["LicenseProgramId"].ToString();
                string value = newProfId.ToString();
                slProf.Add(key, value);

                newProfId++;
            }

            query = @"SELECT DISTINCT LicenseProgramId, FacultyId, ProfileId, ProfileName 
                       FROM ed.qEntry WHERE NOT ProfileId IS NULL AND qEntry.StudyLevelGroupId = " + MainClass.studyLevelGroupId + ((string.IsNullOrEmpty(FacultyId) ? "" : " AND ed.qEntry.FacultyId = " + FacultyId));

            DataSet dsSpec = _bdc.GetDataSet(query);

            int newSpecId = 1;
            foreach (DataRow rowSpec in dsSpec.Tables[0].Rows)
            {               
                string profId = slProf[rowSpec["FacultyId"].ToString() + "_" + rowSpec["LicenseProgramId"].ToString()];

                string queryIns = string.Format("INSERT INTO Specialization (Id, Name, ProfessionId) VALUES ({0}, '{1}', {2})", newSpecId, rowSpec["ProfileName"].ToString(), profId);
                _odc.ExecuteQuery(queryIns);

                string key = string.Format("{0}_{1}_{2}", rowSpec["FacultyId"], rowSpec["LicenseProgramId"], rowSpec["ProfileId"]);
                string value = newSpecId.ToString();
                slSpec.Add(key, value);

                newSpecId++;
            }
        }
        private void MigrateOrders()
        {
            string query = @"SELECT ed.OrderNumbers.*, ed.Protocol.FacultyId, ed.Protocol.StudyFormId, ed.Protocol.StudyBasisId FROM ed.OrderNumbers 
                             INNER JOIN ed.Protocol On ed.OrderNumbers.ProtocolId = ed.Protocol.Id WHERE Protocol.StudyLevelGroupId = " + MainClass.studyLevelGroupId + " " + GetFilter("ed.Protocol");
            string queryAbits;

            DataSet ds = _bdc.GetDataSet(query);
            
            _slIds = new SortedList<string, long>();

            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string s = string.Empty;
                if (dr["OrderNum"].ToString().Length != 0)
                {
                    s = string.Format("INSERT INTO Protocol (Id,FacultyId, SectionId, StudyFormId,Name,FromDate) VALUES ({0},{1},{2},{3},'{4}','{5}')", _NewId, dr["FacultyId"].ToString(), dr["StudyFormId"].ToString(), dr["StudyBasisId"].ToString(), dr["OrderNum"].ToString(), dr["OrderDate"].ToString());
                    _alQueries.Add(s);

                    queryAbits = string.Format(@"SELECT ed.extEntryView.AbiturientId FROM ed.extEntryView INNER JOIN ed.qAbiturient ON ed.extEntryView.AbiturientId = ed.qAbiturient.Id INNER JOIN ed.Person ON ed.qAbiturient.PersonId = ed.Person.Id WHERE ed.extEntryView.Id = '{0}' 
                                                 AND ed.Person.NationalityId=1  ", dr["ProtocolId"].ToString());
                    foreach (DataRow drr in _bdc.GetDataSet(queryAbits).Tables[0].Rows)
                    {
                        _slIds.Add(drr["AbiturientId"].ToString(), _NewId);
                    }
                    _NewId++;          
                }
                if (dr["OrderNumFor"].ToString().Length != 0)
                {
                    s = string.Format("INSERT INTO Protocol (Id,FacultyId, SectionId, StudyFormId,Name,FromDate) VALUES ({0},{1},{2},{3},'{4}','{5}')", _NewId, dr["FacultyId"].ToString(), dr["StudyFormId"].ToString(), dr["StudyBasisId"].ToString(), dr["OrderNumFor"].ToString(), dr["OrderDateFor"].ToString());
                    _alQueries.Add(s);

                    queryAbits = string.Format(@"SELECT ed.extEntryView.AbiturientId FROM ed.extEntryView INNER JOIN ed.qAbiturient ON ed.extEntryView.AbiturientId = ed.qAbiturient.Id INNER JOIN ed.Person ON ed.qAbiturient.PersonId = ed.Person.Id 
                                                 WHERE ed.extEntryView.Id = '{0}' AND ed.Person.NationalityId <> 1 ", dr["ProtocolId"].ToString());
                    foreach (DataRow drr in _bdc.GetDataSet(queryAbits).Tables[0].Rows)
                    {
                        _slIds.Add(drr["AbiturientId"].ToString(), _NewId);
                    }
                    _NewId++;          
                }
            }
        }
        private void MigrateAbits()
        {
            NewWatch wc = new NewWatch(100);
            wc.Show();
            wc.SetText("Загрузка данных...");
            wc.PerformStep();
            using (PriemEntities context = new PriemEntities())
            {
                int iFacultyId = 0;
                int.TryParse(FacultyId, out iFacultyId);

                var abitList =
                    (from Ab in context.Abiturient
                     join Ent in context.extEntry on Ab.EntryId equals Ent.Id
                     join extEV in context.extEntryView on Ab.Id equals extEV.AbiturientId
                     where Ab.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId &&
                     (iFacultyId == 0 ? true : Ab.Entry.FacultyId == iFacultyId)
                     //&& extEV.Date > new DateTime(2013, 9, 18)
                     select new
                     {
                         Ab.Id,
                         Ab.RegNum,
                         Ab.Person.Surname,
                         Ab.Person.Name,
                         Ab.Person.SecondName,
                         Ab.Person.BirthDate,
                         Ab.Person.BirthPlace,
                         Ab.Person.Sex,
                         Ab.Person.PassportTypeId,
                         Ab.Person.PassportSeries,
                         Ab.Person.PassportNumber,
                         Ab.Person.PassportAuthor,
                         Ab.Person.PassportDate,
                         Ab.Person.NationalityId,
                         Ab.Person.Person_Contacts.RegionId,
                         Ab.Person.Person_Contacts.Phone,
                         Ab.Person.Person_Contacts.Mobiles,
                         Ab.LanguageId,
                         Ab.Person.Person_AdditionalInfo.Privileges,
                         Ab.Person.Person_Contacts.Code,
                         Ab.Person.Person_Contacts.City,
                         Ab.Person.Person_Contacts.Street,
                         Ab.Person.Person_Contacts.House,
                         Ab.Person.Person_Contacts.Korpus,
                         Ab.Person.Person_Contacts.Flat,
                         Ab.Person.Person_Contacts.CodeReal,
                         Ab.Person.Person_Contacts.CityReal,
                         Ab.Person.Person_Contacts.StreetReal,
                         Ab.Person.Person_Contacts.HouseReal,
                         Ab.Person.Person_Contacts.KorpusReal,
                         Ab.Person.Person_Contacts.FlatReal,
                         Ab.Person.Person_EducationInfo.SchoolTypeId,
                         Ab.Person.Person_EducationInfo.SchoolCity,
                         Ab.Person.Person_EducationInfo.SchoolName,
                         Ab.Person.Person_EducationInfo.SchoolNum,
                         Ab.Person.Person_EducationInfo.SchoolExitYear,
                         Ab.Person.Person_EducationInfo.AttestatRegion,
                         Ab.Person.Person_EducationInfo.AttestatSeries,
                         Ab.Person.Person_EducationInfo.AttestatNum,
                         Ab.Person.Person_EducationInfo.DiplomSeries,
                         Ab.Person.Person_EducationInfo.DiplomNum,
                         Ab.Person.Person_EducationInfo.IsExcellent,
                         Nation = Ab.Person.Person_Contacts.Country.Name,
                         Ab.Entry.FacultyId,
                         Ab.Entry.StudyFormId,
                         Ab.Entry.StudyBasisId,
                         Ab.Entry.LicenseProgramId,
                         Ab.Entry.ProfileId,
                         Ab.CompetitionId,
                         Ab.StudyNumber,
                         ObrazProgramName = Ab.Entry.SP_ObrazProgram.Name,
                         ObrazProgramCrypt = Ab.Entry.StudyLevel.Acronym + "." + Ab.Entry.SP_ObrazProgram.Number + "." + MainClass.sPriemYear,
                         Ab.DocDate,
                         Ab.IsListener,
                         Ab.Entry.StudyPlanNumber,
                         ListenerTypeId = Ab.Entry.IsSecond ? 1 : (Ab.Entry.IsReduced ? 2 : (Ab.Entry.IsParallel ? 3 : 0)),
                         //EntryProtId = extEV.Id,
                         Ab.Person.Person_EducationInfo.HEExitYear,
                         Ab.Person.Person_AdditionalInfo.HostelEduc,
                         Ab.HasOriginals,
                         Ab.Person.SNILS
                     }).ToList();


                if (abitList.Count() == 0)
                    return;

                wc.SetMax(abitList.Count());
                wc.SetText("Импорт данных...");

                var notApplicableList = abitList.Select(x => x.Id).Except(_slIds.Keys.Select(x => Guid.Parse(x))).ToList();

                if (notApplicableList.Count() > 0)
                {
                    WinFormsServ.Error("Not found orders for " + notApplicableList.Count() + " Ids");
                    return;
                }

                foreach (var Abit in abitList.ToList())
                {
                    string zc = Abit.Code.Replace(" ", "");
                    if (zc.Length > 10)
                        zc = zc.Substring(0, 10);

                    string pa = Abit.PassportAuthor;
                    if (pa.Length > 250)
                        pa = pa.Substring(0, 250);

                    string a = Abit.Code + ", " + Abit.City + ", " + Abit.Street + ", д." + Abit.House + ", " + (Abit.Korpus.Length > 0 ? " к." + Abit.Korpus + ", " : "") + "кв." + Abit.Flat;
                    if (a.Length > 250)
                        a = a.Substring(0, 250);

                    string la = Abit.CodeReal + ", " + Abit.CityReal + ", " + Abit.StreetReal + ", д." + Abit.HouseReal + ", " + (Abit.KorpusReal.Length > 0 ? " к." + Abit.KorpusReal + ", " : "") + "кв." + Abit.FlatReal;
                    if (Abit.CodeReal.Length == 0 && Abit.CityReal.Length == 0 && Abit.StreetReal.Length == 0 && Abit.HouseReal.Length == 0 && Abit.KorpusReal.Length == 0 && Abit.FlatReal.Length == 0)
                        la = "";
                    if (la.Length > 250)
                        la = la.Substring(0, 250);

                    string ph = Abit.Phone + (Abit.Mobiles.Length == 0 ? "" : "; " + Abit.Mobiles);
                    if (ph.Length > 100)
                        ph = ph.Substring(0, 100);

                    string profId = slProf[Abit.FacultyId + "_" + Abit.LicenseProgramId];

                    string specId;
                    if (!Abit.ProfileId.HasValue)
                        specId = "0";
                    else
                        specId = slSpec[Abit.FacultyId + "_" + Abit.LicenseProgramId + "_" + Abit.ProfileId.ToString()];

                    string educSeries = MainClass.studyLevelGroupId == 1 ? Abit.AttestatSeries : Abit.DiplomSeries;
                    string educNum = MainClass.studyLevelGroupId == 1 ? Abit.AttestatNum : Abit.DiplomNum;
                    string educYear = MainClass.studyLevelGroupId == 1 ? (Abit.SchoolExitYear.HasValue ? Abit.SchoolExitYear.ToString() : "") : (Abit.HEExitYear.HasValue ? Abit.HEExitYear.ToString() : "");
                    if (string.IsNullOrEmpty(educYear))
                        educYear = "0";

                    long abId = _slIds[Abit.Id.ToString()];
                    string regionId = _dRegion[Abit.Nation];

                    string s = string.Format(
                        "INSERT INTO Abiturient (" +
                        "[FileNum], [Name], [Patronymic], [Surname], " +
                        "[Privileges], [IsExcellent], [ListenerTypeId], [IsActualListener], " +
                        "[Hostel], [FacultyId], [ProfessionId], [SpecializationId], " +
                        "[StudyFormId], [SectionId], [CompetitionId], " +
                        "[DocDate], [CitizenId], [RegionId], [LanguageId], " +
                        "[AttestatSeries], [AttestatRegion], [AttestatNum], [AttestatCopy], " +
                        "[SchoolName], [SchoolCity], [SchoolNum], [SchoolTypeId], [ExitYear], " +
                        "[Phone], [ZipCode], [Adress], [LifeAddress], " +
                        "[BirthDate], [Sex], " +
                        "[PasswordTypeId], [PaswSeries], [PaswNumber], [PaswDate], [PaswAuthor], " +
                        "[StudyNumber], [EntryOrderId], [EduProgName], [EduProgKod], [StudyPlanNumber], [SNILS])" +
                        "VALUES (" +
                        "'{0}','{1}','{2}','{3}'," +
                        "'{4}','{5}','{6}','{7}', " +
                        "'{8}','{9}','{10}','{11}'," +
                        "'{12}','{13}','{14}'," +
                        "'{15}','{16}','{17}','{18}'," +
                        "'{19}','{20}','{21}','{22}'," +
                        "'{23}','{24}','{25}','{26}','{27}'," +
                        "'{28}','{29}','{30}','{31}'," +
                        "'{32}','{33}'," +
                        "'{34}','{35}','{36}','{37}','{38}'," +
                        "'{39}','{40}', '{41}','{42}', '{43}', '{44}')",
                        Abit.RegNum, Abit.Name, Abit.SecondName, Abit.Surname,
                        Abit.Privileges.ToString(), QueryServ.QueryForBool(Abit.IsExcellent.ToString()), Abit.ListenerTypeId.ToString(), QueryServ.QueryForBool(Abit.IsListener.ToString()),
                        QueryServ.QueryForBool(Abit.HostelEduc.ToString()), Abit.FacultyId.ToString(), profId, specId,
                        Abit.StudyBasisId, Abit.StudyFormId, Abit.CompetitionId,
                        Abit.DocDate, regionId, Abit.RegionId.HasValue ? Abit.RegionId.ToString() : "1",
                        Abit.LanguageId.HasValue ? Abit.LanguageId.Value.ToString() : "1",
                        educSeries, Abit.AttestatRegion, educNum, QueryServ.QueryForBool(Abit.HasOriginals.ToString()),
                        Abit.SchoolName.Replace("'", "").Substring(0, Abit.SchoolName.Length > 200 ? 200 : Abit.SchoolName.Length), Abit.SchoolCity, Abit.SchoolNum, (Abit.SchoolTypeId ?? 1).ToString(), 
                        (string.IsNullOrEmpty(educYear) ? DateTime.Now.Year.ToString() : educYear),
                        ph, zc, a, la,
                        Abit.BirthDate.ToString(), QueryServ.QueryForBool(Abit.Sex.ToString()),
                        Abit.PassportTypeId.ToString(), Abit.PassportSeries, Abit.PassportNumber, Abit.PassportDate.ToString(), pa,
                        Abit.StudyNumber, abId,
                        (Abit.ObrazProgramName.Length > 128 ? Abit.ObrazProgramName.Substring(0, 128) : Abit.ObrazProgramName), Abit.ObrazProgramCrypt, (Abit.StudyPlanNumber ?? ""), Abit.SNILS ?? "");

                    _odc.ExecuteQuery(s);

                    _NewId++;
                    wc.PerformStep();
                }
            }
            wc.Close();
        }        

        //фильтр по факультету
        private string GetFilter(string table)
        {
            string res = string.Empty;

            if (!string.IsNullOrEmpty(FacultyId))
                res += string.Format(" AND {0}.FacultyId={1} ", table, FacultyId);
            
            res += string.Format(" AND {0}.StudyLevelGroupId={1} ", table, MainClass.studyLevelGroupId);

            return res;
        }

        private void btnMetro_Click(object sender, EventArgs e)
        {
            string newfile = folderBrowser.SelectedPath + "/MetroDB.mdb";

            FileInfo fi = new FileInfo(_metroBase);
            fi.CopyTo(newfile, true);

            _alQueries = new ArrayList();

            _odc = new OleDbClass();
            _odc.OpenDatabase(newfile);

            string query = string.Format("SELECT DISTINCT ed.extAbit.Id, ed.Person.Name, ed.Person.SecondName, ed.Person.Surname, " +
                              "ed.Person.BirthDate, ed.extAbit.StudyNumber, ed.extAbit.StudyLevelId, " +
                              "ed.Person.PassportTypeId, case when ed.Person.PassportTypeId=1 then 'Р' when ed.Person.PassportTypeId=3 then 'З' else '' end as PassportType, " +
                              "ed.Person.PassportSeries, ed.Person.PassportNumber, EEE.DateFinishEduc, " +
                              "ed.extEntryView.Id AS EntryProtId " +
                              "FROM ed.extAbit INNER JOIN ed.Person ON ed.extAbit.PersonId = ed.Person.Id " +
                              "INNER JOIN ed.extEntryView ON ed.extEntryView.AbiturientId = ed.extAbit.Id " +
                              "INNER JOIN ed.Entry AS EEE ON EEE.Id = extAbit.EntryId " +
                              "WHERE ed.extAbit.StudyFormId = 1 {0}", GetFilter("extAbit"));

            DataSet ds = _bdc.GetDataSet(query);

            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string ser = dr["PassportSeries"].ToString();
                string ser1 = string.Empty, ser2 = string.Empty;
                string pType = dr["PassportTypeId"].ToString();

                int num = 0;
                string temp = ser.Replace(" ", "").Replace("-", "");
                if (pType == "1" && int.TryParse(temp, out num) && temp.Length == 4)
                {
                    ser = ser.Replace(" ", "").Replace("-", "");
                    ser1 = ser.Substring(0, 2);
                    ser2 = ser.Substring(2, 2);
                }
                else
                    ser1 = ser;

                string dateEnd;
                string course;

                string stLevel = dr["StudyLevelId"].ToString();
                DateTime? dtFinish = dr.Field<DateTime?>("DateFinishEduc");
                
                if (stLevel == "16")
                {
                    course = "1";
                    if (dtFinish.HasValue)
                        dateEnd = dtFinish.Value.ToString("dd.MM.yyyy");
                    else
                        dateEnd = "31.08." + (DateTime.Now.Year + 4).ToString();
                }
                else if (stLevel == "17")
                {
                    course = "5";
                    if (dtFinish.HasValue)
                        dateEnd = dtFinish.Value.ToString("dd.MM.yyyy");
                    else
                        dateEnd = "31.08." + (DateTime.Now.Year + 2).ToString();
                }
                else
                {
                    course = "1";
                    if (dtFinish.HasValue)
                        dateEnd = dtFinish.Value.ToString("dd.MM.yyyy");
                    else
                        dateEnd = "31.08." + (DateTime.Now.Year + 5).ToString();
                }

                string datebirth = ((DateTime)dr["BirthDate"]).ToString("dd.MM.yyyy");
                string OrgCode = "21";
                if (FacultyId == "17")
                    OrgCode = "197";
                if (FacultyId == "29")
                    OrgCode = "105";

                string s = string.Format(
                    "INSERT INTO sList ([DOC_KIND], [DOC_SN], [DOC_S]," +
                    "[DOC_NUM],[SDOCUM],[NAME_F],[NAME_I], [NAME_O]," +
                    "[ORGCODE], [DATEEND], [BIRTHDAY], [COURSE])" +
                    "VALUES ('{0}','{1}','{2}'," +
                    "'{3}','{4}','{5}','{6}','{7}'," +
                    "'" + OrgCode + "','{8}','{9}','{10}')",
                    dr["PassportType"].ToString(), ser1, ser2,
                    dr["PassportNumber"].ToString(), dr["StudyNumber"].ToString(), dr["Surname"].ToString(), dr["Name"].ToString(), dr["SecondName"].ToString(),
                    dateEnd, datebirth, course);

                _alQueries.Add(s);
            }

            _odc.ExecuteWithTrasaction(_alQueries);
            MessageBox.Show("Done!");
        }
    }
}