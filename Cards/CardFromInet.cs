using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data.Objects;
using System.Transactions;

using BaseFormsLib;
using EducServLib;
using PriemLib;

namespace Priem
{
    public partial class CardFromInet : CardFromList
    {
        private DBPriem _bdcInet;
        private int? _abitBarc;
        private int? _personBarc;

        private Guid? personId;
        private bool _closePerson;
        private bool _closeAbit;

        LoadFromInet load;
        private List<ShortCompetition> LstCompetitions;

        private DocsClass _docs;

        // конструктор формы
        public CardFromInet(int? personBarcode, int? abitBarcode, bool closeAbit)
        {
            InitializeComponent();
            _Id = null;
           
            _abitBarc = abitBarcode;
            _personBarc = personBarcode;
            _closeAbit = closeAbit;
            tcCard = tabCard;

            if (_abitBarc == null)
                _closeAbit = true;

            InitControls();     
        }      

        protected override void ExtraInit()
        { 
            base.ExtraInit();

            load = new LoadFromInet();
            _bdcInet = load.BDCInet;
            
            _bdc = MainClass.Bdc;
            _isModified = true;

            if (_personBarc == null)
                _personBarc = (int)_bdcInet.GetValue(string.Format("SELECT Person.Barcode FROM Abiturient INNER JOIN Person ON Abiturient.PersonId = Person.Id WHERE Abiturient.ApplicationCommitNumber = {0}", _abitBarc));

            lblBarcode.Text = _personBarc.ToString();
            if (_abitBarc != null)
                lblBarcode.Text += @"\" + _abitBarc.ToString();

            _docs = new DocsClass(_personBarc.Value, _abitBarc);

            tbNum.Enabled = false;

            rbMale.Checked = true;
            chbEkvivEduc.Visible = false;

            chbHostelAbitYes.Checked = false;
            chbHostelAbitNo.Checked = false;
            chbHostelEducYes.Checked = false;
            chbHostelEducNo.Checked = false;

            cbHEQualification.DropDownStyle = ComboBoxStyle.DropDown;
            
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbPassportType, HelpClass.GetComboListByTable("ed.PassportType"), true, false);
                    ComboServ.FillCombo(cbCountry, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbNationality, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbRegion, HelpClass.GetComboListByTable("ed.Region", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbRegionEduc, HelpClass.GetComboListByTable("ed.Region", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbLanguage, HelpClass.GetComboListByTable("ed.Language"), true, false);
                    ComboServ.FillCombo(cbCountryEduc, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), true, false);                    
                    ComboServ.FillCombo(cbHEStudyForm, HelpClass.GetComboListByTable("ed.StudyForm"), true, false);
                    ComboServ.FillCombo(cbMSStudyForm, HelpClass.GetComboListByTable("ed.StudyForm"), true, false);

                    cbAttestatSeries.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.AttestatSeries AS Name FROM ed.Person_EducationInfo WHERE ed.Person_EducationInfo.AttestatSeries > '' ORDER BY 1");
                    cbHEQualification.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.HEQualification AS Name FROM ed.Person_EducationInfo WHERE NOT ed.Person_EducationInfo.HEQualification IS NULL AND ed.Person_EducationInfo.HEQualification > '' ORDER BY 1");

                    cbAttestatSeries.SelectedIndex = -1;
                    cbHEQualification.SelectedIndex = -1;
                    
                    ComboServ.FillCombo(cbLanguage, HelpClass.GetComboListByTable("ed.Language"), true, false);
                }               

                // магистратура!
                if (MainClass.dbType == PriemType.PriemAspirant)
                {
                    tpEge.Parent = null;
                    tpSecond.Parent = null;

                    
                    gbAtt.Visible = false;
                    gbDipl.Visible = true;
                    chbIsExcellent.Text = "Диплом с отличием";
                    btnAttMarks.Visible = false;
                }
                else
                {
                    tpDocs.Parent = null;
                }

                if (_closeAbit)
                    tpApplication.Parent = null;
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            }
        }

        protected override bool IsForReadOnly()
        {
            return !MainClass.RightsToEditCards();
        }
                
        #region handlers

        //инициализация обработчиков мегакомбов
        protected override void InitHandlers()
        {
            cbCountry.SelectedIndexChanged += new EventHandler(UpdateAfterCountry);
            cbCountryEduc.SelectedIndexChanged += new EventHandler(UpdateAfterCountryEduc);
        }

        protected override void NullHandlers()
        {
            cbCountry.SelectedIndexChanged -= new EventHandler(UpdateAfterCountry);
            cbCountryEduc.SelectedIndexChanged -= new EventHandler(UpdateAfterCountryEduc);
        }

        private void UpdateAfterSchool(object sender, EventArgs e)
        {
            if (SchoolTypeId == MainClass.educSchoolId)
            {
                gbAtt.Visible = true;
                gbDipl.Visible = false;
            }               
            else
            {
                gbAtt.Visible = false;
                gbDipl.Visible = true;
            }
        }
        private void UpdateAfterCountry(object sender, EventArgs e)
        {
            if (CountryId == MainClass.countryRussiaId)
            {
                cbRegion.Enabled = true;
                cbRegion.SelectedItem = "нет";
            }
            else
            {
                cbRegion.Enabled = false;
                cbRegion.SelectedItem = "нет";
            }
        }
        private void UpdateAfterCountryEduc(object sender, EventArgs e)
        {
            if (CountryEducId == MainClass.countryRussiaId)           
                chbEkvivEduc.Visible = false;
            else
                chbEkvivEduc.Visible = true;
        }
        private void chbHostelAbitYes_CheckedChanged(object sender, EventArgs e)
        {
            chbHostelAbitNo.Checked = !chbHostelAbitYes.Checked;           
        }
        private void chbHostelAbitNo_CheckedChanged(object sender, EventArgs e)
        {
            chbHostelAbitYes.Checked = !chbHostelAbitNo.Checked;
        }

        #endregion

        protected override void FillCard()
        {
            try
            {
                FillPersonData(GetPerson());
                FillApplication();
                FillFiles();
            }
            catch (DataException de)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + de.Message);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
            }
        }

        private void FillFiles()
        {
            List<KeyValuePair<string, string>> lstFiles = _docs.UpdateFiles();
            if (lstFiles == null || lstFiles.Count == 0)
                return;

            chlbFile.DataSource = new BindingSource(lstFiles, null);
            chlbFile.ValueMember = "Key";
            chlbFile.DisplayMember = "Value";   
        }

        private extPersonAspirant GetPerson()
        {
            if (_personBarc == null)
                return null;

            try
            {
                if (!MainClass.CheckPersonBarcode(_personBarc))
                {
                    _closePerson = true;

                    using (PriemEntities context = new PriemEntities())
                    {
                        extPersonAspirant person = (from pers in context.extPersonAspirant
                                            where pers.Barcode == _personBarc
                                            select pers).FirstOrDefault();

                        personId = person.Id;

                        tbNum.Text = person.PersonNum.ToString();
                        this.Text = "ПРОВЕРКА ДАННЫХ " + person.FIO;
                        
                        return person;
                    }
                }
                else
                {
                    if (_personBarc == 0)
                        return null;

                    _closePerson = false;
                    personId = null;

                    tcCard.SelectedIndex = 0;
                    tbSurname.Focus();
                                       
                    extPersonAspirant person = load.GetPersonByBarcode(_personBarc.Value); 
                    
                    this.Text = "ЗАГРУЗКА " + person.FIO;
                    return person;
                }
            }

            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
                return null;
            }
        }

        private void FillPersonData(extPersonAspirant person)
        {
            if (person == null)
            {
                WinFormsServ.Error("Не найдены записи!");
                _isModified = false;
                this.Close();
            }

            try
            {
                PersonName = person.Name;
                SecondName = person.SecondName;
                Surname = person.Surname;
                BirthDate = person.BirthDate;
                BirthPlace = person.BirthPlace;
                PassportTypeId = person.PassportTypeId;
                PassportSeries = person.PassportSeries;
                PassportNumber = person.PassportNumber;
                PassportAuthor = person.PassportAuthor;
                PassportDate = person.PassportDate;
                PassportCode = person.PassportCode;
                PersonalCode = person.PersonalCode;
                SNILS = person.SNILS;
                Sex = person.Sex;
                CountryId = person.CountryId;
                NationalityId = person.NationalityId;
                RegionId = person.RegionId;
                Phone = person.Phone;
                Mobiles = person.Mobiles;
                Email = person.Email;
                Code = person.Code;
                City = person.City;
                Street = person.Street;
                House = person.House;
                Korpus = person.Korpus;
                Flat = person.Flat;
                CodeReal = person.CodeReal;
                CityReal = person.CityReal;
                StreetReal = person.StreetReal;
                HouseReal = person.HouseReal;
                KorpusReal = person.KorpusReal;
                FlatReal = person.FlatReal;
                KladrCode = person.KladrCode;
                HostelAbit = person.HostelAbit ?? false;
                HostelEduc = person.HostelEduc ?? false;
                IsExcellent = person.IsExcellent ?? false;
                LanguageId = person.LanguageId;
                SchoolCity = person.SchoolCity;
                SchoolTypeId = person.SchoolTypeId;
                SchoolName = person.SchoolName;
                SchoolNum = person.SchoolNum;
                SchoolExitYear = person.SchoolExitYear;
                CountryEducId = person.CountryEducId;
                RegionEducId = person.RegionEducId;
                IsEqual = person.IsEqual ?? false;
                AttestatRegion = person.AttestatRegion;
                AttestatSeries = person.AttestatSeries;
                AttestatNum = person.AttestatNum;
                DiplomSeries = person.DiplomSeries;
                DiplomNum = person.DiplomNum;
                SchoolAVG = person.SchoolAVG;
                HighEducation = person.HighEducation;
                HEProfession = person.HEProfession;
                HEQualification = person.HEQualification;
                HEEntryYear = person.HEEntryYear;
                HEExitYear = person.HEExitYear;
                HEWork = person.HEWork;
                HEStudyFormId = person.HEStudyFormId;
                Stag = person.Stag;
                WorkPlace = person.WorkPlace;
                MSVuz = person.MSVuz;
                MSCourse = person.MSCourse;
                MSStudyFormId = person.MSStudyFormId;
                Privileges = person.Privileges;
                ExtraInfo = person.ExtraInfo;
                PersonInfo = person.PersonInfo;
                ScienceWork = person.ScienceWork;
                StartEnglish = person.StartEnglish ?? false;
                EnglishMark = person.EnglishMark;

                if (MainClass.dbType == PriemType.Priem)
                {
                    DataTable dtEge = load.GetPersonEgeByBarcode(_personBarc.Value);
                    FillEgeFirst(dtEge);
                }
            }
            catch (DataException de)
            {
                WinFormsServ.Error("Ошибка при заполнении формы (DataException)" + de.Message);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
            } 
        }

        //старый метод FillApplication - прощай, молодость!
        //public void FillApplication()
        //{
        //    try
        //    {
        //        if (_closeAbit || _abitBarc == null)
        //            return;
                               
        //        qAbiturient abit = load.GetAbitByBarcode(_abitBarc.Value);
                
        //        if (abit == null)
        //        {
        //            WinFormsServ.Error("Заявления отсутствуют!");
        //            _isModified = false;
        //            this.Close();
        //        }

        //        IsSecond = abit.IsSecond;
        //        LicenseProgramId = abit.LicenseProgramId;
        //        ObrazProgramId = abit.ObrazProgramId;
        //        ProfileId = abit.ProfileId;
        //        FacultyId = abit.FacultyId;
        //        StudyFormId = abit.StudyFormId;
        //        StudyBasisId = abit.StudyBasisId;
        //        DocDate = (MainClass.dbType == PriemType.PriemMag) ? abit.DocDate : DateTime.Now;
        //        Priority = abit.Priority;

        //        lblHasMotivationLetter.Visible = MainClass.GetHasMotivationLetter(_abitBarc, _personBarc);
        //        lblHasEssay.Visible = MainClass.GetHasEssay(_abitBarc, _personBarc);
        //    }
        //    catch (Exception ex)
        //    {
        //        WinFormsServ.Error("Ошибка при заполнении формы заявления" + ex.Message);
        //    }
        //}

        public void FillApplication()
        {
            try
            {
                string query =
@"SELECT Abiturient.[Id]
,[Priority]
,[PersonId]
,[Priority]
,[Barcode]
,[DateOfStart]
,[EntryId]
,[FacultyId]
,[FacultyName]
,[LicenseProgramId]
,[LicenseProgramCode]
,[LicenseProgramName]
,[ObrazProgramId]
,[ObrazProgramCrypt]
,[ObrazProgramName]
,[ProfileId]
,[ProfileName]
,[StudyBasisId]
,[StudyBasisName]
,[StudyFormId]
,[StudyFormName]
,[StudyLevelId]
,[StudyLevelName]
,[IsSecond]
,[IsReduced]
,[IsParallel]
,[IsGosLine]
,[CommitId]
,[DateOfStart]
,(SELECT MAX(ApplicationCommitVersion.Id) FROM ApplicationCommitVersion WHERE ApplicationCommitVersion.CommitId = [Abiturient].CommitId) AS VersionNum
,(SELECT MAX(ApplicationCommitVersion.VersionDate) FROM ApplicationCommitVersion WHERE ApplicationCommitVersion.CommitId = [Abiturient].CommitId) AS VersionDate
,ApplicationCommit.IntNumber
,[Abiturient].HasInnerPriorities
,[Abiturient].IsApprovedByComission
,[Abiturient].CompetitionId
,[Abiturient].ApproverName
,[Abiturient].DocInsertDate
,[Abiturient].IsCommonRussianCompetition
FROM [Abiturient] 
INNER JOIN ApplicationCommit ON ApplicationCommit.Id = Abiturient.CommitId
WHERE IsCommited = 1 AND IntNumber=@CommitId";

                DataTable tbl = _bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@CommitId", _abitBarc } }).Tables[0];

                LstCompetitions =
                         (from DataRow rw in tbl.Rows
                          select new ShortCompetition(rw.Field<Guid>("Id"), rw.Field<Guid>("CommitId"), rw.Field<Guid>("EntryId"), rw.Field<Guid>("PersonId"),
                              rw.Field<int?>("VersionNum"), rw.Field<DateTime?>("VersionDate"))
                          {
                              Barcode = rw.Field<int>("Barcode"),
                              CompetitionId = rw.Field<int?>("CompetitionId") ?? (rw.Field<int>("StudyBasisId") == 1 ? 4 : 3),
                              CompetitionName = "не указана",
                              HasCompetition = rw.Field<bool>("IsApprovedByComission"),
                              LicenseProgramId = rw.Field<int>("LicenseProgramId"),
                              LicenseProgramName = rw.Field<string>("LicenseProgramName"),
                              ObrazProgramId = rw.Field<int>("ObrazProgramId"),
                              ObrazProgramName = rw.Field<string>("ObrazProgramName"),
                              ProfileId = rw.Field<Guid?>("ProfileId"),
                              ProfileName = rw.Field<string>("ProfileName"),
                              StudyBasisId = rw.Field<int>("StudyBasisId"),
                              StudyBasisName = rw.Field<string>("StudyBasisName"),
                              StudyFormId = rw.Field<int>("StudyFormId"),
                              StudyFormName = rw.Field<string>("StudyFormName"),
                              StudyLevelId = rw.Field<int>("StudyLevelId"),
                              StudyLevelName = rw.Field<string>("StudyLevelName"),
                              FacultyId = rw.Field<int>("FacultyId"),
                              FacultyName = rw.Field<string>("FacultyName"),
                              DocDate = rw.Field<DateTime>("DateOfStart"),
                              DocInsertDate = rw.Field<DateTime?>("DocInsertDate") ?? DateTime.Now,
                              Priority = rw.Field<int>("Priority"),
                              IsGosLine = rw.Field<bool>("IsGosLine"),
                              IsReduced = rw.Field<bool>("IsReduced"),
                              IsSecond = rw.Field<bool>("IsSecond"),
                              HasInnerPriorities = rw.Field<bool>("HasInnerPriorities"),
                              IsApprovedByComission = rw.Field<bool>("IsApprovedByComission"),
                              ApproverName = rw.Field<string>("ApproverName"),
                              lstObrazProgramsInEntry = new List<ShortObrazProgramInEntry>(),
                              IsCommonRussianCompetition = rw.Field<bool>("IsCommonRussianCompetition"),
                          }).ToList();

                if (LstCompetitions.Count == 0)
                {
                    WinFormsServ.Error("Заявления отсутствуют!");
                    _isModified = false;
                    this.Close();
                }

                tbApplicationVersion.Text = (LstCompetitions[0].VersionNum.HasValue ? "№ " + LstCompetitions[0].VersionNum.Value.ToString() : "n/a") +
                    (LstCompetitions[0].VersionDate.HasValue ? (" от " + LstCompetitions[0].VersionDate.Value.ToShortDateString() + " " + LstCompetitions[0].VersionDate.Value.ToShortTimeString()) : "n/a");


                //ObrazProgramInEntry
                foreach (var C in LstCompetitions.Where(x => x.HasInnerPriorities))
                {
                    C.lstObrazProgramsInEntry = new List<ShortObrazProgramInEntry>();
                    query = @"SELECT ObrazProgramInEntryId, ObrazProgramInEntryPriority, ObrazProgramName, ProfileInObrazProgramInEntryId, ProfileInObrazProgramInEntryPriority, ProfileName, 
ISNULL(CurrVersion, 1) AS CurrVersion, ISNULL(CurrDate, GETDATE()) AS CurrDate
FROM [extApplicationDetails] WHERE [ApplicationId]=@AppId";
                    tbl = _bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@AppId", C.Id } }).Tables[0];

                    var data = (from DataRow rw in tbl.Rows
                                select new
                                {
                                    ObrazProgramInEntryId = rw.Field<Guid>("ObrazProgramInEntryId"),
                                    ObrazProgramInEntryPriority = rw.Field<int>("ObrazProgramInEntryPriority"),
                                    ObrazProgramName = rw.Field<string>("ObrazProgramName"),
                                    ProfileInObrazProgramInEntryId = rw.Field<Guid?>("ProfileInObrazProgramInEntryId"),
                                    ProfileInObrazProgramInEntryPriority = rw.Field<int?>("ProfileInObrazProgramInEntryPriority"),
                                    ProfileName = rw.Field<string>("ProfileName"),
                                    CurrVersion = rw.Field<int>("CurrVersion"),
                                    CurrDate = rw.Field<DateTime>("CurrDate")
                                }).ToList();
                    using (PriemEntities context = new PriemEntities())
                    {
                        foreach (var OPIE in data.Select(x => new { x.ObrazProgramInEntryId, x.ObrazProgramInEntryPriority, x.ObrazProgramName, x.CurrDate, x.CurrVersion }).Distinct().OrderBy(x => x.ObrazProgramInEntryPriority))
                        {
                            var OP = new ShortObrazProgramInEntry(OPIE.ObrazProgramInEntryId, OPIE.ObrazProgramName) { ObrazProgramInEntryPriority = OPIE.ObrazProgramInEntryPriority, CurrVersion = OPIE.CurrVersion, CurrDate = OPIE.CurrDate };
                            OP.ListProfiles = new List<ShortProfileInObrazProgramInEntry>();
                            int profPriorVal = 0;
                            foreach (var PROF in data.Where(x => x.ObrazProgramInEntryId == OPIE.ObrazProgramInEntryId && x.ProfileInObrazProgramInEntryId.HasValue).Select(x => new { x.ProfileInObrazProgramInEntryId, ProfileInObrazProgramInEntryPriority = x.ProfileInObrazProgramInEntryPriority, x.ProfileName }).OrderBy(x => x.ProfileInObrazProgramInEntryPriority))
                            {
                                profPriorVal++;
                                OP.ListProfiles.Add(new ShortProfileInObrazProgramInEntry(PROF.ProfileInObrazProgramInEntryId.Value, PROF.ProfileName) { ProfileInObrazProgramInEntryPriority = PROF.ProfileInObrazProgramInEntryPriority ?? profPriorVal });
                            }

                            C.lstObrazProgramsInEntry.Add(OP);
                        }
                    }
                }

                UpdateApplicationGrid();

                //if (_closeAbit || _abitBarc == null)
                //    return;
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы заявления" + ex.Message);
            }
        }

        private void UpdateApplicationGrid()
        {
            dgvApplications.DataSource = LstCompetitions.OrderBy(x => x.Priority)
                .Select(x => new
                {
                    x.Id,
                    x.Priority,
                    x.LicenseProgramName,
                    x.ObrazProgramName,
                    x.ProfileName,
                    x.StudyFormName,
                    x.StudyBasisName,
                    x.HasCompetition,
                    comp = x.lstObrazProgramsInEntry.Count > 0 ? "приоритеты" : ""
                }).ToList();
            dgvApplications.Columns["Id"].Visible = false;
            dgvApplications.Columns["Priority"].HeaderText = "Приор";
            dgvApplications.Columns["Priority"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader;
            dgvApplications.Columns["LicenseProgramName"].HeaderText = "Направление";
            dgvApplications.Columns["ObrazProgramName"].HeaderText = "Образ. программа";
            dgvApplications.Columns["ProfileName"].HeaderText = "Профиль";
            dgvApplications.Columns["StudyFormName"].HeaderText = "Форма обуч";
            dgvApplications.Columns["StudyBasisName"].HeaderText = "Основа обуч";
            dgvApplications.Columns["comp"].HeaderText = "";
            dgvApplications.Columns["HasCompetition"].Visible = false;
        }
        private void dgvApplications_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if ((bool)dgvApplications["HasCompetition", e.RowIndex].Value)
                {
                    e.CellStyle.BackColor = Color.Cyan;
                    e.CellStyle.SelectionBackColor = Color.Cyan;
                }
            }
        }

        protected override void SetReadOnlyFieldsAfterFill()
        {
            base.SetReadOnlyFieldsAfterFill();
            
            if (_closePerson)
            {
                tcCard.SelectedTab = tpApplication;

                foreach (TabPage tp in tcCard.TabPages)
                {
                    if (tp != tpApplication && tp != tpDocs)
                    {
                        foreach (Control control in tp.Controls)
                        {
                            control.Enabled = false;
                            foreach (Control crl in control.Controls)
                                crl.Enabled = false;
                        }
                    }
                }
            }

            if (MainClass.dbType == PriemType.PriemAspirant)
                btnSaveChange.Text = "Одобрить";
        }
       
        private void FillEgeFirst(DataTable dtEge)
        {
            if (MainClass.dbType == PriemType.PriemAspirant)
                return;
            
            try
            {                
                DataTable examTable = new DataTable();

                DataColumn clm;
                clm = new DataColumn();
                clm.ColumnName = "Предмет";
                clm.ReadOnly = true;
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "ExamId";
                clm.ReadOnly = true;
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Баллы";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Номер сертификата";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Типографский номер";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "EgeCertificateId";
                examTable.Columns.Add(clm);


                string defQuery = "SELECT ed.EgeExamName.Name AS 'Предмет', ed.EgeExamName.Id AS ExamId FROM ed.EgeExamName";
                DataSet ds = _bdc.GetDataSet(defQuery);
                foreach (DataRow dsRow in ds.Tables[0].Rows)
                {
                    DataRow newRow;
                    newRow = examTable.NewRow();
                    newRow["Предмет"] = dsRow["Предмет"].ToString();
                    newRow["ExamId"] = dsRow["ExamId"].ToString();
                    examTable.Rows.Add(newRow);
                }

                foreach (DataRow dsRow in dtEge.Rows)
                {
                    for (int i = 0; i < examTable.Rows.Count; i++)
                    {
                        if (examTable.Rows[i]["ExamId"].ToString() == dsRow["ExamId"].ToString())
                        {
                            examTable.Rows[i]["Баллы"] = dsRow["Value"].ToString();
                            examTable.Rows[i]["Номер сертификата"] = dsRow["Number"].ToString();                            
                        }
                    }
                }

                DataView dv = new DataView(examTable);
                dv.AllowNew = false;

                dgvEGE.DataSource = dv;
                dgvEGE.Columns["ExamId"].Visible = false;
                dgvEGE.Columns["EgeCertificateId"].Visible = false;               

                dgvEGE.Columns["Предмет"].Width = 162;
                dgvEGE.Columns["Баллы"].Width = 45;
                dgvEGE.Columns["Номер сертификата"].Width = 110;
                dgvEGE.ReadOnly = false;

                dgvEGE.Update();
            }
            catch (DataException de)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + de.Message);
            }
        }
           

        #region Save

        // проверка на уникальность абитуриента
        private bool CheckIdent()
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

                if(_Id == null)
                    context.CheckPersonIdent(Surname, PersonName, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatRegion, AttestatSeries, AttestatNum, boolPar);
                else
                    context.CheckPersonIdentWithId(Surname, PersonName, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatRegion, AttestatSeries, AttestatNum, GuidId, boolPar);

                return Convert.ToBoolean(boolPar.Value);
            }
        }

        protected override bool CheckFields()
        {
            if (Surname.Length <= 0)
            {
                epError.SetError(tbSurname, "Отсутствует фамилия абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (PersonName.Length <= 0)
            {
                epError.SetError(tbName, "Отсутствует имя абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            //Для О'Коннор сделал добавку в регулярное выражение: \'
            if (!Regex.IsMatch(Surname, @"^[А-Яа-яёЁ\-\'\s]+$"))
            {
                epError.SetError(tbSurname, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Regex.IsMatch(PersonName, @"^[А-Яа-яёЁ\-\s]+$"))
            {
                epError.SetError(tbName, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Regex.IsMatch(SecondName, @"^[А-Яа-яёЁ\-\s]*$"))
            {
                epError.SetError(tbSecondName, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (SecondName.StartsWith("-"))
            {
                SecondName = SecondName.Replace("-", "");
            }

            // проверка на англ. буквы
            if (!Util.IsRussianString(PersonName))
            {
                epError.SetError(tbName, "Имя содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Util.IsRussianString(Surname))
            {
                epError.SetError(tbSurname, "Фамилия содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Util.IsRussianString(SecondName))
            {
                epError.SetError(tbSecondName, "Отчество содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (BirthDate == null)
            {
                epError.SetError(dtBirthDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            int checkYear = DateTime.Now.Year - 12;
            if (BirthDate.Value.Year > checkYear || BirthDate.Value.Year < 1920)
            {
                epError.SetError(dtBirthDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (PassportDate.Value.Year > DateTime.Now.Year || PassportDate.Value.Year < 1970)
            {
                epError.SetError(dtPassportDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (PassportTypeId == MainClass.pasptypeRFId)
            {
                if (!(PassportSeries.Length == 4))
                {
                    epError.SetError(tbPassportSeries, "Неправильно введена серия паспорта РФ абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();

                if (!(PassportNumber.Length == 6))
                {
                    epError.SetError(tbPassportNumber, "Неправильно введен номер паспорта РФ абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();
            }

            if (NationalityId == MainClass.countryRussiaId)
            {
                if (PassportSeries.Length <= 0)
                {
                    epError.SetError(tbPassportSeries, "Отсутствует серия паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();

                if (PassportNumber.Length <= 0)
                {
                    epError.SetError(tbPassportNumber, "Отсутствует номер паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();
            }

            if (PassportSeries.Length > 10)
            {
                epError.SetError(tbPassportSeries, "Слишком длинное значение серии паспорта абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();


            if (PassportNumber.Length > 20)
            {
                epError.SetError(tbPassportNumber, "Слишком длинное значение номера паспорта абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!chbHostelAbitYes.Checked && !chbHostelAbitNo.Checked)
            {
                epError.SetError(chbHostelAbitNo, "Не указаны данные о предоставлении общежития");
                tabCard.SelectedIndex = 1;
                return false;
            }
            else
                epError.Clear();

            if (gbAtt.Visible && AttestatNum.Length <= 0)
            {
                epError.SetError(tbAttestatNum, "Отсутствует номер аттестата абитуриента");
                tabCard.SelectedIndex = 2;
                return false;
            }
            else
                epError.Clear();

            double d = 0;
            if (tbSchoolAVG.Text.Trim() != "")
            {
                if (!double.TryParse(tbSchoolAVG.Text.Trim().Replace(".", ","), out d))
                {
                    epError.SetError(tbSchoolAVG, "Неправильный формат");
                    tabCard.SelectedIndex = 2;
                    return false;
                }
                else
                    epError.Clear();
            }

            //if (tbHEProfession.Text.Length >= 100)
            //{
            //    epError.SetError(tbHEProfession, "Длина поля превышает 100 символов.");
            //    tabCard.SelectedIndex = 2;
            //    return false;
            //}
            //else
            //    epError.Clear();

            if (tbScienceWork.Text.Length >= 2000)
            {
                epError.SetError(tbScienceWork, "Длина поля превышает 2000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (tbExtraInfo.Text.Length >= 1000)
            {
                epError.SetError(tbExtraInfo, "Длина поля превышает 1000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (tbPersonInfo.Text.Length > 1000)
            {
                epError.SetError(tbPersonInfo, "Длина поля превышает 1000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (tbWorkPlace.Text.Length > 1000)
            {
                epError.SetError(tbWorkPlace, "Длина поля превышает 1000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (!CheckIdent())
            {
                WinFormsServ.Error("В базе уже существует абитуриент с такими же либо ФИО, либо данными паспорта, либо данными аттестата!");
                return false;
            }

            return true;
        }

        private bool CheckFieldsAbit()
        {
            //if (LstCompetitions.Where(x => !x.HasCompetition).Count() > 0)
            //{
            //    epError.SetError(dgvApplications, "Не по всем конкурсным позициям указаны типы конкурсов");
            //    tabCard.SelectedIndex = 5;
            //    return false;
            //}
            //else
            //    epError.Clear();

            return true;

            //using (PriemEntities context = new PriemEntities())
            //{
            //    if (LicenseProgramId == null || ObrazProgramId == null || FacultyId == null || StudyFormId == null || StudyBasisId == null)
            //    {
            //        epError.SetError(cbLicenseProgram, "Прием документов на данную программу не осуществляется!");
            //        tabCard.SelectedIndex = 0;
            //        return false;
            //    }
            //    else
            //        epError.Clear();

            //    if (EntryId == null)
            //    {
            //        epError.SetError(cbLicenseProgram, "Прием документов на данную программу не осуществляется!");
            //        tabCard.SelectedIndex = 0;
            //        return false;
            //    }
            //    else
            //        epError.Clear();

            //    if (!CheckIsClosed(context))
            //    {
            //        epError.SetError(cbLicenseProgram, "Прием документов на данную программу закрыт!");
            //        tabCard.SelectedIndex = 0;
            //        return false;
            //    }
            //    else
            //        epError.Clear();


            //    if (!CheckIdent(context))
            //    {
            //        WinFormsServ.Error("У абитуриента уже существует заявление на данный факультет, направление, профиль, форму и основу обучения!");
            //        return false;
            //    }

            //    if (!CheckThreeAbits(context))
            //    {
            //        WinFormsServ.Error("У абитуриента уже существует 3 заявления на различные образовательные программы!");
            //        return false;
            //    }

            //    if (!chbHostelEducYes.Checked && !chbHostelEducNo.Checked)
            //    {
            //        epError.SetError(chbHostelEducNo, "Не указаны данные о предоставлении общежития");
            //        tabCard.SelectedIndex = 0;
            //        return false;
            //    }
            //    else
            //        epError.Clear();

            //    if (DocDate > DateTime.Now)
            //    {
            //        epError.SetError(dtDocDate, "Неправильная дата");
            //        tabCard.SelectedIndex = 1;
            //        return false;
            //    }
            //    else
            //        epError.Clear();               
            //}
            //
            //return true;
        } 
        
        private bool CheckIsClosed(PriemEntities context, Guid EntryId)
        {                  
            bool isClosed = (from ent in context.qEntry
                                where ent.Id == EntryId
                                select ent.IsClosed).FirstOrDefault();
            return !isClosed;
        }

        //// проверка на уникальность заявления
        //private bool CheckIdent(PriemEntities context)
        //{
        //    ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

        //    if (personId != null)
        //        context.CheckAbitIdent(personId, EntryId, boolPar);         

        //    return Convert.ToBoolean(boolPar.Value);
        //}
        //private bool CheckThreeAbits(PriemEntities context)
        //{
        //    return SomeMethodsClass.CheckThreeAbits(context, personId, LicenseProgramId, ObrazProgramId, ProfileId);
        //}

        protected override bool SaveClick()
        {
            try
            {
                if (_closePerson)
                {
                    if (!SaveApplication(personId.Value))
                        return false;
                }
                else
                {
                    if (!CheckFields())
                        return false;

                    using (PriemEntities context = new PriemEntities())
                    {
                        using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                        {
                            try
                            {
                                ObjectParameter entId = new ObjectParameter("id", typeof(Guid));
                                context.Person_insert(_personBarc, PersonName, SecondName, Surname, BirthDate, BirthPlace, PassportTypeId, PassportSeries, PassportNumber,
                                    PassportAuthor, PassportDate, Sex, CountryId, NationalityId, RegionId, Phone, Mobiles, Email,
                                    Code, City, Street, House, Korpus, Flat, CodeReal, CityReal, StreetReal, HouseReal, KorpusReal, FlatReal, KladrCode, HostelAbit, HostelEduc, false,
                                    null, false, null, IsExcellent, LanguageId, SchoolCity, SchoolTypeId, SchoolName, SchoolNum, SchoolExitYear,
                                    SchoolAVG, CountryEducId, RegionEducId, IsEqual, AttestatRegion, AttestatSeries, AttestatNum, DiplomSeries, DiplomNum, HighEducation, HEProfession,
                                    HEQualification, HEEntryYear, HEExitYear, HEStudyFormId, HEWork, Stag, WorkPlace, MSVuz, MSCourse, MSStudyFormId, Privileges, PassportCode,
                                    PersonalCode, PersonInfo, ExtraInfo, ScienceWork, StartEnglish, EnglishMark, EgeInSpbgu, SNILS, entId);

                                personId = (Guid)entId.Value;

                                SaveEgeFirst();
                                transaction.Complete();
                            }
                            catch (Exception exc)
                            {
                                WinFormsServ.Error(exc, "Ошибка при сохранении:");
                            }
                        }
                        if (!SaveApplication(personId.Value))
                        {
                            _closePerson = true;
                            return false;
                        }
                        
                        _bdcInet.ExecuteQuery("UPDATE Person SET IsImported = 1 WHERE Person.Barcode = " + _personBarc);                       
                    }
                }  
                             
                _isModified = false;

                OnSave();               

                this.Close();
                return true;
            }
            catch (Exception de)
            {
                WinFormsServ.Error("Ошибка обновления данных" + de.Message);
                return false;
            }
        }

        private bool SaveApplication(Guid PersonId)
        {
            if (_closeAbit)
                return true;

            if (personId == null)
                return false;

            if (!CheckFieldsAbit())
                return false;

            try
            {
                using (TransactionScope trans = new TransactionScope(TransactionScopeOption.Required))
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        ObjectParameter entId = new ObjectParameter("id", typeof(Guid));

                        if (personId.HasValue)
                        {
                            var notUsedApplications = context.Abiturient.Where(x => x.PersonId == personId && !x.BackDoc && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId).Select(x => x.EntryId).ToList().Except(LstCompetitions.Select(x => x.EntryId)).ToList();
                            if (notUsedApplications.Count > 0)
                            {
                                var dr = MessageBox.Show("У абитуриента в базе имеются " + notUsedApplications.Count +
                                    " конкурсов, не перечисленных в заявлении. Вероятно, по ним был уже произведён отказ. Проставить по данным конкурсным позициям отказ от участия в конкурсе?",
                                    "Внимание!", MessageBoxButtons.YesNo);
                                if (dr == System.Windows.Forms.DialogResult.Yes)
                                {
                                    string str = "У меня есть на руках заявление об отказе в участии в следующих конкурсах:";
                                    int incrmntr = 1;
                                    foreach (var app_entry in notUsedApplications)
                                    {
                                        var entry = context.Entry.Where(x => x.Id == app_entry).FirstOrDefault();
                                        str += "\n" + incrmntr++ + ")" + entry.SP_LicenseProgram.Code + " " + entry.SP_LicenseProgram.Name + "; "
                                            + entry.StudyLevel.Acronym + "." + entry.SP_ObrazProgram.Number + " " + entry.SP_ObrazProgram.Name +
                                            ";\nПрофиль:" + entry.ProfileName + ";" + entry.StudyForm.Acronym + ";" + entry.StudyBasis.Acronym;
                                    }
                                    dr = MessageBox.Show(str, "Внимание!", MessageBoxButtons.YesNo);
                                    if (dr == System.Windows.Forms.DialogResult.Yes)
                                    {
                                        foreach (var app_entry in notUsedApplications)
                                        {
                                            var applst = context.Abiturient.Where(x => x.EntryId == app_entry && x.PersonId == personId && !x.BackDoc && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId).Select(x => x.Id).ToList();
                                            foreach (var app in applst)
                                            {
                                                context.Abiturient_UpdateBackDoc(true, DateTime.Now, app);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        foreach (var Comp in LstCompetitions)
                        {
                            var DocDate = Comp.DocDate;
                            var DocInsertDate = Comp.DocInsertDate == DateTime.MinValue ? DateTime.Now : Comp.DocInsertDate;

                            bool isViewed = Comp.HasCompetition;
                            Guid ApplicationId = Comp.Id;
                            bool hasLoaded = context.Abiturient.Where(x => x.PersonId == PersonId && x.EntryId == Comp.EntryId && !x.BackDoc).Count() == 0;
                            if (hasLoaded)
                            {
                                context.Abiturient_InsertDirectly(PersonId, Comp.EntryId, Comp.CompetitionId, Comp.IsListener,
                                    false, false, false, null, DocDate, DocInsertDate,
                                    false, false, null, Comp.OtherCompetitionId, Comp.CelCompetitionId, Comp.CelCompetitionText,
                                    LanguageId, Comp.HasOriginals, Comp.Priority, Comp.Barcode, Comp.CommitId, _abitBarc, Comp.IsGosLine, isViewed, ApplicationId);
                                context.Abiturient_UpdateIsCommonRussianCompetition(Comp.IsCommonRussianCompetition, ApplicationId);
                            }
                            else
                            {
                                ApplicationId = context.Abiturient.Where(x => x.PersonId == PersonId && x.EntryId == Comp.EntryId && !x.BackDoc).Select(x => x.Id).First();
                                context.Abiturient_UpdatePriority(Comp.Priority, ApplicationId);
                            }
                            if (Comp.lstObrazProgramsInEntry.Count > 0)
                            {
                                //загружаем внутренние приоритеты по профилям
                                int currVersion = Comp.lstObrazProgramsInEntry.Select(x => x.CurrVersion).FirstOrDefault();
                                DateTime currDate = Comp.lstObrazProgramsInEntry.Select(x => x.CurrDate).FirstOrDefault();
                                Guid ApplicationVersionId = Guid.NewGuid();
                                context.ApplicationVersion.AddObject(new ApplicationVersion() { IntNumber = currVersion, Id = ApplicationVersionId, ApplicationId = ApplicationId, VersionDate = currDate });
                                foreach (var OPIE in Comp.lstObrazProgramsInEntry)
                                {
                                    context.Abiturient_UpdateObrazProgramInEntryPriority(OPIE.Id, OPIE.ObrazProgramInEntryPriority, ApplicationId);

                                    foreach (var ProfInOPIE in OPIE.ListProfiles)
                                    {
                                        context.Abiturient_UpdateProfileInObrazProgramInEntryPriority(ProfInOPIE.Id, ProfInOPIE.ProfileInObrazProgramInEntryPriority, ApplicationId);

                                        context.ApplicationVersionDetails.AddObject(new ApplicationVersionDetails()
                                        {
                                            ApplicationVersionId = ApplicationVersionId,
                                            ObrazProgramInEntryId = OPIE.Id,
                                            ObrazProgramInEntryPriority = OPIE.ObrazProgramInEntryPriority,
                                            ProfileInObrazProgramInEntryId = ProfInOPIE.Id,
                                            ProfileInObrazProgramInEntryPriority = ProfInOPIE.ProfileInObrazProgramInEntryPriority
                                        });
                                    }
                                }
                            }
                        }

                        context.SaveChanges();

                        //context.Abiturient_Insert(personId, EntryId, CompetitionId, HostelEduc, IsListener, WithHE, false, false, null, DocDate, DateTime.Now,
                        //AttDocOrigin, EgeDocOrigin, false, false, null, OtherCompetitionId, CelCompetitionId, CelCompetitionText, LanguageId, false,
                        //Priority, _abitBarc, entId);
                    }

                    trans.Complete();

                    if (!MainClass.IsTestDB)
                        _bdcInet.ExecuteQuery("UPDATE ApplicationCommit SET IsImported = 1 WHERE IntNumber = '" + _abitBarc + "'");

                    return true;
                }
            }
            catch (Exception de)
            {
                WinFormsServ.Error("Ошибка обновления данных Abiturient\n" + de.Message + "\n" + de.InnerException.Message);
                return false;
            }
        }
       
        private void SaveEgeFirst()
        {
            if (MainClass.dbType == PriemType.PriemAspirant)
                return;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    EgeList egeLst = new EgeList();

                    foreach (DataGridViewRow dr in dgvEGE.Rows)
                    {
                        if (dr.Cells["Баллы"].Value.ToString().Trim() != string.Empty)
                            egeLst.Add(new EgeMarkCert(dr.Cells["ExamId"].Value.ToString().Trim(), dr.Cells["Баллы"].Value.ToString().Trim(), dr.Cells["Номер сертификата"].Value.ToString().Trim(), dr.Cells["Типографский номер"].Value.ToString()));
                    }
                   
                    foreach (EgeCertificateClass cert in egeLst.EGEs.Keys)
                    {
                        // проверку на отсутствие одинаковых свидетельств
                        int res = (from ec in context.EgeCertificate
                                   where ec.Number == cert.Doc
                                   select ec).Count(); 
                        if (res > 0)
                        {
                            WinFormsServ.Error(string.Format("Свидетельство с номером {0} уже есть в базе, поэтому сохранено не будет!", cert.Doc));
                            continue;
                        }                        

                        ObjectParameter ecId = new ObjectParameter("id", typeof(Guid));
                        context.EgeCertificate_Insert(cert.Doc, cert.Tipograf, "20" + cert.Doc.Substring(cert.Doc.Length - 2, 2), personId, null, false, ecId);

                        Guid? certId = (Guid?)ecId.Value;
                        foreach (EgeMarkCert mark in egeLst.EGEs[cert])
                        {
                            int val;
                            if(!int.TryParse(mark.Value, out val))
                                continue;
                            
                            int subj;
                            if(!int.TryParse(mark.Subject, out subj))
                                continue;
                                                       
                            context.EgeMark_Insert((int?)val, (int?)subj, certId, false, false);                            
                        }
                    }                   
                }

            }
            catch (Exception de)
            {          
                WinFormsServ.Error("Ошибка сохранения данные ЕГЭ - данные не были сохранены. Введите их заново! \n" + de.Message);
            }
        }

        public bool IsMatchEgeNumber(string number)
        {
            string num = number.Trim();
            if (Regex.IsMatch(num, @"^\d{2}-\d{9}-(12|13|14)$"))//не даёт перегрузить воякам свои древние ЕГЭ, добавлен 2010 год
                return true;
            else
                return false;
        }

        #endregion 
        
        protected override void OnClosed()
        {
            base.OnClosed();
            load.CloseDB();                
        }

        protected override void OnSave()
        {
            base.OnSave();
            using (PriemEntities context = new PriemEntities())
            {
                Guid? perId = (from per in context.extPersonAll
                                where per.Barcode == _personBarc
                                select per.Id).FirstOrDefault();

                MainClassCards.OpenCardPerson(MainClass.mainform, perId.ToString(), null, null);
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<string, string>> lstFiles = new List<KeyValuePair<string, string>>();
            foreach (KeyValuePair<string, string> file in chlbFile.CheckedItems)
            {
                lstFiles.Add(file);
            }

            _docs.OpenFile(lstFiles);
        }

        private void tabCard_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.D1)
                this.tcCard.SelectedIndex = 0;
            if (e.Control && e.KeyCode == Keys.D2)
                this.tcCard.SelectedIndex = 1;
            if (e.Control && e.KeyCode == Keys.D3)
                this.tcCard.SelectedIndex = 2;
            if (e.Control && e.KeyCode == Keys.D4)
                this.tcCard.SelectedIndex = 3;
            if (e.Control && e.KeyCode == Keys.D5)
                this.tcCard.SelectedIndex = 4;
            if (e.Control && e.KeyCode == Keys.D6)
                this.tcCard.SelectedIndex = 5;
            if (e.Control && e.KeyCode == Keys.D7)
                this.tcCard.SelectedIndex = 6;
            if (e.Control && e.KeyCode == Keys.D8)
                this.tcCard.SelectedIndex = 7;
            if (e.Control && e.KeyCode == Keys.S)
                SaveRecord();
        }

        private void dgvApplications_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rwNum = e.RowIndex;
            OpenCardCompetitionInInet(rwNum);
        }
        private void btnOpenCompetition_Click(object sender, EventArgs e)
        {
            if (dgvApplications.SelectedCells.Count == 0)
                return;

            int rwNum = dgvApplications.SelectedCells[0].RowIndex;
            OpenCardCompetitionInInet(rwNum);
        }
        private ShortCompetition GetCompFromGrid(int rwNum)
        {
            if (rwNum < 0)
                return null;

            Guid Id = (Guid)dgvApplications["Id", rwNum].Value;
            return LstCompetitions.Where(x => x.Id == Id).FirstOrDefault();
        }
        private void OpenCardCompetitionInInet(int rwNum)
        {
            if (rwNum >= 0)
            {
                var ent = GetCompFromGrid(rwNum);
                if (ent != null)
                {
                    var crd = new CardCompetitionInInet(ent);
                    crd.OnUpdate += UpdateCommitCompetition;
                    crd.Show();
                }
            }
        }
        private void UpdateCommitCompetition(ShortCompetition comp)
        {
            int ind = LstCompetitions.FindIndex(x => comp.Id == x.Id);
            if (ind > -1)
            {
                LstCompetitions[ind].HasCompetition = true;
                LstCompetitions[ind].IsApprovedByComission = true;
                LstCompetitions[ind].CompetitionId = comp.CompetitionId;
                LstCompetitions[ind].CompetitionName = comp.CompetitionName;

                LstCompetitions[ind].DocInsertDate = comp.DocInsertDate;
                LstCompetitions[ind].IsGosLine = comp.IsGosLine;
                LstCompetitions[ind].IsListener = comp.IsListener;
                LstCompetitions[ind].IsReduced = comp.IsReduced;

                LstCompetitions[ind].FacultyId = comp.FacultyId;
                LstCompetitions[ind].FacultyName = comp.FacultyName;
                LstCompetitions[ind].LicenseProgramId = comp.LicenseProgramId;
                LstCompetitions[ind].LicenseProgramName = comp.LicenseProgramName;
                LstCompetitions[ind].ObrazProgramId = comp.ObrazProgramId;
                LstCompetitions[ind].ObrazProgramName = comp.ObrazProgramName;
                LstCompetitions[ind].ProfileId = comp.ProfileId;
                LstCompetitions[ind].ProfileName = comp.ProfileName;

                LstCompetitions[ind].StudyFormId = comp.StudyFormId;
                LstCompetitions[ind].StudyFormName = comp.StudyFormName;
                LstCompetitions[ind].StudyBasisId = comp.StudyBasisId;
                LstCompetitions[ind].StudyBasisName = comp.StudyBasisName;
                LstCompetitions[ind].StudyLevelId = comp.StudyLevelId;
                LstCompetitions[ind].StudyLevelName = comp.StudyLevelName;

                LstCompetitions[ind].HasCompetition = comp.HasCompetition;
                LstCompetitions[ind].ChangeEntry();

                string userName = MainClass.GetUserName();

                string query = "UPDATE [Application] SET IsApprovedByComission=1, ApproverName=@ApproverName, CompetitionId=@CompId, DocInsertDate=@DocInsertDate, IsCommonRussianCompetition=@IsCommonRussianCompetition, IsGosLine=@IsGosLine WHERE Id=@Id";
                _bdcInet.ExecuteQuery(query, new SortedList<string, object>() { { "@Id", comp.Id }, { "@CompId", comp.CompetitionId }, { "@DocInsertDate", comp.DocInsertDate }, 
                { "@ApproverName", userName }, 
                { "@IsGosLine", comp.IsGosLine },
                { "@IsCommonRussianCompetition", comp.IsCommonRussianCompetition }
                });

                UpdateApplicationGrid();
            }
        }
    }
}
