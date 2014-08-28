using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Transactions;

using EducServLib;
using BDClassLib;
using WordOut;
using BaseFormsLib;
using PriemLib;
namespace Priem
{
    public partial class Olymp2Competition : BookList
    {
        //конструктор
        public Olymp2Competition()
        {
            InitializeComponent();
            Dgv = dgvAbitList;
            _title = "Зачет результатов олимпиад как изменение типа конкурса на б/э";

            InitControls();            
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            Dgv.ReadOnly = false;
            btnRemove.Visible = btnAdd.Visible = false;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.qFaculty", "ORDER BY Acronym"), false, false);
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, true);

                    cbStudyBasis.SelectedIndex = 0;
                    FillStudyForm();
                    FillLicenseProgram();
                    FillObrazProgram();
                   

                    ComboServ.FillCombo(cbOlympType, HelpClass.GetComboListByTable("ed.OlympType", " ORDER BY Id"), false, true);
                    ComboServ.FillCombo(cbOlympName, HelpClass.GetComboListByTable("ed.OlympName"), false, true);
                    ComboServ.FillCombo(cbOlympSubject, HelpClass.GetComboListByTable("ed.OlympSubject"), false, true);
                    ComboServ.FillCombo(cbOlympValue, HelpClass.GetComboListByTable("ed.OlympValue"), false, true);
                    ComboServ.FillCombo(cbOlympLevel, HelpClass.GetComboListByTable("ed.OlympLevel"), false, true);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            }
        }

        #region Handlers

        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }

        public int? LicenseProgramId
        {
            get { return ComboServ.GetComboIdInt(cbLicenseProgram); }
            set { ComboServ.SetComboId(cbLicenseProgram, value); }
        }

        public int? ObrazProgramId
        {
            get { return ComboServ.GetComboIdInt(cbObrazProgram); }
            set { ComboServ.SetComboId(cbObrazProgram, value); }
        }

        //public Guid? ProfileId
        //{
        //    get
        //    {
        //        string prId = ComboServ.GetComboId(cbProfile);
        //        if (string.IsNullOrEmpty(prId))
        //            return null;
        //        else
        //            return new Guid(prId);
        //    }
        //    set
        //    {
        //        if (value == null)
        //            ComboServ.SetComboId(cbProfile, (string)null);
        //        else
        //            ComboServ.SetComboId(cbProfile, value.ToString());
        //    }
        //}

        public int? StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }

        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }        

        public int? OlympTypeId
        {
            get { return ComboServ.GetComboIdInt(cbOlympType); }
            set { ComboServ.SetComboId(cbOlympType, value); }
        }

        public int? OlympSubjectId
        {
            get { return ComboServ.GetComboIdInt(cbOlympSubject); }
            set { ComboServ.SetComboId(cbOlympSubject, value); }
        }

        public int? OlympLevelId
        {
            get { return ComboServ.GetComboIdInt(cbOlympLevel); }
            set { ComboServ.SetComboId(cbOlympLevel, value); }
        }

        public int? OlympValueId
        {
            get { return ComboServ.GetComboIdInt(cbOlympValue); }
            set { ComboServ.SetComboId(cbOlympValue, value); }
        }

        public int? OlympNameId
        {
            get { return ComboServ.GetComboIdInt(cbOlympName); }
            set { ComboServ.SetComboId(cbOlympName, value); }
        }

        private void FillStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.StudyFormId.ToString(), u.StudyFormName)).Distinct().ToList();

                ComboServ.FillCombo(cbStudyForm, lst, false, true);
                cbStudyForm.SelectedIndex = 0;
            }
        }

        private void FillLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.LicenseProgramId.ToString(), u.LicenseProgramName)).Distinct().ToList();

                ComboServ.FillCombo(cbLicenseProgram, lst, false, true);
                cbLicenseProgram.SelectedIndex = 0;
            }
        }

        private void FillObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);
                if (LicenseProgramId != null)
                    ent = ent.Where(c => c.LicenseProgramId == LicenseProgramId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ObrazProgramId.ToString(), u.ObrazProgramName + ' ' + u.ObrazProgramCrypt)).Distinct().ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, true);
            }
        }

        //инициализация обработчиков мегакомбов
        public override void InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);
            cbObrazProgram.SelectedIndexChanged += new EventHandler(cbObrazProgram_SelectedIndexChanged);           
            cbStudyForm.SelectedIndexChanged += new EventHandler(cbStudyForm_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);            

            cbOlympType.SelectedIndexChanged += new EventHandler(cbOlympType_SelectedIndexChanged);
            cbOlympName.SelectedIndexChanged += new EventHandler(cbOlympName_SelectedIndexChanged);
            cbOlympSubject.SelectedIndexChanged += new EventHandler(cbOlympSubject_SelectedIndexChanged);
            cbOlympValue.SelectedIndexChanged += new EventHandler(cbOlympValue_SelectedIndexChanged);
            cbOlympLevel.SelectedIndexChanged += new EventHandler(cbOlympLevel_SelectedIndexChanged);
        }

        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillStudyForm();
        }

        void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
        }

        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillObrazProgram();
        }

        void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        void cbExam_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        void cbOlympLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        void cbOlympValue_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        void cbOlympSubject_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        void cbOlympName_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        void cbOlympType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        #endregion
                             
        //обновление грида
        public override void UpdateDataGrid()
        {
            DataTable examTable = new DataTable();
            DataSet ds;

            DataColumn clm;
            clm = new DataColumn();
            clm.ColumnName = "Id";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "Рег_номер";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "ФИО";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "Тип_конкурса";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "Вид";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "Название";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "Предмет";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "Уровень";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "Степень";
            clm.ReadOnly = true;
            examTable.Columns.Add(clm);            

            clm = new DataColumn();
            clm.ColumnName = "Изменить тип конкурса";            
            clm.Caption = "Set";
            clm.ReadOnly = false;
            clm.DataType = typeof(bool);
            examTable.Columns.Add(clm);          

            ds = _bdc.GetDataSet(GetQuery());

            lblCount.Text = ds.Tables[0].Rows.Count.ToString();

            foreach (DataRow dsRow in ds.Tables[0].Rows)
            {
                DataRow newRow;
                newRow = examTable.NewRow();
                newRow["Рег_номер"] = dsRow["Рег_Номер"].ToString();
                newRow["ФИО"] = dsRow["ФИО"].ToString();
                newRow["Тип_конкурса"] = dsRow["Тип_конкурса"].ToString();
                newRow["Вид"] = dsRow["Вид"].ToString();
                newRow["Название"] = dsRow["Название"].ToString();
                newRow["Предмет"] = dsRow["Предмет"].ToString();
                newRow["Уровень"] = dsRow["Уровень"].ToString();                
                newRow["Степень"] = dsRow["Степень"].ToString();                
                newRow["Id"] = dsRow["Id"].ToString();
                
                examTable.Rows.Add(newRow);
            }
            
            foreach (DataColumn dc in examTable.Columns)
                if (dc.Caption == "Set")
                    dc.ReadOnly = false;
                else
                    dc.ReadOnly = true;

            DataView dv = new DataView(examTable);
            dv.AllowNew = false;

            string sortedColumn = string.Empty;
            ListSortDirection order = ListSortDirection.Ascending;
            bool sorted = false;
            int index = 0;

            if (dgvAbitList.SortOrder != SortOrder.None)
            {
                sorted = true;
                sortedColumn = dgvAbitList.SortedColumn.Name;
                order = dgvAbitList.SortOrder == SortOrder.Ascending ? ListSortDirection.Ascending : ListSortDirection.Descending;
                index = dgvAbitList.CurrentRow == null ? -1 : dgvAbitList.CurrentRow.Index;
            }

            dgvAbitList.DataSource = dv;
            dgvAbitList.Columns["Id"].Visible = false;
            dgvAbitList.Update();

            if (dgvAbitList.Rows.Count > 0)
            {
                if (sorted && dgvAbitList.Columns.Contains(sortedColumn))
                    dgvAbitList.Sort(dgvAbitList.Columns[sortedColumn], order);
                if (index >= 0 && index <= dgvAbitList.Rows.Count)
                    dgvAbitList.CurrentCell = dgvAbitList[1, index];
            }
        }

        //сбор фильтров 
        private string GetQuery()
        {
            string sQuery = @"SELECT ed.qAbiturient.Id as Id, ed.qAbiturient.RegNum as Рег_Номер, ed.extPersonAspirant.FIO as ФИО, 
                        ed.Competition.Name as Тип_конкурса, OlympTypeName as Вид, OlympName AS Название, OlympLevelName AS Уровень, OlympSubjectName as Предмет,
                        OlympValueName as Степень
                        FROM ed.qAbiturient LEFT JOIN ed.extPersonAspirant ON ed.qAbiturient.PersonId = ed.extPersonAspirant.Id INNER JOIN ed.extOlympiads ON ed.extOlympiads.AbiturientId = ed.qAbiturient.Id 
                        LEFT JOIN ed.Competition ON ed.Competition.Id = ed.qAbiturient.CompetitionId ";

            sQuery += " WHERE ed.qAbiturient.BackDoc=0 AND ed.qAbiturient.CompetitionId <> 1 ";

            string s1 = string.Empty;

            s1 += " AND ed.qAbiturient.StudyLevelGroupId = " + MainClass.studyLevelGroupId;

            //обработали форму обучения  
            if (StudyFormId != null)
                s1 += " AND ed.qAbiturient.StudyFormId = " + StudyFormId;

            //обработали основу обучения  
            if (StudyBasisId != null)
                s1 += " AND ed.qAbiturient.StudyBasisId = " + StudyBasisId;

            //обработали факультет
            if (FacultyId != null)
                s1 += " AND ed.qAbiturient.FacultyId = " + FacultyId;

            //обработали Направление
            if (LicenseProgramId != null)
                s1 += " AND ed.qAbiturient.LicenseProgramId = " + LicenseProgramId;

            //обработали Образ программу
            if (ObrazProgramId != null)
                s1 += " AND ed.qAbiturient.ObrazProgramId = " + ObrazProgramId;  

            sQuery += s1;
            
            string s2 = string.Empty;

            //обработали вид            
            if (OlympTypeId != null)
                s2 += " AND ed.extOlympiads.OlympTypeId = " + OlympTypeId;

            //обработали уровень            
            if (OlympLevelId != null)
                s2 += " AND ed.extOlympiads.OlympLevelId = " + OlympLevelId;

            //обработали предмет            
            if (OlympSubjectId != null)
                s2 += " AND ed.extOlympiads.OlympSubjectId = " + OlympSubjectId;

            //обработали значение            
            if (OlympValueId != null)
                s2 += " AND ed.extOlympiads.OlympValueId  = " + OlympValueId;

            //обработали название            
            if (OlympNameId != null)
                s2 += " AND ed.extOlympiads.OlympNameId  = " + OlympNameId;    

            sQuery += s2;

            return sQuery;
        }
        
        //поиск по фио
        private void tbFIO_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbitList, "ФИО", tbFIO.Text);
        }         

        private void chbAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dgvr in dgvAbitList.Rows)
            {
                if (chbAll.Checked)
                    dgvr.Cells["Изменить тип конкурса"].Value = true;
                else
                    dgvr.Cells["Изменить тип конкурса"].Value = false;
            }     
        }

        protected override void OpenCard(string id, BaseFormEx formOwner, int? index)
        {
            MainClassCards.OpenCardAbit(id, this, dgvAbitList.CurrentRow.Index);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Подтвердить зачет олимпиад у выбранных заявлений как тип конкурса б/э? (необходимо будет создать новый протокол о допуске для всех выбранных абитуриентов!)", "Зачет олимпиад", MessageBoxButtons.YesNoCancel);
            if (res == DialogResult.Yes)
            {              
                try
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                        {
                            foreach (DataGridViewRow row in dgvAbitList.Rows)
                            {
                                if (row.Cells["Изменить тип конкурса"].Value.ToString() == "True")
                                {
                                    Guid abId = new Guid(row.Cells["Id"].Value.ToString());

                                    context.Abiturient_UpdateCompetititon(1, null, true, abId);
                                    context.Abiturient_UpdateChecked(false, abId);

                                    //исключаем из предыдущего протокола о допуске
                                    Guid? protId = (from pr in context.qProtocolHistory
                                                    where pr.AbiturientId == abId && !pr.Excluded && !pr.IsOld
                                                    select pr.ProtocolId).FirstOrDefault();

                                    if (protId != null)
                                        context.ProtocolHistory_UpdateExcluded(abId, protId, true, null);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    WinFormsServ.Error("Ошибка при сохранении перезачета оценок. Оценки перезачтены не будут. " + ex.Message);
                }

                UpdateDataGrid();
            }
        }

        //печать
        private void btnPrint_Click(object sender, EventArgs e)
        {            
            string s1 = string.Empty;

            //обработали факультет
            if (FacultyId != null)
                s1 += " AND ed.extAbitAspirant.FacultyId = " + FacultyId;

            //обработали Направление
            if (LicenseProgramId != null)
                s1 += " AND ed.extAbitAspirant.LicenseProgramId = " + LicenseProgramId;

            //обработали Образ программу
            if (ObrazProgramId != null)
                s1 += " AND ed.extAbitAspirant.ObrazProgramId = " + ObrazProgramId;
            
            string query = string.Format(@"SELECT ed.extPersonAspirant.PersonNum AS PerNum, ed.extAbitAspirant.RegNum,  LicenseProgramName, ObrazProgramName, 
                    StudyFormName, ed.extAbitAspirant.FIO as fio, ed.Competition.Name as Competition
                    FROM ed.extAbitAspirant INNER JOIN ed.extPersonAspirant ON ed.extAbitAspirant.PersonId = ed.extPersonAspirant.Id 
                    LEFT JOIN ed.Competition ON ed.extAbitAspirant.CompetitionId = ed.Competition.Id
                    WHERE extAbitAspirant.CompFromOlymp = 1 AND extAbitAspirant.CompetitionId = 1 {0} ORDER BY 1", s1);           
            try
            {
                DataSet ds = _bdc.GetDataSet(query);
                if (ds.Tables[0].Rows.Count == 0)
                    return;

                WordDoc wd = new WordDoc(string.Format(@"{0}\Olymp2Competition.dot", MainClass.dirTemplates), true);
                TableDoc td = wd.Tables[0];
                                
                wd.Fields["Факультет"].Text = cbFaculty.Text;
                wd.Fields["Специальность"].Text = cbLicenseProgram.Text;
                wd.Fields["ОбразовательнаяПрограмма"].Text = cbObrazProgram.Text;  

                int i = 1;
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    td.AddRow(1);
                    td[0, i] = i.ToString();
                    td[1, i] = dr["PerNum"].ToString();
                    td[2, i] = dr["RegNum"].ToString();
                    td[3, i] = dr["LicenseProgramName"].ToString();
                    td[4, i] = dr["ObrazProgramName"].ToString();
                    td[5, i] = dr["StudyFormName"].ToString();
                    td[6, i] = dr["fio"].ToString();
                    td[7, i] = dr["Competition"].ToString();
                    i++;
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при выводе в Word: " + ex.Message);
            }            
        } 
    }
}