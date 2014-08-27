using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using BDClassLib;
using BaseFormsLib;
using PriemLib;

namespace Priem
{
    public partial class OnlineChangesList : BookList
    {        
        protected DBPriem _bdcInet;      
        private LoadFromInet loadClass;

        private Dictionary<int, ChangeAbitClass> lstAbits;
        
        //конструктор
        public OnlineChangesList()
        {
            // ChangeType: 1- изменил приоритет, 2 - забрал документы в инет базе, 3 - забрал документы в нашей базе, 4 - добавил документы
            InitializeComponent();
            Dgv = dgvAbiturients;
            this.Text = "Изменения";                    
            
            InitControls();            
        }

        #region Fields
        public DBPriem BdcInet
        {
            get { return _bdcInet; }
        }
        public int? ChangeTypeId
        {
            get { return ComboServ.GetComboIdInt(cbChangeType); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
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
        public Guid? ProfileId
        {
            get
            {
                string prId = ComboServ.GetComboId(cbProfile);
                if (string.IsNullOrEmpty(prId))
                    return null;
                else
                    return new Guid(prId);
            }
            set
            {
                if (value == null)
                    ComboServ.SetComboId(cbProfile, (string)null);
                else
                    ComboServ.SetComboId(cbProfile, value.ToString());
            }
        }
        public int? StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }
        #endregion

        protected override void ExtraInit()
        {
            base.ExtraInit();
            
            loadClass = new LoadFromInet();
            _bdcInet = loadClass.BDCInet;

            if (MainClass.IsReadOnly())
                btnLoad.Enabled = false;

            btnAdd.Visible = btnCard.Visible = btnRemove.Visible = false;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {                    
                    ComboServ.FillCombo(cbChangeType, HelpClass.GetComboListByTable("ed.ChangeType", "ORDER BY Id"), false, false);
                    ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.qFaculty", "ORDER BY Acronym"), false, false);
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, true);

                    cbStudyBasis.SelectedIndex = 0;
                    FillLicenseProgram();
                    FillObrazProgram();
                    FillProfile();

                    UpdateDataGrid();
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            } 
        }

        private void FillLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = ((from ent in MainClass.GetEntry(context)
                                                           where ent.FacultyId == FacultyId
                                                           select new
                                                           {
                                                               Id = ent.LicenseProgramId,
                                                               Name = ent.LicenseProgramName
                                                           }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                ComboServ.FillCombo(cbLicenseProgram, lst, false, false);
                cbLicenseProgram.SelectedIndex = 0;
            }
        }
        private void FillObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = ((from ent in MainClass.GetEntry(context)
                                                           where ent.FacultyId == FacultyId && ent.LicenseProgramId == LicenseProgramId
                                                           select new
                                                           {
                                                               Id = ent.ObrazProgramId,
                                                               Name = ent.ObrazProgramName,
                                                               Crypt = ent.ObrazProgramCrypt
                                                           }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name + ' ' + u.Crypt)).ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, false);
            }
        }
        private void FillProfile()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = ((from ent in MainClass.GetEntry(context)
                                                           where ent.FacultyId == FacultyId && ent.LicenseProgramId == LicenseProgramId && ent.ObrazProgramId == ObrazProgramId && ent.ProfileId != null
                                                           select new
                                                           {
                                                               Id = ent.ProfileId,
                                                               Name = ent.ProfileName
                                                           }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                if (lst.Count() > 0)
                {
                    ComboServ.FillCombo(cbProfile, lst, false, false);
                    cbProfile.Enabled = true;
                }
                else
                {
                    ComboServ.FillCombo(cbProfile, new List<KeyValuePair<string, string>>(), true, false);
                    cbProfile.Enabled = false;
                }
            }
        }

        //инициализация обработчиков мегакомбов
        public override void InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);
            cbObrazProgram.SelectedIndexChanged += new EventHandler(cbObrazProgram_SelectedIndexChanged);
            cbProfile.SelectedIndexChanged += new EventHandler(cbProfile_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
            cbChangeType.SelectedIndexChanged += new EventHandler(cbChangeType_SelectedIndexChanged);
        }

        void cbChangeType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateAfterChangeType();
            UpdateDataGrid();
        }
        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }
        void cbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }
        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
            UpdateDataGrid();
        }
        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillObrazProgram();
            UpdateDataGrid();
        }
        void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillProfile();
            UpdateDataGrid();
        }      

        private void UpdateAfterChangeType()
        {            
            switch (ChangeTypeId)
            {
                // приоритеты
                case 1:
                    {
                        btnLoad.Text = "Загрузить обновленные данные";                        
                        break;
                    }
                // забрал документы в инете
                case 2:
                    {
                        btnLoad.Text = "Загрузить обновленные данные";                        
                        break;
                    }
                // забрал документы в нашей базе    
                case 3:
                    {
                        btnLoad.Text = "Обновить информацию в онлайн базе";                        
                        break;
                    }
                // новые документы    
                case 4:
                    {
                        btnLoad.Text = "Просмотреть документы";
                        break;
                    }
            }
        }

        //строим запрос фильтров
        private string GetFilterString(string tableName)
        {
            string s = string.Empty;

            //обработали основу обучения 
            if (StudyBasisId != null)
                s += string.Format(" AND {0}.StudyBasisId = {1}", tableName, StudyBasisId);

            //обработали факультет            
            if (FacultyId != null)
                s += string.Format(" AND {0}.FacultyId = {1}", tableName, FacultyId);

            //обработали Направление
            if (LicenseProgramId != null)
                s += string.Format(" AND {0}.LicenseProgramId = {1}", tableName, LicenseProgramId);

            //обработали Образ программу
            if (ObrazProgramId != null)
                s += string.Format(" AND {0}.ObrazProgramId = {1}", tableName, ObrazProgramId);

            //обработали специализацию 
            if (ProfileId != null)
                s += string.Format(" AND {0}.ProfileId = '{1}'", tableName, ProfileId);

            return s;
        }  
       
        //обновление грида
        public override void UpdateDataGrid()
        {
            if (ChangeTypeId == 4)
                UpdateDataGridPerson();
            else
                UpdateDataGridAbits();
        }

        private void UpdateDataGridPerson()
        {
            string query = @"SELECT DISTINCT qAbiturient.Id, qAbiturient.Barcode FROM qAbiturient WHERE 
(qAbiturient.Id IN 
(
SELECT DISTINCT ab.Id FROM extAbitFiles INNER JOIN qAbiturient AS ab ON extAbitFiles.ApplicationId = ab.Id 
WHERE (ab.DateReviewDocs IS NULL OR DateDiff(MINUTE, ab.DateReviewDocs, extAbitFiles.LoadDate) > 0)
)
OR
qAbiturient.PersonId IN 
(
SELECT DISTINCT pers.Id FROM extAbitFiles INNER JOIN Person AS pers ON extAbitFiles.PersonId = pers.Id 
WHERE extAbitFiles.ApplicationId IS NULL AND (pers.DateReviewDocs IS NULL OR DateDiff(MINUTE, pers.DateReviewDocs, extAbitFiles.LoadDate) > 0)
))
AND qAbiturient.IsImported = 1 "; 
            
            DataSet dsAbits = _bdcInet.GetDataSet(query);

            List<string> lstAbits = (from dr in dsAbits.Tables[0].AsEnumerable()
                      select dr["Barcode"].ToString()).ToList<string>();
       
            if (lstAbits.Count == 0)
            {
                SetNullGrid();
                return;
            }

            string s = string.Join(", ", lstAbits.ToArray());
            
            string queryOur = string.Format(
                   @"SELECT extAbitAspirant.Id, extAbitAspirant.Barcode, Person.Surname + ' ' + Person.Name + ' ' + Person.SecondName as ФИО, Person.BirthDate AS Дата_рождения,
                   extAbitAspirant.FacultyName AS Факультет, extAbitAspirant.LicenseProgramName AS Направление, extAbitAspirant.ObrazProgramName AS Образ_программа, 
                   extAbitAspirant.ProfileName AS Профиль, extAbitAspirant.StudyBasisName AS Основа, Person.Barcode AS PersonBarcode                  
                   FROM ed.extAbitAspirant INNER JOIN ed.Person ON extAbitAspirant.PersonId = Person.Id 
                   WHERE 1=1 {0} AND extAbitAspirant.Barcode IN({1}) ORDER BY ФИО", GetFilterString("extAbitAspirant"), s);

            HelpClass.FillDataGrid(dgvAbiturients, _bdc, queryOur, "");
            dgvAbiturients.Columns["Barcode"].Visible = false;
            dgvAbiturients.Columns["PersonBarcode"].Visible = false;            
            lblCount.Text = "Всего: " + dgvAbiturients.RowCount.ToString();           

            btnLoad.Enabled = !(dgvAbiturients.RowCount == 0);           
        }

        //обновление грида
        public void UpdateDataGridAbits()
        {
            try
            {
                int? changeTp = ChangeTypeId;

                string abitQuery = string.Format("SELECT Barcode, (Case When BackDoc = 1 then 1 else 0 end) AS BackDoc, Cast(Priority AS nvarchar(10)) AS Priority FROM ed.extAbitAspirant WHERE Barcode > 0  {0} ", GetFilterString("ed.extAbitAspirant"));
                string abitQueryInet = string.Format("SELECT qAbiturient.Barcode, Cast(qAbiturient.Priority AS nvarchar(10)) AS Priority, (Case When qAbiturient.Enabled = 1 then 0 else 1 end) AS BackDoc FROM qAbiturient WHERE  IsImported = 1 AND Barcode > 0 {0} ", GetFilterString("qAbiturient"));
                                
                Dictionary<int, ChangeAbitClass> abits;
                Dictionary<int, ChangeAbitClass> abitsInet;
                
                lstAbits = new Dictionary<int, ChangeAbitClass>();
                
                switch (changeTp)
                {
                    // приоритеты
                    case 1:
                        {
                            abitQuery += " AND BackDoc = 0 ";
                            abitQueryInet += " AND Enabled = 1 AND NOT Priority IS NULL ";                            

                            abits = GetAbitList(_bdc.GetDataSet(abitQuery));
                            abitsInet = GetAbitList(_bdcInet.GetDataSet(abitQueryInet));

                            if (abitsInet.Count == 0)
                            {
                                SetNullGrid();
                                return;
                            }

                            foreach (ChangeAbitClass cl in abits.Values)
                            {
                                if (!abitsInet.Keys.Contains(cl.Barcode))
                                    continue;
                                
                                ChangeAbitClass inAb = abitsInet[cl.Barcode];

                                if (cl.Priority != inAb.Priority)
                                    lstAbits.Add(cl.Barcode, new ChangeAbitClass(cl.Barcode, inAb.Priority, cl.BackDoc));
                            }

                            break;
                        }
                    // забрал документы в инете
                    case 2:
                        {
                            abitQuery += " AND BackDoc = 0 ";
                            abitQueryInet += " AND Enabled = 0 ";

                            abits = GetAbitList(_bdc.GetDataSet(abitQuery));
                            abitsInet = GetAbitList(_bdcInet.GetDataSet(abitQueryInet));

                            if (abitsInet.Count == 0)
                            {
                                SetNullGrid();
                                return;
                            }

                            foreach (ChangeAbitClass cl in abits.Values)
                            {
                                if (!abitsInet.Keys.Contains(cl.Barcode))
                                    continue;

                                ChangeAbitClass inAb = abitsInet[cl.Barcode];

                                if (!cl.BackDoc && inAb.BackDoc)
                                    lstAbits.Add(cl.Barcode, new ChangeAbitClass(cl.Barcode, cl.Priority, inAb.BackDoc));
                            }

                            break;
                        }
                    // забрал документы в нашей базе    
                    case 3:
                        {
                            abitQuery += " AND BackDoc = 1 ";
                            abitQueryInet += " AND Enabled = 1 ";

                            abits = GetAbitList(_bdc.GetDataSet(abitQuery));
                            abitsInet = GetAbitList(_bdcInet.GetDataSet(abitQueryInet));

                            if (abitsInet.Count == 0)
                            {
                                SetNullGrid();
                                return;
                            }

                            foreach (ChangeAbitClass cl in abits.Values)
                            {
                                if (!abitsInet.Keys.Contains(cl.Barcode))
                                continue;

                                ChangeAbitClass inAb = abitsInet[cl.Barcode];

                                if (cl.BackDoc && !inAb.BackDoc)
                                    lstAbits.Add(cl.Barcode, new ChangeAbitClass(cl.Barcode, cl.Priority, cl.BackDoc));
                            }

                            break;
                        }                   
                }                              
                
                if(lstAbits.Count == 0)
                {
                    SetNullGrid();
                    return;
                }                   

                string query = string.Format(
                   @"SELECT ed.extAbitAspirant.Id, Person.Surname + ' ' + Person.Name + ' ' + Person.SecondName as ФИО, 
                   Person.BirthDate AS Дата_рождения, ed.extAbitAspirant.Barcode, Person.Barcode AS PersonBarcode,
                   ed.extAbitAspirant.FacultyName as Факультет, ed.extAbitAspirant.LicenseProgramName AS Направление, ed.extAbitAspirant.ObrazProgramName AS Образ_программа, 
                   ed.extAbitAspirant.ProfileName AS Профиль, ed.extAbitAspirant.StudyBasisName AS Основа 
                   FROM ed.extAbitAspirant INNER JOIN ed.Person ON ed.extAbitAspirant.PersonId = Person.Id 
                   WHERE 1=1 {0} AND ed.extAbitAspirant.Barcode IN({1}) ORDER BY ФИО", GetFilterString("ed.extAbitAspirant"), GetListKeys(lstAbits));


                HelpClass.FillDataGrid(dgvAbiturients, _bdc, query, "");
                dgvAbiturients.Columns["Barcode"].Visible = false;
                dgvAbiturients.Columns["PersonBarcode"].Visible = false;        
                lblCount.Text = "Всего: " + dgvAbiturients.RowCount.ToString();
            }
            catch(Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }

            btnLoad.Enabled = !(dgvAbiturients.RowCount == 0);           
        }

        private void SetNullGrid()
        {
            dgvAbiturients.DataSource = null;
            lblCount.Text = "Всего: 0";
            btnLoad.Enabled = false;                 
        }

        private Dictionary<int, ChangeAbitClass> GetAbitList(DataSet dsAbits)
        {
            Dictionary<int, ChangeAbitClass> abits = (from dr in dsAbits.Tables[0].AsEnumerable()
                                                      select new
                                                      {
                                                          key = dr.Field<int>("Barcode"),
                                                          val = new ChangeAbitClass(
                                                         dr.Field<int>("Barcode"),
                                                         dr.Field<string>("Priority"), 
                                                         QueryServ.ToBoolValue(dr.Field<int>("BackDoc"))
                                                         )
                                                      }).ToDictionary(n => n.key, n => n.val);
            return abits;
        }

        private string GetListKeys(Dictionary<int, ChangeAbitClass> lst)
        {
            List<string> t = lst.Select(kst=>kst.Key.ToString()).ToList();
            return string.Join(", ", t.ToArray());           
        }

        //поле поиска
        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbiturients, "ФИО", tbSearch.Text);
        }
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            UpdateDataGrid();            
        }
        private void btnClose_Click(object sender, EventArgs e)
        {            
            this.Close();
        }

        private void ApplicationInetList_FormClosing(object sender, FormClosingEventArgs e)
        {
            loadClass.CloseDB();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            int? changeTp = ChangeTypeId;

            foreach (DataGridViewRow dgvr in dgvAbiturients.SelectedRows)
            {
                string code = dgvr.Cells["Barcode"].Value.ToString();
                string perCode = dgvr.Cells["PersonBarcode"].Value.ToString();
                
                try
                {
                    int abitBarcode;
                    int persBarcode;

                    if (!int.TryParse(code, out abitBarcode))
                    {
                        WinFormsServ.Error("Не распознан баркод!");
                        return;
                    }

                    if (abitBarcode == 0)
                    {
                        WinFormsServ.Error("Не распознан баркод!");
                        return;
                    }

                    if (!int.TryParse(perCode, out persBarcode))
                    {
                        WinFormsServ.Error("Не распознан баркод!");
                        return;
                    }

                    if (persBarcode == 0)
                    {
                        WinFormsServ.Error("Не распознан баркод!");
                        return;
                    }

                    ChangeAbitClass chCl;

                    string updateQuery = string.Empty;

                    switch (changeTp)
                    {
                        // приоритеты
                        case 1:
                            {
                                if (!MessageAttention())
                                    return;

                                chCl = lstAbits[abitBarcode];

                                using (PriemEntities context = new PriemEntities())
                                {
                                    Guid abitId = (from ab in context.extAbitAspirant
                                                   where ab.Barcode == abitBarcode
                                                   select ab.Id).FirstOrDefault();
                                    double pr;
                                    if(!double.TryParse(chCl.Priority, out pr))
                                        break;

                                    context.Abiturient_UpdatePriority(pr, abitId);                                   
                                    break;
                                }
                            }
                        // забрал доки в инете
                        case 2:
                            {
                                if (!MessageAttention())
                                    return;

                                chCl = lstAbits[abitBarcode];
                                DataRow dr = _bdcInet.GetDataSet(string.Format("SELECT qAbiturient.DateOfDisable FROM qAbiturient WHERE Barcode = {0}", abitBarcode)).Tables[0].Rows[0];
                                DateTime? backDate = dr.Field<DateTime?>("DateOfDisable");

                                using (PriemEntities context = new PriemEntities())
                                {
                                    Guid abitId = (from ab in context.extAbitAspirant
                                                   where ab.Barcode == abitBarcode
                                                   select ab.Id).FirstOrDefault();

                                    context.Abiturient_UpdateBackDoc(true, backDate, abitId);
                                    break;
                                }
                            }
                        // забрал доки в нашей базе
                        case 3:
                            {
                                if (!MessageAttention())
                                    return;

                                chCl = lstAbits[abitBarcode];
                                DataRow dr = _bdc.GetDataSet(string.Format("SELECT qAbiturient.BackDocDate FROM ed.qAbiturient WHERE Barcode = {0}", abitBarcode)).Tables[0].Rows[0];
                                DateTime? backDate = dr.Field<DateTime?>("BackDocDate");
                                updateQuery = string.Format("UPDATE qAbiturient SET Enabled = 0, DateOfDisable = '{1}' WHERE Barcode = {0}", abitBarcode, backDate == null ? "NULL" : backDate.Value.Date.ToShortDateString());
                                _bdcInet.ExecuteQuery(updateQuery);
                                break;
                            }

                        // новые документы
                        case 4:
                            {
                                new DocCard(persBarcode, abitBarcode).ShowDialog();
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    WinFormsServ.Error("Ошибка обновления данных: " + code + ":" + ex.Message);
                    goto Next;
                }
            Next: ;
            }
            UpdateDataGrid();
        }

        private bool MessageAttention()
        {
            if (MessageBox.Show("Вы хотите обновить данные?", "Обновить", MessageBoxButtons.YesNo) == DialogResult.Yes)
                return true;
            else
                return false;
            
        }
    }
}