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
    public partial class ApplicationInetList : BookList
    {
        protected DBPriem _bdcInet;       
        private LoadFromInet loadClass;
        
        //конструктор
        public ApplicationInetList()
        {
            InitializeComponent();
            Dgv = dgvAbiturients;

            _title = "Список online заявлений для загрузки";
                        
            InitControls();
        }

        public DBPriem BdcInet
        {
            get { return _bdcInet; }
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();

            btnCard.Visible = btnAdd.Visible = btnRemove.Visible = false;
            
            loadClass = new LoadFromInet();
            _bdcInet = loadClass.BDCInet;

            if (MainClass.IsReadOnly())
                btnLoad.Enabled = false;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.qFaculty", "ORDER BY Acronym"), false, true);
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, true);

                    cbStudyBasis.SelectedIndex = 0;
                    FillLicenseProgram();
                    FillObrazProgram();
                    FillProfile();

                    UpdateDataGrid();

                    if (MainClass.IsPasha())
                        gbUpdateImport.Visible = true;
                    else
                        gbUpdateImport.Visible = false;

                    chbSelectAll.Checked = false;

                    tbAbitBarcode.Focus();
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            } 
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

                ComboServ.FillCombo(cbLicenseProgram, lst, false, true);
                cbLicenseProgram.SelectedIndex = 0;
            }
        }
        private void FillObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = ((from ent in MainClass.GetEntry(context)
                                                           where ent.FacultyId == FacultyId
                                                           && (LicenseProgramId.HasValue ? ent.LicenseProgramId == LicenseProgramId : true)
                                                           select new
                                                           {
                                                               Id = ent.ObrazProgramId,
                                                               Name = ent.ObrazProgramName,
                                                               Crypt = ent.ObrazProgramCrypt
                                                           }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name + ' ' + u.Crypt)).ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, true);
            }
        }
        private void FillProfile()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = ((from ent in MainClass.GetEntry(context)
                                                           where ent.FacultyId == FacultyId
                                                           && (LicenseProgramId.HasValue ? ent.LicenseProgramId == LicenseProgramId : true)
                                                           && (ObrazProgramId.HasValue ? ent.ObrazProgramId == ObrazProgramId : true)
                                                           && ent.ProfileId != null
                                                           select new
                                                           {
                                                               Id = ent.ProfileId,
                                                               Name = ent.ProfileName
                                                           }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                if (lst.Count() > 0)
                {
                    ComboServ.FillCombo(cbProfile, lst, false, true);
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
        public override void  InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);
            cbObrazProgram.SelectedIndexChanged += new EventHandler(cbObrazProgram_SelectedIndexChanged);
            cbProfile.SelectedIndexChanged += new EventHandler(cbProfile_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);            
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
            //UpdateDataGrid();
        }
        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillObrazProgram();
            //UpdateDataGrid();
        }
        void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {            
            FillProfile();
            UpdateDataGrid();
        }        

        //строим запрос фильтров
        private string GetFilterString()
        {
            string s = "AND qAbiturient.StudyLevelId = 15 ";

            //обработали основу обучения 
            if (StudyBasisId != null)
                s += string.Format(" AND qAbiturient.StudyBasisId = {0}", StudyBasisId);

            //обработали факультет            
            if (FacultyId != null)
                s += string.Format(" AND qAbiturient.FacultyId = {0}", FacultyId);         

            //обработали Направление
            if (LicenseProgramId != null)
                s += string.Format(" AND qAbiturient.LicenseProgramId = {0}", LicenseProgramId); 

            //обработали Образ программу
            if (ObrazProgramId != null)
                s += string.Format(" AND qAbiturient.ObrazProgramId = {0}", ObrazProgramId); 

            //обработали специализацию 
            if (ProfileId != null)
                s += string.Format(" AND qAbiturient.ProfileId = '{0}'", ProfileId); 

            return s;
        }       

        //обновление грида
        protected override void GetSource()
        {
            _orderBy =" ORDER BY Приложены_файлы, Дата_обновления DESC, ФИО";
            _sQuery = @"SELECT DISTINCT qAbiturient.CommitId AS Id, extPerson.Surname + ' ' + extPerson.Name + ' ' + extPerson.SecondName as ФИО, 
                       extPerson.BirthDate AS Дата_рождения, qAbiturient.CommitNumber AS Barcode, 
                       (Case When EXISTS(SELECT extAbitFiles.Id FROM extAbitFiles WHERE extAbitFiles.PersonId = extPerson.Id) then 'да' else 'нет' end) AS Приложены_файлы,
                       (SELECT Max(extAbitFiles.LoadDate) FROM extAbitFiles WHERE extAbitFiles.PersonId = extPerson.Id AND (extAbitFiles.ApplicationId = qAbiturient.Id OR extAbitFiles.ApplicationId IS NULL)) AS Дата_обновления                     
                       FROM qAbiturient INNER JOIN extPerson ON qAbiturient.PersonId = extPerson.Id
                       WHERE qAbiturient.IsImported = 0 AND Enabled = 1 AND qAbiturient.Id NOT IN (SELECT Id FROM qForeignApplicationOnly)";

            HelpClass.FillDataGrid(dgvAbiturients, _bdcInet, _sQuery + GetFilterString(), _orderBy);
            dgvAbiturients.Columns["Barcode"].Visible = false;
            btnLoad.Enabled = !(dgvAbiturients.RowCount == 0);
        }

        //поле поиска
        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbiturients, "ФИО", tbSearch.Text);
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            try
            {   
                string barcText = tbAbitBarcode.Text.Trim();

                if (barcText.Length > 0)
                {
                    if (barcText.Length == 7)
                    {
                        if (barcText.StartsWith("1"))
                        {
                            WinFormsServ.Error("Выбран человек, подавший заявления на первый курс");
                            return;
                        }
                        barcText = barcText.Substring(1);
                    }

                    int code;
                    if (!int.TryParse(barcText, out code))
                    {
                        WinFormsServ.Error("Не распознан баркод!");
                        return;
                    }
                    if (code == 0)
                    {
                        WinFormsServ.Error("Не распознан баркод!");
                        return;
                    }

                    using (PriemEntities context = new PriemEntities())
                    {
                        int cnt = (from ab in context.Abiturient
                                   where ab.CommitNumber == code
                                   select ab).Count();

                        if (cnt > 0)
                        {
                            WinFormsServ.Error("Запись уже добавлена!");
                            return;
                        }

                        bool t = false;
                        for (int i = 0; i < dgvAbiturients.Rows.Count; i++)
                        {
                            if (dgvAbiturients.Rows[i].Cells["Barcode"].Value.ToString() == code.ToString())
                            {
                                dgvAbiturients.CurrentCell = dgvAbiturients[1, i];
                                t = true;
                                break;
                            }
                        }

                        if (!t)
                        {
                            //поиск заявления в зоне видимости приёмной комиссии
                            var entries = context.qEntry.Select(x => x.Id).ToList();
                            string lstEntries = Util.BuildStringWithCollection(entries.Select(x => "'" + x.ToString() + "'").ToList());

                            string query = string.Format("SELECT Barcode, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId, StudyBasisId FROM qAbiturient WHERE Barcode=@Barcode AND EntryId IN ({0})", lstEntries);
                            DataTable tbl = BdcInet.GetDataSet(query, new SortedList<string, object>() { { "@Barcode", code } }).Tables[0];

                            if (tbl.Rows.Count == 0)
                            {
                                WinFormsServ.Error("Запись не найдена!");
                                return;
                            }

                            //StudyBasisId = tbl.Rows[0].Field<int?>("StudyBasisId");
                            //FacultyId = tbl.Rows[0].Field<int?>("FacultyId");
                            //LicenseProgramId = tbl.Rows[0].Field<int?>("LicenseProgramId");
                            //ObrazProgramId = tbl.Rows[0].Field<int?>("ObrazProgramId");
                            //if (!tbl.Rows[0].Field<Guid?>("ProfileId").HasValue)
                            //    ProfileId = tbl.Rows[0].Field<Guid?>("ProfileId");
                        }

                        tbAbitBarcode.Text = string.Empty;

                        CardFromInet crd = new CardFromInet(null, code, false);
                        crd.ToUpdateList += new UpdateListHandler(UpdateDataGrid);
                        crd.Show();
                    }
                }
            }            
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            UpdateDataGrid();            
        }

        protected override void OnClosed()
        {
            loadClass.CloseDB();
        }

        private void dgvAbiturients_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == dgvAbiturients.Columns["ФИО"].Index && dgvAbiturients["Приложены_файлы", e.RowIndex].Value.ToString() == "да")
            {
                dgvAbiturients["ФИО", e.RowIndex].Style.BackColor = Color.LightGreen;
            }
        }

        private void btnUpdateImport_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsPasha())
                return;            

            if (MessageBox.Show("Обновить IsImported = true?", "Обновление", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (DataGridViewRow dgvr in dgvAbiturients.SelectedRows)
                {
                    string sId = dgvr.Cells["Id"].Value.ToString();
                    try
                    {
                        string query; 
                        query = string.Format("UPDATE [Application] SET IsImported = 1 WHERE Id = '{0}'",  sId);                        

                        _bdcInet.ExecuteQuery(query);
                    }
                    catch (Exception ex)
                    {
                        WinFormsServ.Error("Ошибка обновления данных" + ex.Message);
                        goto Next;
                    }
                Next: ;
                }
                UpdateDataGrid();
            }
        }

        private void chbSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            dgvAbiturients.MultiSelect = chbSelectAll.Checked;
            btnLoad.Enabled = !chbSelectAll.Checked;
        }

        protected override void Dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Dgv.CurrentCell != null && Dgv.CurrentCell.RowIndex > -1)            
                tbAbitBarcode.Text = Dgv.Rows[Dgv.CurrentCell.RowIndex].Cells["Barcode"].Value.ToString();

            btnLoad_Click(null, null);
        }       
    }
}