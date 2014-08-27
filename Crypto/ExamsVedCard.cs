using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Data.Objects;
using System.Transactions;

using EducServLib;
using BDClassLib;
using BaseFormsLib;
using PriemLib;

namespace Priem
{
    public partial class ExamsVedCard : BaseFormEx
    {
        private ExamsVedList owner;
        private DBPriem bdc; 
               
        private Guid? _Id;
        private bool isAdditional;
        private int? examId;   
        private int? facultyId;
        private int? studyBasisId;
        private DateTime passDate;
        private bool isAddVed;
        private int? addCount;

        private string sQuery = @"SELECT DISTINCT ed.extPersonAspirant.Id, ed.extPersonAspirant.PersonNum as Ид_Номер, ed.extPersonAspirant.FIO as ФИО, 
                                   ed.extPersonAspirant.EducDocument as Документ_об_Образовании, 
                                   ed.extPersonAspirant.PassportData as Паспорт, ed.qAbiturient.FacultyId 
                                   FROM ed.qAbiturient INNER JOIN ed.extPersonAspirant ON ed.qAbiturient.PersonId = ed.extPersonAspirant.Id 
                                   LEFT JOIN ed.extProtocol ON extProtocol.AbiturientId = qAbiturient.Id ";

        private string sOrderby = " ORDER BY ФИО ";

        public ExamsVedCard(ExamsVedList owner, Guid? vedId)
        {
            InitializeComponent();

            _Id = vedId;
            this.owner = owner;
            bdc = MainClass.Bdc;

            using (PriemEntities context = new PriemEntities())
            {
                extExamsVed ved = (from vd in context.extExamsVed
                                   where vd.Id == _Id
                                   select vd).FirstOrDefault();
                
                this.facultyId = ved.FacultyId;            
                this.passDate = ved.Date;
                this.isAdditional = ved.IsAdditional;
                this.examId = ved.ExamId;
                this.studyBasisId = ved.StudyBasisId;
                this.isAddVed = false;
                this.addCount = 0;
            }
            
            InitControls();       
        }

        public ExamsVedCard(ExamsVedList owner, int? facId, int? examId, DateTime date, int? basisId, bool isAdd)
        {
            InitializeComponent();

            _Id = null;            
            this.owner = owner;
            bdc = MainClass.Bdc;
            
            this.facultyId = facId;            
            this.passDate = date;
            this.examId = examId;
            this.studyBasisId = basisId;
            this.isAddVed = isAdd;

            using (PriemEntities context = new PriemEntities())
            {
                if (isAddVed)
                {
                    addCount = (from vd in context.extExamsVed
                                where vd.ExamId == examId && vd.Date == passDate.Date
                                && vd.FacultyId == facultyId && vd.StudyLevelGroupId == MainClass.studyLevelGroupId
                                && (studyBasisId != null ? vd.StudyBasisId == studyBasisId : true == true)
                                select vd.AddCount).Max();
                    
                    if (addCount == null)
                        addCount = 0;

                    addCount++;
                }
                else
                    addCount = 0;

                isAdditional = (from vd in context.extExamsVed
                                 where vd.ExamId == examId
                                select vd.IsAdditional).FirstOrDefault(); 
            }

            InitControls();
        }

        public int? cbStudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }

        public int? cbStudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }

        public int? cbObrazProgramId
        {
            get { return ComboServ.GetComboIdInt(cbObrazProgram); }
            set { ComboServ.SetComboId(cbObrazProgram, value); }
        }


        //дополнительная инициализация
        protected virtual void InitControls()
        {
            this.MdiParent = MainClass.mainform;
            this.CenterToParent();
            InitFocusHandlers();

            Dgv = dgvRight;

            if (MainClass.IsPasha())
                dtPassDate.Enabled = true;
            else
                dtPassDate.Enabled = false;

            using (PriemEntities context = new PriemEntities())
            {
                tbExam.Text = (from vd in context.extExamInEntry
                               where vd.ExamId == examId
                               select vd.ExamName).FirstOrDefault();                    
                   
                dtPassDate.Value = passDate;
                if (isAddVed)
                {
                    dtPassDate.Enabled = false;
                    lblAdd.Visible = true;
                    if (addCount == 0)
                        lblAddCount.Text = "";
                    else
                        lblAddCount.Text = "(" + addCount + ")";
                }
                if (_Id == null)
                    dtPassDate.Enabled = false;

                if (studyBasisId == null)
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, true);                    
                else
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByQuery(string.Format("SELECT CONVERT(varchar(100), Id) AS Id, Name FROM ed.StudyBasis WHERE Id = {0} ORDER BY Name", studyBasisId)), false, false);

                ComboServ.FillCombo(cbStudyForm, HelpClass.GetComboListByQuery(string.Format("SELECT DISTINCT CONVERT(varchar(100), StudyFormId) AS Id, StudyFormName AS Name FROM ed.qEntry WHERE StudyLevelGroupId = {0} AND FacultyId = {1} ORDER BY Name", MainClass.studyLevelGroupId, facultyId)), false, true);
                FillObrazProgram();

                //заполнение гридов            
                FillGridLeft();
                UpdateRightGrid();               

                cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
                cbStudyForm.SelectedIndexChanged += new EventHandler(cbStudyForm_SelectedIndexChanged);
                cbObrazProgram.SelectedIndexChanged += new EventHandler(cbObrazProgram_SelectedIndexChanged);                   
            }
        }

        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateRightGrid();
        }

        void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillObrazProgram();
            UpdateRightGrid();
        }       

        void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateRightGrid();
        }        

        private void FillObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = ((from ent in MainClass.GetEntry(context)
                                                           where ent.FacultyId == facultyId
                                                           && (cbStudyBasisId != null ? ent.StudyBasisId == cbStudyBasisId : true == true)
                                                           && (cbStudyFormId != null ? ent.StudyBasisId == cbStudyFormId : true == true)                                                           
                                                           select new
                                                           {
                                                               Id = ent.ObrazProgramId,
                                                               Name = ent.ObrazProgramName,
                                                               Crypt = ent.ObrazProgramCrypt
                                                           }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name + ' ' + u.Crypt)).ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, true);
            }
        }       

        private void UpdateRightGrid()
        {
            string rez = string.Empty;            

            //обработали форму обучения  
            if (cbStudyFormId != null)
                rez += " AND ed.qAbiturient.StudyFormId = " + cbStudyFormId;

            //обработали основу обучения 
            if (cbStudyBasisId != null)
                rez += " AND ed.qAbiturient.StudyBasisId = " + cbStudyBasisId; 
            
            //Направление
            if (cbObrazProgramId != null)
                rez += " AND ed.qAbiturient.ObrazProgramId = " + cbObrazProgramId;
            
            List<string> lstIds = new List<string>();
            foreach(DataGridViewRow dgrw in dgvLeft.Rows)
            {
                lstIds.Add(string.Format("'{0}'", dgrw.Cells["Id"].Value.ToString()));
            }

            if (lstIds.Count > 0)
                rez += string.Format(" AND ed.extPersonAspirant.Id NOT IN ({0}) ", Util.BuildStringWithCollection(lstIds));

            FillGridRight(rez);                        
        }

        private void FillGridRight(string flt_prof)
        {
            string flt_where = "";
            string flt_backDoc = "";
            string flt_enable = "";
            string flt_protocol = "";
            string flt_existMark = "";
            string flt_hasExam = "";
            string flt_notInVed = "";
            string flt_notAdd = "";
                           
            flt_backDoc = " AND ed.qAbiturient.BackDoc = 0 ";
            flt_enable = " AND ed.qAbiturient.NotEnabled = 0 ";
            //flt_protocol = " AND ed.qAbiturient.Id IN (SELECT ed.extProtocol.AbiturientId FROM ed.extProtocol WHERE ed.extProtocol.ProtocolTypeId = 1 AND ed.extProtocol.IsOld = 0 AND ed.extProtocol.Excluded = 0) ";
            flt_protocol = " AND ProtocolTypeId = 1 AND IsOld = 0 AND Excluded = 0";
            flt_hasExam = string.Format(" AND ed.qAbiturient.EntryId IN (SELECT ed.extExamInEntry.EntryId FROM ed.extExamInEntry WHERE ed.extExamInEntry.ExamId = {0})", examId);

            flt_where = string.Format(" WHERE ed.qAbiturient.FacultyId = {0} AND ed.qAbiturient.StudyLevelGroupId = {1} ", facultyId, MainClass.studyLevelGroupId) + flt_prof;

            if (studyBasisId != 2)
            {
                flt_existMark = string.Format(" AND ed.qAbiturient.Id NOT IN (SELECT ed.qMark.AbiturientId FROM ed.qMark INNER JOIN ed.extExamInEntry ON ed.qMark.ExamInEntryId = ed.extExamInEntry.Id WHERE ed.extExamInEntry.ExamId = {0} AND ed.extExamInEntry.FacultyId = {1}) ", examId, facultyId);

                if (isAdditional)
                    flt_notInVed = string.Format(" AND ed.extPersonAspirant.Id NOT IN (Select ed.ExamsVedHistory.PersonId FROM ed.ExamsVedHistory INNER JOIN ed.ExamsVed ON ed.ExamsVedHistory.ExamsVedId = ed.ExamsVed.Id WHERE ((ed.ExamsVed.IsLoad = 1 AND NOT ed.ExamsVedHistory.Mark IS NULL) OR (ed.ExamsVed.IsLoad = 0)) AND ed.ExamsVed.ExamId = {0} AND ed.ExamsVed.FacultyId = {1} {2}) ", examId, facultyId, (studyBasisId == null ? "" : " AND ed.ExamsVed.StudyBasisId = " + studyBasisId));
                else
                    flt_notInVed = string.Format(" AND ed.extPersonAspirant.Id NOT IN (Select ed.ExamsVedHistory.PersonId FROM ed.ExamsVedHistory INNER JOIN ed.ExamsVed ON ed.ExamsVedHistory.ExamsVedId = ed.ExamsVed.Id WHERE ((ed.ExamsVed.IsLoad = 1 AND NOT ed.ExamsVedHistory.Mark IS NULL) OR (ed.ExamsVed.IsLoad = 0)) AND ed.ExamsVed.ExamId = {0} {1}) ", examId, (studyBasisId == null ? "" : " AND ed.ExamsVed.StudyBasisId = " + studyBasisId));
            }
            
            if (!isAdditional)
                flt_notAdd = Exams.GetFilterForNotAddExam(examId, facultyId);
                                  
            FillGrid(dgvRight, sQuery + flt_where + flt_backDoc + flt_enable + flt_protocol + flt_existMark + flt_hasExam + flt_notInVed + flt_notAdd, "", sOrderby);
        }

        private void FillGridLeft()
        {
            string flt_where = string.Format(" WHERE ed.qAbiturient.FacultyId = {0} AND ed.qAbiturient.StudyLevelGroupId = {1} ", facultyId, MainClass.studyLevelGroupId);
            
            //заполнили левый
            if (_Id != null)
            {
                string sFilter = string.Format(" AND ed.extPersonAspirant.Id IN (Select ed.ExamsVedHistory.PersonId FROM ed.ExamsVedHistory WHERE ed.ExamsVedHistory.ExamsVedId = '{0}') ", _Id);
                FillGrid(dgvLeft, sQuery + flt_where, sFilter, sOrderby);
            }
            else //новый
            {
                InitGrid(dgvLeft);
            }
        }

        //подготовка нужного грида
        private void InitGrid(DataGridView dgv)
        {
            dgv.Columns.Clear();

            DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
            column.Width = 20;
            column.ReadOnly = false;
            column.Resizable = DataGridViewTriState.False;
            dgv.Columns.Add(column);
            dgv.Columns.Add("Id", "Id");
            dgv.Columns.Add("Num", "Ид_номер");
            dgv.Columns.Add("FIO", "ФИО");
            dgv.Columns.Add("Attestat", "Документ_об_образовании");
            dgv.Columns.Add("Pasport", "Паспорт");

            dgv.Columns["Id"].Visible = false;

            dgv.Update();
        }

        //функция заполнения грида
        private void FillGrid(DataGridView dgv, string sQuery, string sWhere, string sOrderby)
        {
            try
            {
                //подготовили грид
                InitGrid(dgv);
                DataSet ds = bdc.GetDataSet(sQuery + sWhere + sOrderby);

                //заполняем строки
                foreach (DataRow dr in ds.Tables[0].Rows)
                {                   
                    DataGridViewRow r = new DataGridViewRow();
                    r.CreateCells(dgv, false, dr["Id"].ToString(), dr["Ид_номер"].ToString(), dr["ФИО"].ToString(), dr["Документ_об_образовании"].ToString(), dr["Паспорт"].ToString());                   
                    dgv.Rows.Add(r);
                }

                //блокируем редактирование
                for (int i = 1; i < dgv.ColumnCount; i++)
                    dgv.Columns[i].ReadOnly = true;
                
                dgv.Update();               
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении грида данными протокола: " + ex.Message);
            }
        }

        private bool Save()
        {
            try
            {
                if (dgvLeft.Rows.Count == 0)
                {
                    MessageBox.Show("Нельзя создать пустую ведомость!", "Внимание");
                    return false;
                }

                using (PriemEntities context = new PriemEntities())
                {
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                    {
                        if (_Id == null)
                        {
                            ObjectParameter vedParId = new ObjectParameter("id", typeof(Guid));
                            context.ExamsVed_Insert(MainClass.studyLevelGroupId, facultyId, studyBasisId, passDate.Date, examId, false, false, isAddVed, (isAddVed ? addCount : null), vedParId);
                            _Id = (Guid)vedParId.Value;
                        }
                        else
                        {
                            if (dtPassDate.Value.Date != passDate.Date)
                            {
                                int cnt = (from vd in context.extExamsVed
                                           where vd.ExamId == examId && vd.Date == dtPassDate.Value.Date
                                           && vd.FacultyId == facultyId && (studyBasisId == null ? vd.StudyBasisId == null : vd.StudyBasisId == studyBasisId)
                                           select vd).Count();

                                if (cnt > 0)
                                {
                                    WinFormsServ.Error(string.Format("Ведомость на этот экзамен на эту дату {0} уже существует! ", studyBasisId == null ? "" : "на эту основу обучения"));
                                    return false;
                                }
                            }

                            context.ExamsVedHistory_DeleteByVedId(_Id);
                            if (dtPassDate.Value.Date != passDate.Date)
                                context.ExamsVed_UpdateDate(dtPassDate.Value.Date, _Id);
                        }

                        //записи в ведомостьхистори
                        foreach (DataGridViewRow r in dgvLeft.Rows)
                        {
                            Guid? persId = new Guid(r.Cells["Id"].Value.ToString());
                            context.ExamsVedHistory_InsertToVed(_Id, persId);                            
                        }

                        transaction.Complete();

                        return true;
                    }
                }
                
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при создании новой ведомости: " + ex.Message);
                return false;
            }
        }
        
        //отмена
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //все хорошо
        private void btnOk_Click(object sender, EventArgs e)
        {
            if(Save())
                this.Close();
        }

        //ресайз
        private void VedExamCrypto_Resize(object sender, EventArgs e)
        {
            int width = (this.Width - 90 - 11 * 2) / 2;
            dgvLeft.Width = dgvRight.Width = width;
            dgvLeft.Location = new Point(12, 112);
            dgvRight.Location = new Point(90 + width, 112);

            btnLeft.Left = 11 + width + 11;
            btnLeftAll.Left = 11 + width + 11;

            btnRight.Left = 90 + width - btnRight.Width - 11;
            btnRightAll.Left = 90 + width - btnRightAll.Width - 11;
        }               

        //перенос строк
        private void MoveRows(DataGridView from, DataGridView to, bool isAll)
        {
            for (int i = from.Rows.Count - 1; i >= 0; i--) 
            {
                DataGridViewRow dr = from.Rows[i];               

                if (isAll || (bool)dr.Cells[0].Value)
                {
                    dr.Cells[0].Value = false;
                    from.Rows.Remove(dr);
                    to.Rows.Add(dr);
                }
            }
        }
                
        //убрать влево
        private void btnLeft_Click(object sender, EventArgs e)
        {
            MoveRows(dgvRight, dgvLeft, false);
        }

        //убрать вправо
        private void btnRight_Click(object sender, EventArgs e)
        {
            MoveRows(dgvLeft, dgvRight, false);
        }

        //все влево
        private void btnLeftAll_Click(object sender, EventArgs e)
        {
            MoveRows(dgvRight, dgvLeft, true);
        }

        //все вправо
        private void btnRightAll_Click(object sender, EventArgs e)
        {
            MoveRows(dgvLeft, dgvRight, true);
        }

        private void ExamsVed_FormClosed(object sender, FormClosedEventArgs e)
        {
            owner.UpdateVedList();
            owner.SelectVed(_Id);
        }

        private void dgvRight_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                string abId = dgvRight.Rows[dgvRight.CurrentCell.RowIndex].Cells["Id"].Value.ToString();
                if (abId != "")
                {
                    MainClassCards.OpenCardPerson(MainClass.mainform, abId, this, dgvRight.CurrentRow.Index);
                }
            }                
        }        
    }
}
