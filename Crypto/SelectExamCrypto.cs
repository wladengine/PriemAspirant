using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;

using EducServLib;
using BDClassLib;
using BaseFormsLib;
using PriemLib;

namespace Priem
{
    public partial class SelectExamCrypto : BaseForm
    {
        private DBPriem bdc;
        private int? facultyId;
        private int? studyBasisId;
        private ExamsVedList owner;
        private bool isAdditional;

        public SelectExamCrypto(ExamsVedList owner, int? facId, int? basisId, DateTime passDate, int? examId)
        {
            InitializeComponent();            
            this.owner = owner;
            this.facultyId = facId;
            this.studyBasisId = basisId;
            this.isAdditional = true;            

            InitControls(); 
            cbExam.Enabled = false;
            dtDateExam.Enabled = false;

            dtDateExam.Value = passDate;            
            this.ExamId = examId;
        }

        public int? ExamId
        {
            get { return ComboServ.GetComboIdInt(cbExam); }
            set { ComboServ.SetComboId(cbExam, value); }
        }

        public SelectExamCrypto(ExamsVedList owner, int? facId, int? basisId)
        {
            InitializeComponent();
            this.owner = owner;
            this.facultyId = facId;
            this.studyBasisId = basisId;
            this.isAdditional = false;

            InitControls();
        }  

        //дополнительная инициализация контролов
        private void InitControls()
        {
            this.CenterToParent();
            InitFocusHandlers();
            this.MdiParent = MainClass.mainform;
            bdc = MainClass.Bdc; 
           
            FillExams();
        }

        private void FillExams()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = Exams.GetExamsWithFilters(context, facultyId, null, null, null, null, studyBasisId, null, null, null);
                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ExamId.ToString(), u.ExamName)).Distinct().ToList();
                ComboServ.FillCombo(cbExam, lst, false, false);                 
            }          
        } 

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!isAdditional)
            {
                using (PriemEntities context = new PriemEntities())
                {
                    int cnt = (from ev in context.extExamsVed
                               where ev.ExamId == ExamId && ev.Date == dtDateExam.Value.Date 
                               && ev.FacultyId == facultyId 
                               && (studyBasisId == null ? ev.StudyBasisId == null : ev.StudyBasisId == studyBasisId) 
                               && !ev.IsAddVed
                               select ev).Count();
                    
                    if (cnt > 0)
                    {
                        WinFormsServ.Error(string.Format("{0}едомость на этот экзамен на эту дату {1} уже существует! ", isAdditional ? "Дополнительная в" : "В", studyBasisId == null ? "" : "на эту основу обучения"));
                        return;
                    }
                }
            }

            ExamsVedCard frm = new ExamsVedCard(owner, facultyId, ExamId, dtDateExam.Value, studyBasisId, isAdditional);
            frm.Show();
            this.Close();
        }                     
    }
}