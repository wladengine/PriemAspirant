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
using PriemLib;

namespace Priem
{
    public partial class LoadEgeMarks : Form
    {
        private DBPriem bdc;
        private Watch wtc;
        private int marksCount;

        public LoadEgeMarks()
        {
            InitializeComponent();           
            InitControls();            
        }  

        //дополнительная инициализация контролов
        private void InitControls()
        {
            this.CenterToParent(); 
            this.MdiParent = MainClass.mainform;
            bdc = MainClass.Bdc;

            ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.SP_Faculty"), false, false);             
            FillExams();

            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);                 
        }

        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillExams();
        }

        
        private void FillExams()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = Exams.GetExamsWithFilters(context, FacultyId, null, null, null, null, null, null, null, null).Where(c=>!c.IsAdditional);               

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ExamId.ToString(), u.ExamName)).Distinct().ToList();
                ComboServ.FillCombo(cbExam, lst, false, true);
            }             
        }

        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }

        public int? ExamId
        {
            get { return ComboServ.GetComboIdInt(cbExam); }
            set { ComboServ.SetComboId(cbExam, value); }
        }

        //строим запрос фильтров для абитуриентов
        private string GetAbitFilterString()
        {
            string s = " AND ed.qAbiturient.StudyLevelGroupId = 1";
 
            //обработали факультет            
            if (FacultyId != null)
                s += " AND ed.qAbiturient.FacultyId = " + FacultyId.ToString(); 

            return s;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsPasha())
                return;
            
            LoadMarks();
        }

        private void LoadMarks()
        {
            if (!MainClass.IsPasha())
                return;

            try
            {
                wtc = new Watch(2);
                wtc.Show();
                marksCount = 0;

                if (ExamId == null)
                {
                    foreach (KeyValuePair<string, string> ex in cbExam.Items)
                    {
                        int exId;
                        if (int.TryParse(ex.Key, out exId))
                            SetMarksForExam(exId);
                    }
                }
                else
                    SetMarksForExam(ExamId);
                
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка загрузки оценок " + ex.Message);
            }
            finally
            {
                wtc.Close();
                MessageBox.Show(string.Format("Зачтено {0} оценок", marksCount));
            }
        }

        private void SetMarksForExam(int? examId)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {                    
                    int examInostr = 65;                       
                    int filFacId = 18;

                    string flt_backDoc = " AND ed.qAbiturient.BackDoc = 0  ";
                    string flt_enable = " AND ed.qAbiturient.NotEnabled = 0 ";
                    string flt_protocol = " AND ProtocolTypeId = 1 AND IsOld = 0 AND Excluded = 0 ";
                    string flt_status = " AND ed.extFBSStatus.FBSStatusId IN (1,4) ";
                    string flt_mark = string.Format(" AND NOT EXISTS( SELECT ed.Mark.Value FROM ed.Mark INNER JOIN ed.extExamInEntry ON ed.Mark.ExamInEntryId = ed.extExamInEntry.Id WHERE ed.Mark.AbiturientId = ed.qAbiturient.Id AND ed.extExamInEntry.ExamId = {0} AND ed.extExamInEntry.IsAdditional=0) ", examId);
                    string flt_hasEge = string.Format(" AND ed.Person.Id IN (SELECT PersonId FROM  ed.extEgeMark LEFT JOIN ed.EgeToExam ON ed.extEgeMark.EgeExamNameId = ed.EgeToExam.EgeExamNameId WHERE ed.EgeToExam.ExamId = {0})", examId);
                    string flt_hasExam = string.Format(" AND ed.qAbiturient.EntryId IN (SELECT ed.extExamInEntry.EntryId FROM ed.extExamInEntry WHERE ed.extExamInEntry.ExamId = {0})", examId);

                    string queryAbits = @"SELECT ed.qAbiturient.Id, ed.qAbiturient.FacultyId, ed.qAbiturient.EntryId FROM ed.qAbiturient 
                            LEFT JOIN ed.Person ON ed.qAbiturient.PersonId = ed.Person.Id LEFT JOIN ed.extFBSStatus ON ed.extFBSStatus.PersonId = Person.Id 
                            LEFT JOIN ed.extProtocol ON extProtocol.AbiturientId = qAbiturient.Id WHERE 1 = 1 ";

                    DataSet ds = bdc.GetDataSet(queryAbits + GetAbitFilterString() + flt_backDoc + flt_enable + flt_protocol + flt_status + flt_mark + flt_hasExam + flt_hasEge);

                    wtc.pBar.Maximum += ds.Tables[0].Rows.Count;

                    //using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                    //{
                        try
                        {
                            foreach (DataRow dsRow in ds.Tables[0].Rows)
                            {
                                int? balls = null;
                                Guid abId = new Guid(dsRow["Id"].ToString());
                                Guid entryId = new Guid(dsRow["EntryId"].ToString());
                                
                                int? exInEntryId = (from eie in context.extExamInEntry
                                                    where eie.EntryId == entryId && eie.ExamId == examId
                                                    select eie.Id).FirstOrDefault();

                                if (exInEntryId == null)
                                    continue;

                                Guid egeCertificateId;

                                if (examId != examInostr)
                                {
                                    balls = (from emm in context.extEgeMarkMaxAbit
                                             join ete in context.EgeToExam
                                             on emm.EgeExamNameId equals ete.EgeExamNameId
                                             where emm.AbiturientId == abId && ete.ExamId == examId
                                             select emm.Value).Max();
                                    egeCertificateId = (from emm in context.extEgeMarkMaxAbit
                                             join ete in context.EgeToExam
                                             on emm.EgeExamNameId equals ete.EgeExamNameId
                                             where emm.AbiturientId == abId && ete.ExamId == examId && emm.Value == balls
                                             select emm.EgeCertificateId).FirstOrDefault();
                                }
                                else
                                {
                                    List<int> lstInostr = (from ete in context.EgeToExam
                                                           where ete.ExamId == examInostr
                                                           select ete.EgeExamNameId).ToList<int>();

                                    if (dsRow["FacultyId"].ToString() == filFacId.ToString())
                                    {
                                        int? egeExamNameId = (from etl in context.EgeToLanguage
                                                              join ab in context.qAbiturient
                                                              on etl.LanguageId equals ab.LanguageId
                                                              where etl.ExamId == examInostr && ab.Id == abId
                                                              select etl.EgeExamNameId).FirstOrDefault();

                                        if (egeExamNameId != null)
                                            balls = (from emm in context.extEgeMarkMaxAbit
                                                     where emm.AbiturientId == abId && emm.EgeExamNameId == egeExamNameId
                                                     select emm.Value).Max();

                                        egeCertificateId = (from emm in context.extEgeMarkMaxAbit
                                             join ete in context.EgeToExam
                                             on emm.EgeExamNameId equals ete.EgeExamNameId
                                             where emm.AbiturientId == abId && ete.ExamId == examId && emm.Value == balls
                                             select emm.EgeCertificateId).FirstOrDefault();
                                    }
                                    else
                                    {
                                        int cntEM = (from emm in context.extEgeMarkMaxAbit
                                                     where lstInostr.Contains(emm.EgeExamNameId) && emm.AbiturientId == abId
                                                     select emm.EgeMarkId).Count();

                                        if (cntEM > 1)
                                        {
                                            int? egeExamNameId = (from etl in context.EgeToLanguage
                                                                  join ab in context.qAbiturient
                                                                  on etl.LanguageId equals ab.LanguageId
                                                                  where etl.ExamId == examInostr && ab.Id == abId
                                                                  select etl.EgeExamNameId).FirstOrDefault();

                                            if (egeExamNameId != null)
                                                balls = (from emm in context.extEgeMarkMaxAbit
                                                         where emm.AbiturientId == abId && emm.EgeExamNameId == egeExamNameId
                                                         select emm.Value).Max();
                                            egeCertificateId = 
                                                (from emm in context.extEgeMarkMaxAbit
                                                 join ete in context.EgeToExam
                                                 on emm.EgeExamNameId equals ete.EgeExamNameId
                                                 where emm.AbiturientId == abId && ete.ExamId == examId && emm.Value == balls
                                                 select emm.EgeCertificateId).FirstOrDefault();
                                        }
                                        else
                                        {
                                            balls = (from emm in context.extEgeMarkMaxAbit
                                                     join ete in context.EgeToExam
                                                     on emm.EgeExamNameId equals ete.EgeExamNameId
                                                     where emm.AbiturientId == abId && ete.ExamId == examId
                                                     select emm.Value).Max();
                                        
                                            egeCertificateId = 
                                                (from emm in context.extEgeMarkMaxAbit
                                                 join ete in context.EgeToExam
                                                 on emm.EgeExamNameId equals ete.EgeExamNameId
                                                 where emm.AbiturientId == abId && ete.ExamId == examId && emm.Value == balls
                                                 select emm.EgeCertificateId).FirstOrDefault();
                                        }
                                    }
                                }

                                if (balls != null)
                                    context.Mark_Insert(abId, exInEntryId, balls, dtDateExam.Value.Date, true, false, false, null, null, egeCertificateId);
                                else
                                    continue;

                                wtc.PerformStep();
                                marksCount++;
                            }

                            //transaction.Complete();
                        }
                        catch (Exception exc)
                        {
                            throw new Exception("Ошибка загрузки оценок1: " + exc.Message);
                        }
                   // }
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка загрузки оценок " + ex.Message);                
                wtc.Close();
            }            
        }
    }
}