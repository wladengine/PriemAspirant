using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using BaseFormsLib;
using EducServLib;
using PriemLib;
namespace Priem
{
    public partial class EgeStatistics : BaseFormEx
    {
        public EgeStatistics()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            Dgv = dgv;
            FillComboFaculty();
        }

        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst =
                    (from x in context.qEntry
                     where x.StudyLevelGroupId == 1
                     select new { x.FacultyId, x.FacultyName }
                     ).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).Distinct().ToList();
                ComboServ.FillCombo(cbFaculty, lst, false, false);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst =
                    (from x in context.qEntry
                     where x.StudyLevelGroupId == 1 && x.FacultyId == FacultyId
                     select new { x.LicenseProgramId, x.LicenseProgramCode, x.LicenseProgramName }
                     ).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.LicenseProgramId.ToString(), x.LicenseProgramCode + " " + x.LicenseProgramName)).Distinct().ToList();
                ComboServ.FillCombo(cbLicenseProgram, lst, false, true);
            }
        }
        private void FillComboObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var OP = from x in context.qEntry
                         where x.StudyLevelGroupId == 1 && x.FacultyId == FacultyId
                         select new { x.LicenseProgramId, x.ObrazProgramId, x.ObrazProgramName };
                if (LicenseProgramId != null)
                    OP = OP.Where(x => x.LicenseProgramId == LicenseProgramId);
                
                List<KeyValuePair<string, string>> lst = OP.Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.ObrazProgramId.ToString(), x.ObrazProgramName)).Distinct().ToList();
                ComboServ.FillCombo(cbObrazProgram, lst, false, true);
            }
        }
        private void FillComboProfile()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var Prof = from x in context.qEntry
                           where x.StudyLevelGroupId == 1 && x.FacultyId == FacultyId && x.ProfileId != null
                           select new { x.LicenseProgramId, x.ObrazProgramId, x.ProfileId, x.ProfileName };
                if (LicenseProgramId != null)
                    Prof = Prof.Where(x => x.LicenseProgramId == LicenseProgramId);
                if (ObrazProgramId != null)
                    Prof = Prof.Where(x => x.ObrazProgramId == ObrazProgramId);
                
                List<KeyValuePair<string, string>> lst = Prof.Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.ProfileId.ToString(), x.ProfileName)).Distinct().ToList();
                ComboServ.FillCombo(cbProfile, lst, false, true);
            }
        }
        private void FillComboStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var SF = from x in context.qEntry
                         where x.StudyLevelGroupId == 1 && x.FacultyId == FacultyId
                         select new { x.LicenseProgramId, x.ObrazProgramId, x.ProfileId, x.StudyFormId, x.StudyFormName };
                if (LicenseProgramId != null)
                    SF = SF.Where(x => x.LicenseProgramId == LicenseProgramId);
                if (ObrazProgramId != null)
                    SF = SF.Where(x => x.ObrazProgramId == ObrazProgramId);
                if (ProfileId != null)
                    SF = SF.Where(x => x.ProfileId == ProfileId);

                List<KeyValuePair<string, string>> lst = SF.Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.StudyFormId.ToString(), x.StudyFormName)).Distinct().ToList();
                ComboServ.FillCombo(cbStudyForm, lst, false, true);
            }
        }
        private void FillComboStudyBasis()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var SB = from x in context.qEntry
                         where x.StudyLevelGroupId == 1 && x.FacultyId == FacultyId
                         select new { x.LicenseProgramId, x.ObrazProgramId, x.ProfileId, x.StudyFormId, x.StudyBasisId, x.StudyBasisName };
                if (LicenseProgramId != null)
                    SB = SB.Where(x => x.LicenseProgramId == LicenseProgramId);
                if (ObrazProgramId != null)
                    SB = SB.Where(x => x.ObrazProgramId == ObrazProgramId);
                if (ProfileId != null)
                    SB = SB.Where(x => x.ProfileId == ProfileId);
                if (StudyFormId != null)
                    SB = SB.Where(x => x.StudyFormId == StudyFormId);

                List<KeyValuePair<string, string>> lst = SB.Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.StudyBasisId.ToString(), x.StudyBasisName)).Distinct().ToList();
                ComboServ.FillCombo(cbStudyBasis, lst, false, true);
            }
        }

        private void FillGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                //стартовый микро-запрос количества людей по фильтру
                string query = @"SELECT COUNT(DISTINCT extAbitAspirant.PersonId)
FROM ed.extAbitAspirant
WHERE extAbitAspirant.StudyLevelGroupId=1 AND FacultyId=" + FacultyId.ToString() + " ";
                string where = "";
                if (LicenseProgramId != null)
                    where += " AND extAbitAspirant.LicenseProgramId=" + LicenseProgramId.ToString();
                if (ObrazProgramId != null)
                    where += " AND extAbitAspirant.ObrazProgramId=" + ObrazProgramId.ToString();
                if (ProfileId != null)
                    where += " AND extAbitAspirant.ProfileId='" + ProfileId.ToString() + "'";
                if (StudyFormId != null)
                    where += " AND extAbitAspirant.StudyFormId=" + StudyFormId.ToString();
                if (StudyBasisId != null)
                    where += " AND extAbitAspirant.StudyBasisId=" + StudyBasisId.ToString();
                int cnt = (int)MainClass.Bdc.GetValue(query + where);
                Watch wc;
                if (cnt > 0)
                {
                    wc = new Watch(cnt);
                    wc.Show();
                }
                else
                {
                    wc = null;
                    return;
                }

                query = @"
                    SELECT DISTINCT 
extAbitAspirant.PersonId, extAbitAspirant.FIO, EgeExamName.Name, hlpStatMaxApprovedEgeMarks.Value, (CASE WHEN hlpStatMaxApprovedEgeMarks.EgeExamNameId IN (11, 12, 13, 14) THEN 1 ELSE 0 END) AS ForeignLang
FROM ed.extAbitAspirant
INNER JOIN ed.extExamInEntry ON extExamInEntry.EntryId = extAbitAspirant.EntryId
INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = extExamInEntry.ExamId
INNER JOIN ed.EgeExamName ON EgeExamName.Id = EgeToExam.EgeExamNameId
INNER JOIN ed.hlpStatMaxApprovedEgeMarks ON hlpStatMaxApprovedEgeMarks.PersonId = extAbitAspirant.PersonId AND hlpStatMaxApprovedEgeMarks.EgeExamNameId = EgeExamName.Id
WHERE extAbitAspirant.StudyLevelGroupId=1  AND extAbitAspirant.FacultyId=" + FacultyId.ToString() + " ";
                where = "";
                if (LicenseProgramId != null)
                    where += " AND extAbitAspirant.LicenseProgramId=" + LicenseProgramId.ToString();
                if (ObrazProgramId != null)
                    where += " AND extAbitAspirant.ObrazProgramId=" + ObrazProgramId.ToString();
                if (ProfileId != null)
                    where += " AND extAbitAspirant.ProfileId='" + ProfileId.ToString() + "'";
                if (StudyFormId != null)
                    where += " AND extAbitAspirant.StudyFormId=" + StudyFormId.ToString();
                if (StudyBasisId != null)
                    where += " AND extAbitAspirant.StudyBasisId=" + StudyBasisId.ToString();

                DataTable tblPersons = MainClass.Bdc.GetDataSet(query + where + " ORDER BY hlpStatMaxApprovedEgeMarks.Value ").Tables[0];
                var p_RAW = (from DataRow rw in tblPersons.Rows
                             select new
                             {
                                 PersonId = rw.Field<Guid>("PersonId"),
                                 FIO = rw.Field<string>("FIO"),
                                 ExamName = rw.Field<string>("Name"),
                                 IsForeign = rw.Field<int>("ForeignLang") == 1 ? true : false,
                                 Value = rw.Field<int>("Value")
                             });

                query = @"
                SELECT PersonId, (case when extAbitAspirant.CompetitionId IN (1,2,7,8) then 1 else 0 end) AS VK 
                FROM ed.extAbitAspirant
                WHERE extAbitAspirant.StudyLevelGroupId=1 AND extAbitAspirant.BackDoc=0
                AND extAbitAspirant.FacultyId=" + FacultyId.ToString() + " ";
                where = "";
                if (LicenseProgramId != null)
                    where += " AND extAbitAspirant.LicenseProgramId=" + LicenseProgramId.ToString();
                if (ObrazProgramId != null)
                    where += " AND extAbitAspirant.ObrazProgramId=" + ObrazProgramId.ToString();
                if (ProfileId != null)
                    where += " AND extAbitAspirant.ProfileId='" + ProfileId.ToString() + "'";
                if (StudyFormId != null)
                    where += " AND extAbitAspirant.StudyFormId=" + StudyFormId.ToString();
                if (StudyBasisId != null)
                    where += " AND extAbitAspirant.StudyBasisId=" + StudyBasisId.ToString();

                DataTable tblVKs = MainClass.Bdc.GetDataSet(query + where).Tables[0];

                var VKs = from DataRow rw in tblVKs.Rows
                          select new
                          {
                              PersonId = rw.Field<Guid>("PersonId"),
                              VK = rw.Field<int>("VK") == 1 ? true : false
                          };

                List<Guid> persons = p_RAW.Select(x => x.PersonId).Distinct().ToList();
                
                query = @"SELECT DISTINCT 
EgeExamName.Name
FROM ed.extExamInEntry
INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = extExamInEntry.ExamId
INNER JOIN ed.EgeExamName ON EgeExamName.Id = EgeToExam.EgeExamNameId";
                where = " WHERE extExamInEntry.FacultyId=" + FacultyId.ToString() + " ";
                if (LicenseProgramId != null)
                    where += " AND extExamInEntry.LicenseProgramId=" + LicenseProgramId.ToString();
                if (ObrazProgramId != null)
                    where += " AND extExamInEntry.ObrazProgramId=" + ObrazProgramId.ToString();
                if (ProfileId != null)
                    where += " AND extExamInEntry.ProfileId='" + ProfileId.ToString();
                if (StudyFormId != null)
                    where += " AND extExamInEntry.StudyFormId=" + StudyFormId.ToString();
                if (StudyBasisId != null)
                    where += " AND extExamInEntry.StudyBasisId=" + StudyBasisId.ToString();

                DataTable tblExams = MainClass.Bdc.GetDataSet(query + where).Tables[0];

                dgv.DataSource = null;
                DataTable tblSource = new DataTable();
                tblSource.Columns.Add("Id", typeof(Guid));
                tblSource.Columns.Add("ФИО", typeof(string));
                foreach (DataRow rw in tblExams.Rows)
                {
                    tblSource.Columns.Add(rw.Field<string>("Name"), typeof(int));
                }
                tblSource.Columns.Add("Сумма", typeof(int));
                tblSource.Columns.Add("Green", typeof(bool));

                Guid PersonId;
                foreach (Guid pId in persons)
                {
                    var rwData = from x in p_RAW
                                 where x.PersonId == pId
                                 select new { x.PersonId, x.FIO, x.ExamName, x.Value, x.IsForeign };
                    DataRow rw = tblSource.NewRow();
                    PersonId = rwData.First().PersonId;
                    rw["Id"] = PersonId;
                    rw["ФИО"] = rwData.First().FIO;
                    foreach (var ex in rwData.Select(x => new { x.ExamName, x.Value }))
                    {
                        rw[ex.ExamName] = ex.Value;
                    }
                    rw["Сумма"] = rwData.Where(x => x.IsForeign == false).Select(x => x.Value).DefaultIfEmpty(0).Sum() + rwData.Where(x => x.IsForeign == true).Select(x => x.Value).DefaultIfEmpty(0).Max();
                    rw["Green"] = VKs.Where(x => x.PersonId == PersonId && x.VK == true).Select(x => x.VK).DefaultIfEmpty(false).First();
                    tblSource.Rows.Add(rw);
                    wc.PerformStep();
                }

                query = "SELECT SUM(KCP) FROM ed.qEntry WHERE StudyLevelGroupId=1 AND FacultyId=" + FacultyId.ToString() + " ";
                where = "";
                if (LicenseProgramId != null)
                    where += " AND LicenseProgramId=" + LicenseProgramId.ToString();
                if (ObrazProgramId != null)
                    where += " AND ObrazProgramId=" + ObrazProgramId.ToString();
                if (ProfileId != null)
                    where += " AND ProfileId='" + ProfileId.ToString();
                if (StudyFormId != null)
                    where += " AND StudyFormId=" + StudyFormId.ToString();
                if (StudyBasisId != null)
                    where += " AND StudyBasisId=" + StudyBasisId.ToString();

                int kcp = (int)MainClass.Bdc.GetValue(query + where);
                tbKCP.Text = kcp.ToString();

                var p_sums = (from DataRow rw in tblSource.Rows
                              select new { Id = rw.Field<Guid>("Id"), Sum = rw.Field<int>("Сумма") }).OrderByDescending(x => x.Sum);

                int i = VKs.Where(x => x.VK == true).Select(x => x.PersonId).Distinct().Count();

                foreach (var p in p_sums)
                {
                    if (++i > kcp)
                        break;

                    for (int j = 0; j < tblSource.Rows.Count; j++)
                    {
                        if (tblSource.Rows[j].Field<Guid>("Id") == p.Id)
                        {
                            tblSource.Rows[j].SetField<bool>("Green", true);
                            break;
                        }
                    }
                }

                dgv.DataSource = tblSource;
                dgv.Columns["Id"].Visible = false;
                dgv.Columns["Green"].Visible = false;
                dgv.Sort(dgv.Columns["Сумма"], ListSortDirection.Ascending);
                dgv.Columns["ФИО"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                
                wc.Close();
                wc = null;
            }
        }
        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboObrazProgram();
        }
        private void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboProfile();
        }
        private void cbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyForm();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyBasis();
        }
        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > -1 && e.RowIndex > -1)
                MainClassCards.OpenCardPerson(MainClass.mainform, dgv.Rows[e.RowIndex].Cells["Id"].Value.ToString(), this, e.RowIndex);
        }
        private void dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (e.ColumnIndex == Dgv.Columns["ФИО"].Index || e.ColumnIndex == Dgv.Columns["Сумма"].Index)
            {
                if ((bool)dgv.Rows[e.RowIndex].Cells["Green"].Value)
                    e.CellStyle.BackColor = Color.LightGreen;
            }
        }
    }
}
