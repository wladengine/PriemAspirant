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
    public partial class OlympAbitBallsAndRatings : Form
    {
        #region Vals
        public int FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty).Value; }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
        public int LicenseProgramId
        {
            get { return ComboServ.GetComboIdInt(cbLicenseProgram).Value; }
            set { ComboServ.SetComboId(cbLicenseProgram, value); }
        }
        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }
        public int? OlympLevelId
        {
            get { return ComboServ.GetComboIdInt(cbOlympLevel); }
            set { ComboServ.SetComboId(cbOlympLevel, value); }
        }
        public int? OlympSubjectId
        {
            get { return ComboServ.GetComboIdInt(cbOlympSubject); }
            set { ComboServ.SetComboId(cbOlympSubject, value); }
        }
        #endregion

        //private bool started = false;
        public OlympAbitBallsAndRatings()
        {
            InitializeComponent();
            FillComboFaculty();
        }

        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.qEntry
                            where x.StudyLevelGroupId == 1 && x.StudyBasisId == 1
                            select new
                            {
                                x.FacultyId,
                                x.FacultyName
                            }).Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).ToList();

                ComboServ.FillCombo(cbFaculty, bind, false, false);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.qEntry
                            where x.StudyLevelGroupId == 1 && x.FacultyId == FacultyId
                            select new
                            {
                                x.LicenseProgramId,
                                x.LicenseProgramCode,
                                x.LicenseProgramName
                            }).Distinct().ToList()
                            .Select(x => new KeyValuePair<string, string>(x.LicenseProgramId.ToString(), x.LicenseProgramCode + " " + x.LicenseProgramName)).ToList();

                ComboServ.FillCombo(cbLicenseProgram, bind, false, false);
            }
        }
        private void FillComboStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.qEntry
                            where x.StudyLevelGroupId == 1 && x.StudyBasisId == 1
                            && x.FacultyId == FacultyId
                            select new
                            {
                                x.StudyFormId,
                                x.StudyFormName,
                            }).Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.StudyFormId.ToString(), x.StudyFormName)).ToList();

                ComboServ.FillCombo(cbStudyForm, bind, false, true);
            }
        }
        private void FillComboOlympLevel()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = 
                    (from olymp in context.OlympLevel
                     select new
                     {
                         olymp.Id,
                         olymp.Name
                     }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbOlympLevel, bind, false, true);
            }
        }
        private void FillComboOlympSubject()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind =
                    (from olymp in context.Olympiads
                     join ab in context.Abiturient
                     on olymp.AbiturientId equals ab.Id
                     join subj in context.OlympSubject
                     on olymp.OlympSubjectId equals subj.Id
                     where ab.Entry.FacultyId == FacultyId && ab.Entry.LicenseProgramId == LicenseProgramId
                     select new
                     {
                         subj.Id,
                         subj.Name
                     }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbOlympSubject, bind, false, true);
            }
        }

        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyForm();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboOlympLevel();
        }
        private void cbOlympLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboOlympSubject();
        }
        private void cbOlympSubject_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }

        private void FillGrid()
        {
            string query = @"SELECT ISNULL(t.OlympLevelName, 'Всеросс.') AS 'Уровень', t.OlympSubjectName AS 'Предмет', MIN(Ege) AS 'МинСуммаЕГЭ',
(SELECT MIN(hlpEntryMarksSum.TotalSum) FROM ed.hlpEntryMarksSum WHERE hlpEntryMarksSum.EntryId=t.EntryId AND hlpEntryMarksSum.CompetitionId IN (3,4)) AS 'ПроходнойБаллОбщКонк'
FROM
(
SELECT extOlympiads.OlympLevelName, extOlympiads.OlympSubjectName, qAbitAll.EntryId, SUM(hlpStatMaxApprovedEgeMarks.Value) AS Ege
FROM ed.qAbitAll
INNER JOIN ed.extOlympiads ON extOlympiads.AbiturientId = qAbitAll.Id
INNER JOIN ed.hlpStatMaxApprovedEgeMarks ON hlpStatMaxApprovedEgeMarks.PersonId = qAbitAll.PersonId
INNER JOIN ed.EgeToExam ON EgeToExam.EgeExamNameId = hlpStatMaxApprovedEgeMarks.EgeExamNameId
INNER JOIN ed.ExamInEntry ON ExamInEntry.EntryId = qAbitAll.EntryId AND ExamInEntry.ExamId = EgeToExam.ExamId AND ExamInEntry.EntryId = qAbitAll.EntryId
WHERE qAbitAll.StudyLevelGroupId=1 AND qAbitAll.FacultyId=@FacultyId AND qAbitAll.StudyBasisId=1
AND qAbitAll.LicenseProgramId=@LicenseProgramId
";
            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@FacultyId ", FacultyId);
            sl.Add("@LicenseProgramId", LicenseProgramId);

            if (StudyFormId.HasValue)
            {
                query += " AND qAbitAll.StudyFormId=@StudyFormId";
                sl.Add("@StudyFormId", StudyFormId.Value);
            }
            if (OlympLevelId.HasValue)
            {
                query += " AND extOlympiads.OlympLevelId=@OlympLevelId";
                sl.Add("@OlympLevelId", OlympLevelId.Value);
            }
            if (OlympSubjectId.HasValue)
            {
                query += " AND extOlympiads.OlympSubjectId=@OlympSubjectId";
                sl.Add("@OlympSubjectId", OlympSubjectId.Value);
            }

            //"хвост" запроса
            string tail = @"
GROUP BY extOlympiads.OlympLevelName, extOlympiads.OlympSubjectName, qAbitAll.EntryId, qAbitAll.Id
) t
GROUP BY t.OlympLevelName, t.OlympSubjectName, t.EntryId";
            DataTable tbl = MainClass.Bdc.GetDataSet(query + tail, sl).Tables[0];
            dgv.DataSource = tbl;
        }
    }
}
