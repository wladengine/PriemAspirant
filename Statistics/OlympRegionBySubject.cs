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
    public partial class OlympRegionBySubject : Form
    {
        #region Vals
        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
        public int? SubjectId
        {
            get { return ComboServ.GetComboIdInt(cbSubject); }
            set { ComboServ.SetComboId(cbSubject, value); }
        }
        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }
        public int? StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }
        #endregion

        private bool started = false;
        public OlympRegionBySubject()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            FillComboFaculty();
            started = true;
        }

        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.qEntry
                            where x.StudyLevelGroupId == 1
                            select new
                            {
                                x.FacultyId,
                                x.FacultyName
                            }).Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).ToList();

                ComboServ.FillCombo(cbFaculty, bind, false, false);
            }
        }
        private void FillComboSubject()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.OlympSubject
                            select new
                            {
                                x.Id,
                                x.Name
                            }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbSubject, bind, false, false);
            }
        }
        private void FillComboStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.qEntry
                            where x.StudyLevelGroupId == 1
                            && (FacultyId.HasValue ? x.FacultyId == FacultyId.Value : 1 == 1)
                            select new
                            {
                                x.StudyFormId,
                                x.StudyFormName,
                            }).Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.StudyFormId.ToString(), x.StudyFormName)).ToList();

                ComboServ.FillCombo(cbStudyForm, bind, false, true);
            }
        }
        private void FillComboStudyBasis()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.qEntry
                            where x.StudyLevelGroupId == 1
                            && (FacultyId.HasValue ? x.FacultyId == FacultyId.Value : true)
                            select new
                            {
                                x.StudyBasisId,
                                x.StudyBasisName
                            }).Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.StudyBasisId.ToString(), x.StudyBasisName)).ToList();

                ComboServ.FillCombo(cbStudyBasis, bind, false, true);
            }
        }

        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyForm();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyBasis();
        }
        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!started)
                FillComboSubject();
            else
                FillGrid();
        }
        private void cbSubject_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }

        private void FillGrid()
        {
            string query = @"SELECT extPersonAspirant.RegionName AS 'Регион', OlympSubject.Name AS 'Предмет', COUNT (DISTINCT qAbitAll.PersonId) AS 'Число абитуриентов', COUNT (DISTINCT extEntryView.AbiturientId) AS 'Число поступивших'
FROM ed.qAbitAll
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id = qAbitAll.PersonId
INNER JOIN ed.Olympiads ON Olympiads.AbiturientId = qAbitAll.Id
INNER JOIN ed.OlympSubject ON OlympSubject.Id = Olympiads.OlympSubjectId
LEFT JOIN ed.extEntryView ON extEntryView.AbiturientId = qAbitAll.Id
WHERE qAbitAll.StudyLevelGroupId = 1 AND OlympSubject.Id = @OlympSubjectId AND extPersonAspirant.RegionId IS NOT NULL
";
            string groupby = " GROUP BY extPersonAspirant.RegionName, OlympSubject.Name";
            SortedList<string, object> sl = new SortedList<string,object>();
            sl.Add("@OlympSubjectId", SubjectId);

            if (FacultyId.HasValue)
            {
                query += " AND qAbitAll.FacultyId = @FacultyId";
                sl.Add("@FacultyId", FacultyId);
            }
            if (StudyFormId.HasValue)
            {
                query += " AND qAbitAll.StudyFormId = @StudyFormId";
                sl.Add("@StudyFormId", StudyFormId);
            }
            if (StudyBasisId.HasValue)
            {
                query += " AND qAbitAll.StudyBasisId = @StudyBasisId";
                sl.Add("@StudyBasisId", StudyBasisId);
            }

            DataTable tbl = MainClass.Bdc.GetDataSet(query + groupby, sl).Tables[0];
            dgv.DataSource = tbl;
        }
    }
}
