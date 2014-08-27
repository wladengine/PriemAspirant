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
    public partial class OlympLevelAbitRating : Form
    {
        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
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

        public OlympLevelAbitRating()
        {
            InitializeComponent();
            FillComboFaculty();
        }

        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.qEntry
                            select new
                            {
                                x.FacultyId,
                                x.FacultyName
                            }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).ToList();
                
                ComboServ.FillCombo(cbFaculty, bind, false, true);
            }
        }
        private void FillComboOlympLevel()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.Olympiads
                            where FacultyId.HasValue ? x.Abiturient.Entry.FacultyId == FacultyId.Value : true
                            select new
                            {
                                x.OlympLevelId,
                                x.OlympLevel.Name
                            }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.OlympLevelId.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbOlympLevel, bind, false, true);
            }
        }
        private void FillComboOlympSubject()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.Olympiads
                            where FacultyId.HasValue ? x.Abiturient.Entry.FacultyId == FacultyId.Value : true
                            && OlympLevelId.HasValue ? x.OlympLevelId == OlympLevelId.Value : true
                            select new
                            {
                                x.OlympSubjectId,
                                x.OlympSubject.Name
                            }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.OlympSubjectId.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbOlympSubject, bind, false, true);
            }
        }

        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
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
            string query = @"SELECT extOlympiads.OlympSubjectName AS 'Предмет', extOlympiads.OlympLevelName AS 'Уровень', 
COUNT(DISTINCT extAbitAspirant.PersonId) AS 'Число подавших документы', COUNT(DISTINCT extEntryView.AbiturientId) AS 'Число зачисленных'
FROM ed.extAbitAspirant
INNER JOIN ed.extOlympiads ON extOlympiads.AbiturientId=extAbitAspirant.Id
LEFT JOIN ed.extEntryView ON extEntryView.AbiturientId=extAbitAspirant.Id
WHERE extAbitAspirant.StudyLevelGroupId=@StudyLevelGroupId ";
            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@StudyLevelGroupId", MainClass.studyLevelGroupId);
            
            if (FacultyId.HasValue)
            {
                query += " AND extAbitAspirant.FacultyId=@FacultyId";
                sl.Add("@FacultyId", FacultyId.Value);
            }
            if (OlympLevelId.HasValue)
            {
                query += " AND OlympLevelId=@OlympLevelId";
                sl.Add("@OlympLevelId", OlympLevelId.Value);
            }
            if (OlympSubjectId.HasValue)
            {
                query += " AND OlympSubjectId=@OlympSubjectId";
                sl.Add("@OlympSubjectId", OlympSubjectId.Value);
            }

            string groupBy = " GROUP BY extOlympiads.OlympSubjectName, extOlympiads.OlympLevelName";

            DataTable tbl = MainClass.Bdc.GetDataSet(query + groupBy, sl).Tables[0];
            dgv.DataSource = tbl;
        }

        
    }
}
