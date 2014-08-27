using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using RtfWriter;
using PriemLib;
namespace Priem
{
    public partial class AbitEgeMarksStatistics : Form
    {
        public int? FacultyId
        {
            get
            {
                if (cbFaculty.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbFaculty);
            }
        }
        public int? LicenseProgramId
        {
            get
            {
                if (cbLicenseProgram.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbLicenseProgram);
            }
        }
        public int? ObrazProgramId
        {
            get
            {
                if (cbObrazProgram.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbObrazProgram);
            }
        }
        public Guid? ProfileId
        {
            get
            {
                if (cbObrazProgram.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                Guid g;
                Guid.TryParse(ComboServ.GetComboId(cbObrazProgram), out g);
                return g;
            }
        }
        public int? StudyFormId
        {
            get
            {
                if (cbStudyForm.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbStudyForm);
            }
        }
        public int? StudyBasisId
        {
            get
            {
                if (cbStudyBasis.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbStudyBasis);
            }
        }
        public int EgeExamNameId
        {
            get
            {
                return ComboServ.GetComboIdInt(cbEgeExamName).Value;
            }
            set
            {
                ComboServ.SetComboId(cbEgeExamName, (int?)value);
            }
        }
        public Dictionary<int, string> EgeExams { get; set; }
        
        public AbitEgeMarksStatistics()
        {
            InitializeComponent();

            this.MdiParent = MainClass.mainform;

            FillComboFaculty();
        }

        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.qEntry
                           where x.StudyLevelGroupId == MainClass.studyLevelGroupId
                           select new
                           {
                               x.FacultyId,
                               x.FacultyName
                           }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).ToList();
                
                ComboServ.FillCombo(cbFaculty, src, false, true);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.qEntry
                           where x.StudyLevelGroupId == MainClass.studyLevelGroupId && x.FacultyId == FacultyId
                           select new
                           {
                               x.LicenseProgramId,
                               x.LicenseProgramCode,
                               x.LicenseProgramName
                           }).Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.LicenseProgramId.ToString(), x.LicenseProgramCode + " " + x.LicenseProgramName)).ToList();

                ComboServ.FillCombo(cbLicenseProgram, src, false, true);
            }
        }
        private void FillComboObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.qEntry
                           where x.StudyLevelGroupId == MainClass.studyLevelGroupId && x.FacultyId == FacultyId
                           select new
                           {
                               x.LicenseProgramId,
                               x.ObrazProgramId,
                               x.ObrazProgramCrypt,
                               x.ObrazProgramName
                           }).Distinct();

                if (LicenseProgramId.HasValue)
                    src = src.Where(x => x.LicenseProgramId == LicenseProgramId);

                var bind = src.Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.ObrazProgramId.ToString(), x.ObrazProgramCrypt + " " + x.ObrazProgramName)).ToList();

                ComboServ.FillCombo(cbObrazProgram, bind, false, true);
            }
        }
        private void FillComboStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.qEntry
                           where x.StudyLevelGroupId == MainClass.studyLevelGroupId && x.FacultyId == FacultyId
                           select new
                           {
                               x.LicenseProgramId,
                               x.ObrazProgramId,
                               x.ProfileId,
                               x.StudyFormId,
                               x.StudyFormName
                           }).Distinct();

                if (LicenseProgramId.HasValue)
                    src = src.Where(x => x.LicenseProgramId == LicenseProgramId);
                if (ObrazProgramId.HasValue)
                    src = src.Where(x => x.ObrazProgramId == ObrazProgramId);
                if (ProfileId.HasValue)
                    src = src.Where(x => x.ProfileId == ProfileId);

                var bind = src.Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.StudyFormId.ToString(), x.StudyFormName)).ToList();

                ComboServ.FillCombo(cbStudyForm, bind, false, true);
            }
        }
        private void FillComboStudyBasis()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.qEntry
                           where x.StudyLevelGroupId == MainClass.studyLevelGroupId && x.FacultyId == FacultyId
                           select new
                           {
                               x.LicenseProgramId,
                               x.ObrazProgramId,
                               x.ProfileId,
                               x.StudyFormId,
                               x.StudyBasisId,
                               x.StudyBasisName
                           }).Distinct();

                if (LicenseProgramId.HasValue)
                    src = src.Where(x => x.LicenseProgramId == LicenseProgramId);
                if (ObrazProgramId.HasValue)
                    src = src.Where(x => x.ObrazProgramId == ObrazProgramId);
                if (ProfileId.HasValue)
                    src = src.Where(x => x.ProfileId == ProfileId);
                if (StudyFormId.HasValue)
                    src = src.Where(x => x.StudyFormId == StudyFormId);

                var bind = src.Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.StudyBasisId.ToString(), x.StudyBasisName)).ToList();

                ComboServ.FillCombo(cbStudyBasis, bind, false, true);
            }
        }
        private void FillComboEgeExam()
        {
            string query = @"
                SELECT DISTINCT EgeExamName.Id, EgeExamName.Name
                FROM ed.Abiturient
                INNER JOIN ed.ExamInEntry ON ExamInEntry.EntryId = Abiturient.EntryId
                INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = ExamInEntry.ExamId
                INNER JOIN ed.EgeExamName ON EgeExamName.Id = EgeToExam.EgeExamNameId
                INNER JOIN ed.Entry ON Entry.Id = Abiturient.EntryId
                INNER JOIN ed.StudyLevel ON StudyLevel.Id = Entry.StudyLevelId
                WHERE StudyLevel.LevelGroupId = @LevelGroupId";

            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@LevelGroupId", MainClass.studyLevelGroupId);

            if (FacultyId.HasValue)
            {
                query += " AND Entry.FacultyId=@FacultyId ";
                sl.Add("@FacultyId", FacultyId);
            }
            if (LicenseProgramId.HasValue)
            {
                query += " AND Entry.LicenseProgramId=@LicenseProgramId ";
                sl.Add("@LicenseProgramId", LicenseProgramId);
            }
            if (ObrazProgramId.HasValue)
            {
                query += " AND Entry.ObrazProgramId=@ObrazProgramId ";
                sl.Add("@ObrazProgramId", ObrazProgramId);
            }
            if (StudyFormId.HasValue)
            {
                query += " AND Entry.StudyFormId=@StudyFormId ";
                sl.Add("@StudyFormId", StudyFormId);
            }
            if (StudyBasisId.HasValue)
            {
                query += " AND Entry.StudyBasisId=@StudyBasisId ";
                sl.Add("@StudyBasisId", StudyBasisId);
            }

            DataTable tbl = MainClass.Bdc.GetDataSet(query, sl).Tables[0];
            List<KeyValuePair<string, string>> bind =
                (from DataRow rw in tbl.Rows
                 select new
                     KeyValuePair<string, string>(rw.Field<int>("Id").ToString(), rw.Field<string>("Name"))
                ).Distinct().ToList();

            EgeExams =
                (from DataRow rw in tbl.Rows
                 select new
                 {
                     Id = rw.Field<int>("Id"),
                     Name = rw.Field<string>("Name")
                 }).ToDictionary(x => x.Id, x => x.Name);

            ComboServ.FillCombo(cbEgeExamName, bind, false, false);
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
            FillComboStudyForm();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyBasis();
        }
        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboEgeExam();
        }
        private void cbEgeExamName_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void chbEntered_CheckedChanged(object sender, EventArgs e)
        {
            FillGrid();
        }

        private void FillGrid()
        {
            string query = @"
                SELECT DISTINCT 
                hlpStatMaxApprovedEgeMarks.EgeExamNameId, Region.Name AS RegionName, 
                COUNT(DISTINCT Person.Id) AS EgeCNT, 
                MIN(hlpStatMaxApprovedEgeMarks.Value) AS EgeMin, 
                AVG(convert(float, hlpStatMaxApprovedEgeMarks.Value)) AS EgeAvg,
                MAX(hlpStatMaxApprovedEgeMarks.Value) AS EgeMax
                FROM ed.hlpStatMaxApprovedEgeMarks
                INNER JOIN ed.Person ON Person.Id = hlpStatMaxApprovedEgeMarks.PersonId
                INNER JOIN ed.Region ON Region.Id = Person.RegionId
                INNER JOIN ed.Abiturient ON Abiturient.PersonId = Person.Id
                INNER JOIN ed.Entry ON Entry.Id = Abiturient.EntryId
                INNER JOIN ed.StudyLevel ON StudyLevel.Id = Entry.StudyLevelId
                LEFT JOIN ed.extEntryView ON extEntryView.AbiturientId = Abiturient.Id
                WHERE StudyLevel.LevelGroupId = @LevelGroupId ";

            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@LevelGroupId", MainClass.studyLevelGroupId);

            if (FacultyId.HasValue)
            {
                query += " AND Entry.FacultyId=@FacultyId ";
                sl.Add("@FacultyId", FacultyId);
            }
            if (LicenseProgramId.HasValue)
            {
                query += " AND Entry.LicenseProgramId=@LicenseProgramId ";
                sl.Add("@LicenseProgramId", LicenseProgramId);
            }
            if (ObrazProgramId.HasValue)
            {
                query += " AND Entry.ObrazProgramId=@ObrazProgramId ";
                sl.Add("@ObrazProgramId", ObrazProgramId);
            }
            if (StudyFormId.HasValue)
            {
                query += " AND Entry.StudyFormId=@StudyFormId ";
                sl.Add("@StudyFormId", StudyFormId);
            }
            if (StudyBasisId.HasValue)
            {
                query += " AND Entry.StudyBasisId=@StudyBasisId ";
                sl.Add("@StudyBasisId", StudyBasisId);
            }
            if (chbEntered.Checked)
                query += " AND extEntryView.AbiturientId IS NOT NULL ";

            string groupby = "GROUP BY hlpStatMaxApprovedEgeMarks.EgeExamNameId, Region.Name";

            NewWatch wc = new NewWatch(5);
            wc.Show();
            wc.SetText("Загрузка данных...");
            DataTable tbl = MainClass.Bdc.GetDataSet(query + groupby, sl).Tables[0];

            wc.PerformStep();
            dgv.DataSource = null;
            var data = from DataRow rw in tbl.Rows
                       select new
                       {
                           EgeExamNameId = rw.Field<int>("EgeExamNameId"),
                           RegionName = rw.Field<string>("RegionName"),
                           CNT = rw.Field<int?>("EgeCNT"),
                           EgeMin = rw.Field<int?>("EgeMin"),
                           EgeAvg = rw.Field<double?>("EgeAvg"),
                           EgeMax = rw.Field<int?>("EgeMax")
                       };

            wc.PerformStep();

            DataTable src_tbl = new DataTable();

            src_tbl.Columns.Add("Регион", typeof(string));
            src_tbl.Columns.Add("Кол-во абитуриентов", typeof(int));
            src_tbl.Columns.Add("Мин.балл", typeof(int));
            src_tbl.Columns.Add("Сред.балл", typeof(double));
            src_tbl.Columns.Add("Макс.балл", typeof(int));
            wc.SetText("Построение списка...");
            foreach (var region in data.Where(x => x.EgeExamNameId == EgeExamNameId).Select(x => x.RegionName).Distinct())
            {
                DataRow row = src_tbl.NewRow();

                row["Регион"] = region;

                var balls = data.Where(x => x.EgeExamNameId == EgeExamNameId && x.RegionName == region);

                row["Кол-во абитуриентов"] = balls.Select(x => x.CNT).DefaultIfEmpty(0).First();
                row["Мин.балл"] = balls.Select(x => x.EgeMin).DefaultIfEmpty(0).First();
                row["Сред.балл"] = Math.Round(balls.Select(x => x.EgeAvg).DefaultIfEmpty(0d).First().Value, 2);
                row["Макс.балл"] = balls.Select(x => x.EgeMax).DefaultIfEmpty(0).First();

                src_tbl.Rows.Add(row);
            }
            
            wc.PerformStep();
            dgv.DataSource = src_tbl;
            wc.Close();
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            try
            {
                RtfDocument doc = new RtfDocument(PaperSize.A4, PaperOrientation.Portrait, Lcid.Russian);
                RtfParagraph p = doc.addParagraph();
                p.Alignment = Align.Center;
                p.DefaultCharFormat.FontSize = 14.0f;
                p.DefaultCharFormat.FontStyle.addStyle(FontStyleFlag.Bold);
                p.Text = "Статистика баллов ЕГЭ по регионам";
                if (chbEntered.Checked)
                    p.Text += " среди зачисленных абитуриентов";
                
                doc.addParagraph().Text = "";
                p = doc.addParagraph();
                p.Alignment = Align.Left;
                if (FacultyId.HasValue)
                    p.Text = cbFaculty.Text;
                else
                    p.Text = "По всему университету";

                if (LicenseProgramId.HasValue)
                {
                    p = doc.addParagraph();
                    p.Alignment = Align.Left;
                    p.Text = "Направление: " + cbLicenseProgram.Text;
                }

                if (ObrazProgramId.HasValue)
                {
                    p = doc.addParagraph();
                    p.Alignment = Align.Left;
                    p.Text = "Образовательная программа: " + cbObrazProgram.Text;
                }

                p = doc.addParagraph();
                //p.LineSpacing = RtfConstants.MILLIMETERS_TO_POINTS
                if (StudyFormId.HasValue)
                    p.Text = "Форма обучения: " + cbStudyForm.Text;
                else
                    p.Text = "Все формы обучения";

                p = doc.addParagraph();
                //p.LineSpacing = RtfConstants.MILLIMETERS_TO_POINTS
                if (StudyFormId.HasValue)
                    p.Text = "Основа обучения: " + cbStudyBasis.Text;
                else
                    p.Text = "Все основы обучения";

                p = doc.addParagraph();
                //p.LineSpacing = RtfConstants.MILLIMETERS_TO_POINTS
                p.Text = "Предмет: " + cbEgeExamName.Text;

                doc.addParagraph().Text = "";
                
                RtfTable table = doc.addTable(dgv.Rows.Count + 1, dgv.ColumnCount);

                table.setInnerBorder(RtfWriter.BorderStyle.Single, 0.25f);
                table.setOuterBorder(RtfWriter.BorderStyle.Single, 0.25f);
                //заголовки
                for (int i = 0; i < dgv.ColumnCount; i++)
                {
                    table.cell(0, i).Alignment = Align.Center;
                    table.cell(0, i).addParagraph().Text = dgv.Columns[i].Name;
                }

                table.setColWidth(0, (float)(60.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(1, (float)(25.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(2, (float)(20.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(3, (float)(20.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(4, (float)(20.0m * RtfConstants.MILLIMETERS_TO_POINTS));

                DataView dv = new DataView((DataTable)dgv.DataSource);
                dv.Sort = "Кол-во абитуриентов DESC";
                int rowId = 1;
                foreach (DataRow rw in dv.ToTable().Rows)
                {
                    int colId = 0;
                    foreach (DataColumn col in dv.ToTable().Columns)
                    {
                        var par = table.cell(rowId, colId).addParagraph();
                        par.Text = rw[col.ColumnName].ToString();
                        par.Alignment = Align.Center;
                        colId++;
                    }
                    //запрет переноса строк
                    table.setRowKeepInSamePage(rowId++, false);
                }

                doc.save(MainClass.saveTempFolder + "/Report.rtf");
                System.Diagnostics.Process.Start(MainClass.saveTempFolder + "/Report.rtf");
            }
            catch (Exception ex)
            {
                WinFormsServ.Error(ex.Message);
            }
        }
    }
}
