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
    public partial class RegionAbitEGEMarksStatistics : Form
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
        public int RegionId
        {
            get
            {
                return ComboServ.GetComboIdInt(cbRegion).Value;
            }
            set
            {
                ComboServ.SetComboId(cbRegion, (int?)value);
            }
        }
        public int? EgeExamNameId
        {
            get
            {
                if (cbEgeExamName.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbEgeExamName);
            }
            set
            {

            }
        }
        public List<int> EgeExams { get; set; }

        public RegionAbitEGEMarksStatistics()
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
                           && LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId.Value : true
                           select new
                           {
                               x.LicenseProgramId,
                               x.ObrazProgramId,
                               x.ObrazProgramCrypt,
                               x.ObrazProgramName
                           }).Distinct();

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
                           && LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId.Value : true
                           && ObrazProgramId.HasValue ? x.ObrazProgramId == ObrazProgramId.Value : true
                           select new
                           {
                               x.StudyFormId,
                               x.StudyFormName
                           }).Distinct();

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
                           && LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId.Value : true
                           && ObrazProgramId.HasValue ? x.ObrazProgramId == ObrazProgramId.Value : true
                           && StudyFormId.HasValue ? x.StudyFormId == StudyFormId.Value : true
                           select new
                           {
                               x.StudyBasisId,
                               x.StudyBasisName
                           }).Distinct();

                var bind = src.Distinct().ToList()
                           .Select(x => new KeyValuePair<string, string>(x.StudyBasisId.ToString(), x.StudyBasisName)).ToList();

                ComboServ.FillCombo(cbStudyBasis, bind, false, true);
            }
        }
        private void FillComboRegion()
        {
            string query = @"
                SELECT DISTINCT Region.Id, Region.Name
                FROM ed.Abiturient
                INNER JOIN ed.Person ON Person.Id = Abiturient.PersonId
                INNER JOIN ed.Region ON Region.Id = Person.RegionId
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

            string tail = " ORDER BY Region.Id";

            DataTable tbl = MainClass.Bdc.GetDataSet(query + tail, sl).Tables[0];
            List<KeyValuePair<string, string>> bind =
                (from DataRow rw in tbl.Rows
                 select new
                     KeyValuePair<string, string>(rw.Field<int>("Id").ToString(), rw.Field<string>("Name"))
                ).Distinct().ToList();

            //EgeExams =
            //    (from DataRow rw in tbl.Rows
            //     select new
            //     {
            //         Id = rw.Field<int>("Id"),
            //         Name = rw.Field<string>("Name")
            //     }).ToDictionary(x => x.Id, x => x.Name);

            ComboServ.FillCombo(cbRegion, bind, false, false);
        }
        private void FillComboEgeExamName()
        {
            string query = @"
                SELECT DISTINCT EgeExamName.Id, EgeExamName.Name
                FROM ed.Abiturient
                INNER JOIN ed.Person ON Person.Id = Abiturient.PersonId
                INNER JOIN ed.ExamInEntry ON ExamInEntry.EntryId = Abiturient.EntryId
                INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = ExamInEntry.ExamId
                INNER JOIN ed.EgeExamName ON EgeExamName.Id = EgeToExam.EgeExamNameId
                INNER JOIN ed.Entry ON Entry.Id = Abiturient.EntryId
                INNER JOIN ed.StudyLevel ON StudyLevel.Id = Entry.StudyLevelId
                WHERE StudyLevel.LevelGroupId = @LevelGroupId AND Person.RegionId=@RegionId ";

            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@LevelGroupId", MainClass.studyLevelGroupId);
            sl.Add("@RegionId", RegionId);

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
                 select rw.Field<int>("Id")).ToList();

            
            ComboServ.FillCombo(cbEgeExamName, bind, false, true);
            
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
            FillComboRegion();
        }
        private void cbRegion_SelectedIndexChanged(object sender, EventArgs e)
        {
            int? oldVal = EgeExamNameId;
            FillComboEgeExamName();
            if (oldVal.HasValue && EgeExams.Contains(oldVal.Value))
                EgeExamNameId = oldVal;
        }
        private void chbEntered_CheckedChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void cbEgeExamName_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void chbEgeToExamOnly_CheckedChanged(object sender, EventArgs e)
        {
            FillGrid();
        }

        private void FillGrid()
        {
            string query = @"
                SELECT DISTINCT 
                EgeExamName.Name AS 'Экзамен', SP_Faculty.Name AS 'Факультет', 
                /*COUNT(DISTINCT Person.Id) AS EgeCNT, */
                MIN(hlpStatMaxApprovedEgeMarks.Value) AS 'Мин ЕГЭ', 
                AVG(convert(float, hlpStatMaxApprovedEgeMarks.Value)) AS 'Сред ЕГЭ',
                MAX(hlpStatMaxApprovedEgeMarks.Value) AS 'Макс ЕГЭ'
                FROM ed.hlpStatMaxApprovedEgeMarks
                INNER JOIN ed.Person ON Person.Id = hlpStatMaxApprovedEgeMarks.PersonId
                INNER JOIN ed.EgeExamName ON EgeExamName.Id = hlpStatMaxApprovedEgeMarks.EgeExamNameId
                INNER JOIN ed.Abiturient ON Abiturient.PersonId = Person.Id
                INNER JOIN ed.Entry ON Entry.Id = Abiturient.EntryId
                LEFT JOIN ed.ExamInEntry ON ExamInEntry.EntryId = Abiturient.EntryId
                LEFT JOIN ed.EgeToExam ON EgeToExam.ExamId = ExamInEntry.ExamId
                INNER JOIN ed.SP_Faculty ON SP_Faculty.Id = Entry.FacultyId
                INNER JOIN ed.StudyLevel ON StudyLevel.Id = Entry.StudyLevelId
                LEFT JOIN ed.extEntryView ON extEntryView.AbiturientId = Abiturient.Id
                WHERE StudyLevel.LevelGroupId=@LevelGroupId AND Person.RegionId=@RegionId ";

            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@LevelGroupId", MainClass.studyLevelGroupId);
            sl.Add("@RegionId", RegionId);

            if (FacultyId.HasValue)
            {
                query += " AND Entry.FacultyId=@FacultyId ";
                sl.Add("@FacultyId", FacultyId.Value);
            }
            if (LicenseProgramId.HasValue)
            {
                query += " AND Entry.LicenseProgramId=@LicenseProgramId ";
                sl.Add("@LicenseProgramId", LicenseProgramId.Value);
            }
            if (ObrazProgramId.HasValue)
            {
                query += " AND Entry.ObrazProgramId=@ObrazProgramId ";
                sl.Add("@ObrazProgramId", ObrazProgramId.Value);
            }
            if (StudyFormId.HasValue)
            {
                query += " AND Entry.StudyFormId=@StudyFormId ";
                sl.Add("@StudyFormId", StudyFormId.Value);
            }
            if (StudyBasisId.HasValue)
            {
                query += " AND Entry.StudyBasisId=@StudyBasisId ";
                sl.Add("@StudyBasisId", StudyBasisId.Value);
            }
            if (EgeExamNameId.HasValue)
            {
                query += " AND hlpStatMaxApprovedEgeMarks.EgeExamNameId=@EgeExamNameId ";
                sl.Add("@EgeExamNameId", EgeExamNameId.Value);
            }
            if (chbEntered.Checked)
                query += " AND extEntryView.AbiturientId IS NOT NULL ";
            if (chbEgeToExamOnly.Checked)
                query += " AND EgeToExam.EgeExamNameId IS NOT NULL AND EgeToExam.EgeExamNameId=hlpStatMaxApprovedEgeMarks.EgeExamNameId ";

            string groupby = "GROUP BY EgeExamName.Name, SP_Faculty.Name ORDER BY 2, 1";

            NewWatch wc = new NewWatch(5);
            wc.Show();
            wc.SetText("Загрузка данных...");
            DataTable tbl = MainClass.Bdc.GetDataSet(query + groupby, sl).Tables[0];

            wc.PerformStep();
            var col = dgv.SortedColumn != null ? dgv.SortedColumn.Name : null;
            var order = dgv.SortOrder;
            dgv.DataSource = null;

            foreach (DataRow rw in tbl.Rows)
                rw.SetField<double>("Сред ЕГЭ", Math.Round(rw.Field<double>("Сред ЕГЭ"), 2));

            wc.PerformStep();

            //dgv.DataSource = src_tbl;
            dgv.DataSource = tbl;
            if (col != null)
                dgv.Sort(dgv.Columns[col], order == SortOrder.Descending ? ListSortDirection.Descending : ListSortDirection.Ascending);
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
                p.Text = "Предмет: " + cbRegion.Text;

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

                table.setColWidth(0, (float)(37.5m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(1, (float)(92.5m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(2, (float)(12.5m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(3, (float)(12.5m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(4, (float)(12.5m * RtfConstants.MILLIMETERS_TO_POINTS));

                DataView dv = new DataView((DataTable)dgv.DataSource);
                dv.Sort = "Факультет, Экзамен";
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
