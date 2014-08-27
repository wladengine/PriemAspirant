using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using BaseFormsLib;
using RtfWriter;
using PriemLib;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Priem
{
    public partial class AbitFacultyIntersection : BaseFormEx
    {
        public AbitFacultyIntersection()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            InitHandlers();
            FillComboFaculty();
            Dgv = dgv;

            //FillGrid();
        }

        private void InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);
        }

        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }

        public int FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty).Value; }
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

        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.qEntry
                           orderby x.FacultyId
                           select new
                           {
                               x.FacultyId,
                               x.FacultyName
                           }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).ToList();

                ComboServ.FillCombo(cbFaculty, src, false, false);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.qEntry
                           orderby x.LicenseProgramCode
                           where x.FacultyId == FacultyId && x.StudyLevelGroupId == MainClass.studyLevelGroupId
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

        private void FillGrid()
        {
            DataTable tbl = GetTableIntersect(FacultyId);

            dgv.DataSource = tbl;
            dgv.Columns["FacultyId"].Visible = false;
            dgv.Columns["Факультет"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgv.Columns["Количество людей"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        }

        private DataTable GetTableIntersect(int iFacultyId)
        {
            string query = string.Format(@"
                SELECT qq.FacultyId, qq.FacultyName AS Факультет, COUNT(DISTINCT PersonId) AS 'Количество людей'
                FROM ed.qAbitAll as qq
                WHERE PersonId IN
                (
                    SELECT DISTINCT q.PersonId
                    FROM ed.qAbitAll as q
                    WHERE PersonId IN
                    (
	                    SELECT PersonId
	                    FROM ed.qAbitAll
	                    WHERE qAbitAll.FacultyId <> q.FacultyId
	                    AND StudyLevelGroupId = @StudyLevelGroupId
                        AND qAbitAll.BackDoc = 0
                    )
                    AND StudyLevelGroupId = @StudyLevelGroupId
                    AND q.FacultyId = @FacultyId
                    AND q.BackDoc = 0
                    {0}
                )
                AND StudyLevelGroupId = @StudyLevelGroupId
                AND FacultyId <> @FacultyId
                AND qq.BackDoc = 0
                GROUP BY qq.FacultyId, qq.FacultyName", LicenseProgramId == null ? "" : "AND q.LicenseProgramId=@LicenseProgramId ");
            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@FacultyId", iFacultyId);
            sl.Add("@StudyLevelGroupId", MainClass.studyLevelGroupId);
            if (LicenseProgramId != null)
                sl.Add("@LicenseProgramId", LicenseProgramId);
            
            return MainClass.Bdc.GetDataSet(query, sl).Tables[0];
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            int iOtherFacId = (int)dgv.Rows[e.RowIndex].Cells["FacultyId"].Value;
            new AbitFacultyIntersectionPersons(FacultyId, iOtherFacId).Show();
        }

        private void btnPrintMatrix_Click(object sender, EventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var facs = context.SP_Faculty.Select(x => new {x.Id, x.Name, x.Acronym }).ToList();

                DataTable dt = new DataTable();
                dt.Columns.Add("Fac");
                foreach (string s in facs.Select(x => x.Name))
                    dt.Columns.Add(s);

                Watch wc = new Watch(facs.Count * 2);
                wc.Show();

                foreach (var f in facs)
                {
                    DataRow rw = dt.NewRow();
                    rw["Fac"] = f.Name;

                    DataTable tblData = GetTableIntersect(f.Id);
                    var data =
                        from DataRow row in tblData.Rows
                        select new
                            {
                                Id = row.Field<int>("FacultyId"),
                                FacultyName = row.Field<string>("Факультет"),
                                CNT = row.Field<int>("Количество людей")
                            };

                    foreach (var d in data)
                        rw[d.FacultyName] = d.CNT;

                    dt.Rows.Add(rw);
                    wc.PerformStep();
                }

                RtfDocument doc = new RtfDocument(PaperSize.A4, PaperOrientation.Landscape, Lcid.Russian);
                RtfTable table = doc.addTable(dt.Rows.Count + 1, dt.Columns.Count);
                table.cell(0, 0).addParagraph().Text = "";

                int m = 0;
                foreach (var f in facs)
                {
                    table.FillCell(0, ++m, f.Acronym, FontStyleFlag.Normal);
                }

                int i = 1;
                table.CellPadding = 1.2f;
                table.setInnerBorder(RtfWriter.BorderStyle.Single, 0.7f);
                table.setOuterBorder(RtfWriter.BorderStyle.Single, 1f);
                foreach (DataRow r in dt.Rows)
                {
                    table.setColWidth(0, 260f);
                    for (int j = 0; j < r.ItemArray.Count(); j++)
                    {
                        table.FillCell(i, j, i == j ? "-" : r[j].ToString(), j > 0 ? FontStyleFlag.Bold : FontStyleFlag.Normal);
                    }
                    i++;
                    wc.PerformStep();
                }

                wc.Close();
                string fname = MainClass.saveTempFolder + "/stat.rtf";
                doc.save(fname);
                System.Diagnostics.Process.Start(fname);
            }
        }
    }
}
