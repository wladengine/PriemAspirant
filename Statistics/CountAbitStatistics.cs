using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

using BaseFormsLib;
using EducServLib;
using PriemLib;
namespace Priem
{
    public partial class CountAbitStatistics : BaseForm
    {
        private int iStudyLevelGroupId;
        public int? FacultyId
        {
            get
            {
                if (cbFaculty.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbFaculty);
            }
            set
            {
                ComboServ.SetComboId(cbFaculty, value);
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
            set
            {
                ComboServ.SetComboId(cbStudyBasis, value);
            }
        }
        public int? RegionId
        {
            get
            {
                if (cbRegion.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbRegion);
            }
        }

        public CountAbitStatistics()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;

            iStudyLevelGroupId = MainClass.dbType == PriemType.Priem ? 1 : 2;
            DateTime priemStartDay = new DateTime(2012, 6, 20);
            dtpStart.Value = DateTime.Now.AddDays(-7).Date < priemStartDay ? priemStartDay : DateTime.Now.AddDays(-7).Date;
            dtpEnd.Value = DateTime.Now;
            
            FillCombos();
            FillGrid();

            this.dtpStart.ValueChanged += new System.EventHandler(this.dtpStart_ValueChanged);
            this.dtpEnd.ValueChanged += new System.EventHandler(this.dtpEnd_ValueChanged);
            this.cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            this.cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
            this.cbRegion.SelectedIndexChanged += new EventHandler(cbRegion_SelectedIndexChanged);
        }

        private void FillCombos()
        {
            FillComboFaculty();
            FillComboBasis();
            FillComboRegion();
        }
        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> bind =
                    (from x in context.qEntry
                     select new
                     {
                         x.FacultyId,
                         x.FacultyName
                     }
                    ).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).ToList();
                ComboServ.FillCombo(cbFaculty, bind, false, true);
            }
        }
        private void FillComboBasis()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src =
                    (from x in context.qEntry
                     select new
                     {
                         x.FacultyId,
                         x.StudyBasisId,
                         x.StudyBasisName
                     });
                if (FacultyId != null)
                    src = src.Where(x => x.FacultyId == FacultyId);
                List<KeyValuePair<string, string>> bind =
                    src.Select(x => new { x.StudyBasisId, x.StudyBasisName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.StudyBasisId.ToString(), x.StudyBasisName)).ToList();
                ComboServ.FillCombo(cbStudyBasis, bind, false, true);
            }
        }
        private void FillComboRegion()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from ab in context.qAbitAll
                            join p in context.Person
                            on ab.PersonId equals p.Id
                            where ab.StudyLevelGroupId == MainClass.studyLevelGroupId
                            && FacultyId.HasValue ? ab.FacultyId == FacultyId.Value : true
                            && StudyBasisId.HasValue ? ab.StudyBasisId == StudyBasisId.Value : true
                            && p.Person_Contacts.RegionId != null
                            select new
                            {
                                p.Person_Contacts.Region.Id,
                                p.Person_Contacts.Region.Name
                            }).Distinct().ToList().OrderBy(x => x.Id).Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbRegion, bind, false, true);
            }
        }

        private void FillGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                Watch wc = new Watch(300);
                wc.Show();

                var dates = (from ab in context.qAbitAll
                             where ab.DocInsertDate != null && ab.DocInsertDate > dtpStart.Value && ab.DocInsertDate <= dtpEnd.Value
                             && ab.StudyLevelGroupId == MainClass.studyLevelGroupId && (FacultyId.HasValue ? (ab.FacultyId == FacultyId.Value) : (true))
                            select ab.DocInsertDate.Value).ToList().Select(x => x.Date).Distinct().OrderBy(x => x);

                string query = string.Format(@"SELECT DISTINCT LicenseProgramId, LicenseProgramCode + ' ' + LicenseProgramName AS Profession, 
            ObrazProgramId, ObrazProgramCrypt + ' ' + ObrazProgramName AS ObrazProgram, ProfileId, ProfileName, convert(date, DocInsertDate) AS Date, COUNT(extAbitAspirant.Id) AS CNT
            FROM ed.extAbitAspirant
            INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id = extAbitAspirant.PersonId
            WHERE StudyLevelGroupId='{0}' AND DocInsertDate IS NOT NULL AND convert(date, DocInsertDate)>=@DateStart AND convert(date, DocInsertDate)<=@DateEnd {1} {2} {3}
            GROUP BY LicenseProgramId, LicenseProgramCode, LicenseProgramName, ObrazProgramId, ObrazProgramCrypt, ObrazProgramName, ProfileId, ProfileName, convert(date, DocInsertDate)", 
            iStudyLevelGroupId.ToString(), FacultyId == null ? "" : " AND FacultyId='" + FacultyId.ToString() + "' ",
            StudyBasisId == null ? "" : " AND StudyBasisId='" + StudyBasisId.ToString() + "' ", 
            RegionId.HasValue ? " AND Person.RegionId='" + RegionId.Value.ToString() + "' " : "");
                System.Data.DataTable tblStatRaw = MainClass.Bdc.GetDataSet(query,
                    new SortedList<string, object>() { { "@DateStart", dtpStart.Value }, { "@DateEnd", dtpEnd.Value } }).Tables[0];


                query = string.Format(@"SELECT LicenseProgramId, ObrazProgramId, ProfileId, SUM(KCP) AS KCP
        FROM ed.qEntry
        WHERE StudyLevelGroupId='{0}' {1} {2}
        GROUP BY LicenseProgramId, ObrazProgramId, ProfileId", iStudyLevelGroupId.ToString(),
                FacultyId == null ? "" : " AND FacultyId='" + FacultyId.ToString() + "' ",
                StudyBasisId == null ? "" : " AND StudyBasisId='" + StudyBasisId.ToString() + "' ");
                System.Data.DataTable tblKC = MainClass.Bdc.GetDataSet(query).Tables[0];

                query = string.Format(@"SELECT LicenseProgramId, ObrazProgramId, ProfileId, COUNT(extAbitAspirant.Id) AS CNT
        FROM ed.extAbitAspirant
        INNER JOIN ed.Person ON Person.Id = extAbitAspirant.PersonId
        WHERE StudyLevelGroupId='{0}' {1} {2} {3}
        GROUP BY LicenseProgramId, ObrazProgramId, ProfileId", 
                iStudyLevelGroupId.ToString(),
                FacultyId == null ? "" : " AND FacultyId='" + FacultyId.ToString() + "' ",
                StudyBasisId == null ? "" : " AND StudyBasisId='" + StudyBasisId.ToString() + "' ",
                RegionId.HasValue ? " AND Person.RegionId='" + RegionId.Value.ToString() + "' " : "");
                System.Data.DataTable tblSumAbit = MainClass.Bdc.GetDataSet(query).Tables[0];

                var SumAbit = from DataRow rw in tblSumAbit.Rows
                               select new
                               {
                                   LicenseProgramId = rw.Field<int>("LicenseProgramId"),
                                   ObrazProgramId = rw.Field<int>("ObrazProgramId"),
                                   ProfileId = rw.Field<Guid?>("ProfileId"),
                                   CNT = rw.Field<int>("CNT")
                               };

                var KCP_Data = from DataRow rw in tblKC.Rows
                               select new
                               {
                                   LicenseProgramId = rw.Field<int>("LicenseProgramId"),
                                   ObrazProgramId = rw.Field<int>("ObrazProgramId"),
                                   ProfileId = rw.Field<Guid?>("ProfileId"),
                                   KCP = rw.Field<int>("KCP")
                               };

                var _dataFULL = (from DataRow rw in tblStatRaw.Rows
                             select new
                             {
                                 LicenseProgramId = rw.Field<int>("LicenseProgramId"),
                                 LicenseProgramName = rw.Field<string>("Profession"),
                                 ObrazProgramId = rw.Field<int>("ObrazProgramId"),
                                 ObrazProgramName = rw.Field<string>("ObrazProgram"),
                                 ProfileId = string.IsNullOrEmpty(rw["ProfileId"].ToString()) ? (Guid?)null : rw.Field<Guid?>("ProfileId"),
                                 ProfileName = rw.Field<string>("ProfileName"),
                                 Date = rw.Field<DateTime>("Date"),
                                 CNT = rw.Field<int>("CNT")
                             }).ToList();

                System.Data.DataTable tblStat = new System.Data.DataTable();
                dgvStatGrid.DataSource = null;
                tblStat.Columns.Add(new DataColumn("Name"));
                tblStat.Columns.Add(new DataColumn("KCP"));
                tblStat.Columns.Add(new DataColumn("SUM"));
                foreach (var d in dates)
                {
                    tblStat.Columns.Add(new DataColumn(d.ToShortDateString()));
                }

                var lp_ids =
                    (from x in _dataFULL
                     select new { x.LicenseProgramId, x.LicenseProgramName } ).Distinct().OrderBy(x => x.LicenseProgramName);
                foreach (var lp in lp_ids)
                {
                    //создаём строку
                    DataRow rwLP = tblStat.NewRow();
                    rwLP.SetField<string>("Name", lp.LicenseProgramName);

                    int KCP = KCP_Data.Where(x => x.LicenseProgramId == lp.LicenseProgramId).Select(x => x.KCP).Sum();
                    rwLP.SetField<int>("KCP", KCP);

                    int AbSum = SumAbit.Where(x => x.LicenseProgramId == lp.LicenseProgramId).Select(x => x.CNT).Sum();
                    rwLP.SetField<int>("SUM", AbSum);
                    
                    foreach (var d in dates)
                    {
                        int cnt = (from x in _dataFULL
                                   where x.LicenseProgramId == lp.LicenseProgramId && x.Date.Date == d
                                   select x.CNT).DefaultIfEmpty(0).Sum();
                        rwLP.SetField<int>(d.ToShortDateString(), cnt);
                    }
                    tblStat.Rows.Add(rwLP);
                    //собираем вложенные образовательные программы
                    var op_lst = (from x in _dataFULL
                                  where x.LicenseProgramId == lp.LicenseProgramId
                                  select new { x.ObrazProgramId, x.ObrazProgramName }).Distinct().OrderBy(x => x.ObrazProgramName);
                    foreach (var op in op_lst)
                    {
                        //создаём строку
                        DataRow rwOP = tblStat.NewRow();
                        rwOP.SetField<string>("Name", "     " + op.ObrazProgramName);

                        KCP = KCP_Data.Where(x => x.LicenseProgramId == lp.LicenseProgramId && x.ObrazProgramId == op.ObrazProgramId).Select(x => x.KCP).Sum();
                        rwOP.SetField<int>("KCP", KCP);

                        AbSum = SumAbit.Where(x => x.LicenseProgramId == lp.LicenseProgramId && x.ObrazProgramId == op.ObrazProgramId).Select(x => x.CNT).Sum();
                        rwOP.SetField<int>("SUM", AbSum);

                        foreach (var d in dates)
                        {
                            int cnt = (from x in _dataFULL
                                       where x.LicenseProgramId == lp.LicenseProgramId && x.Date.Date == d && x.ObrazProgramId == op.ObrazProgramId
                                       select x.CNT).DefaultIfEmpty(0).Sum();
                            rwOP.SetField<int>(d.ToShortDateString(), cnt);
                        }
                        tblStat.Rows.Add(rwOP);
                        var prof_lst = (from x in _dataFULL
                                        where x.LicenseProgramId == lp.LicenseProgramId && x.ObrazProgramId == op.ObrazProgramId && x.ProfileId != null
                                        select new { x.ProfileId, x.ProfileName }).Distinct().OrderBy(x => x.ProfileName);
                        foreach (var prof in prof_lst)
                        {
                            //создаём строку
                            DataRow rwProf = tblStat.NewRow();
                            rwProf.SetField<string>("Name", "           " + prof.ProfileName);
                            
                            KCP = KCP_Data.Where(x => x.LicenseProgramId == lp.LicenseProgramId && 
                                x.ObrazProgramId == op.ObrazProgramId && x.ProfileId == prof.ProfileId)
                                .Select(x => x.KCP).Sum();
                            rwProf.SetField<int>("KCP", KCP);

                            AbSum = SumAbit.Where(x => x.LicenseProgramId == lp.LicenseProgramId 
                                && x.ObrazProgramId == op.ObrazProgramId && x.ProfileId == prof.ProfileId).Select(x => x.CNT).Sum();
                            rwProf.SetField<int>("SUM", AbSum);

                            foreach (var d in dates)
                            {
                                int cnt = (from x in _dataFULL
                                           where x.LicenseProgramId == lp.LicenseProgramId && x.Date.Date == d 
                                           && x.ObrazProgramId == op.ObrazProgramId && x.ProfileId == prof.ProfileId
                                           select x.CNT).DefaultIfEmpty(0).Sum();
                                rwProf.SetField<int>(d.ToShortDateString(), cnt);
                            }
                            tblStat.Rows.Add(rwProf);
                            wc.PerformStep();
                        }
                        wc.PerformStep();
                    }
                    wc.PerformStep();
                }
                dgvStatGrid.DataSource = tblStat;
                dgvStatGrid.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                foreach (var d in dates)
                    dgvStatGrid.Columns[d.ToShortDateString()].Width = 63;

                dgvStatGrid.Columns["Name"].HeaderText = "Название";
                dgvStatGrid.Columns["KCP"].HeaderText = "КЦ";
                dgvStatGrid.Columns["KCP"].Width = 50;
                dgvStatGrid.Columns["SUM"].HeaderText = "Заявлений с начала приёма";
                dgvStatGrid.Columns["SUM"].Width = 63;

                dgvStatGrid.Columns["KCP"].Frozen = true;
                dgvStatGrid.Columns["Name"].Frozen = true;
                dgvStatGrid.Columns["SUM"].Frozen = true;

                dgvStatGrid.AllowUserToOrderColumns = false;

                wc.Close();
                wc = null;
            }
        }

        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboBasis();
        }
        private void cbRegion_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void dtpStart_ValueChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void dtpEnd_ValueChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void btnExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = string.Format("Динамика подачи заявлений с {0} по {1}", dtpStart.Value.ToShortDateString(), dtpEnd.Value.ToShortDateString());
            sfd.Filter = "Файлы Excel (.xls)|*.xls";
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                try
                {
                    List<string> columnList = new List<string>();
                    foreach (DataGridViewColumn column in dgvStatGrid.Columns)
                        if (column.Visible)
                            columnList.Add(column.Name);
                    Excel.Application exc = new Excel.Application();
                    Excel.Workbook wb = exc.Workbooks.Add(System.Reflection.Missing.Value);
                    Excel.Worksheet ws = (Excel.Worksheet)exc.ActiveSheet;

                    ws.Name = "Динамика подачи заявлений";

                    for (int j = 0; j < columnList.Count; j++)
                        ws.Cells[1, j + 1] = dgvStatGrid.Columns[columnList[j]].HeaderText;

                    //cellXY.NumberFormat = 0.00;

                    int i = 0;
                    // печать из грида
                    foreach (DataGridViewRow dgvr in dgvStatGrid.Rows)
                    {
                        for (int j = 0; j < columnList.Count; j++)
                            ws.Cells[i + 2, j + 1] = dgvStatGrid.Rows[i].Cells[columnList[j]].Value.ToString();

                        i++;
                    }

                    wb.SaveAs(sfd.FileName, Excel.XlFileFormat.xlExcel7,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        Excel.XlSaveAsAccessMode.xlExclusive,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value);
                    exc.Visible = true;
                    //по идее, из приложения Excel выходить не надо, пользователь в силах это сделать и самостоятельно, когда насмотрится на свой отсчёт
                    //exc.Quit();
                    //exc = null;
                }
                catch (System.Runtime.InteropServices.COMException exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
        }
    }
}