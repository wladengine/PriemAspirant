using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization;
using System.Windows.Forms.DataVisualization.Charting;

using EducServLib;
using PriemLib;

namespace Priem
{
    public partial class RegionFacultyAbitCountStatistics : Form
    {
        private class StatVals
        {
            public string FacultyName { get; set; }
            public int CNT_Region { get; set; }
            public int CNT_Faculty { get; set; }
            public int MinSUM { get; set; }
            public double AvgSUM { get; set; }

        }

        /// <summary>
        /// Статистика - количество заявлений с одного региона пофакультетно
        /// </summary>
        public RegionFacultyAbitCountStatistics()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            if (!MainClass.IsOwner() || !MainClass.IsPasha())
            {
                MessageBox.Show("В разработке");
                this.Close();
                return;
            }
            FillComboRegion();
        }

        private void FillComboRegion()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var bind = (from x in context.Region
                           where x.RegionNumber != null
                           orderby x.RegionNumber
                           select new
                           {
                               x.Id,
                               x.Name
                           }).Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbRegion, bind, false, false);
            }
        }
        private void FillGrid()
        {
            string query = @"SELECT t.FacultyName AS 'Факультет', t.CNT_Reg AS 'Абитуриентов из региона', t.CNT_All AS 'Абитуриентов всего', 
convert(float, t.CNT_Reg)/convert(float, t.CNT_All) AS '% от общ', 
t.MinEGE AS 'Мин сумма ЕГЭ', t.AvgEGE AS 'Сред сумма ЕГЭ', t.MaxEGE AS 'Макс сумма ЕГЭ'
FROM 
(
SELECT DISTINCT FacultyName, COUNT(DISTINCT Person.Id) AS 'CNT_Reg', 
(SELECT COUNT(DISTINCT q.PersonId) FROM ed.qAbitAll AS q WHERE qAbitAll.FacultyName=q.FacultyName AND q.StudyLevelGroupId = 1) AS 'CNT_All',
MIN(hlpStatMarksSum.SUM) AS 'MinEGE', AVG(convert(float, hlpStatMarksSum.SUM)) AS 'AvgEGE', MAX(hlpStatMarksSum.SUM) AS 'MaxEGE'
FROM ed.qAbitAll
INNER JOIN ed.Person ON Person.Id = qAbitAll.PersonId
INNER JOIN ed.hlpStatMarksSum ON hlpStatMarksSum.AbiturientId = qAbitAll.Id
LEFT JOIN ed.extEntryView ON extEntryView.AbiturientId = qAbitAll.Id
WHERE qAbitAll.StudyLevelGroupId=@StudyLevelGroupId AND RegionId=@RegionId
";
            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@RegionId", ComboServ.GetComboIdInt(cbRegion));
            sl.Add("@StudyLevelGroupId", MainClass.studyLevelGroupId);

            if (cbEntered.Checked)
                query += " AND extEntryView.Id IS NOT NULL";

            string groupby = " GROUP BY FacultyName ) t ORDER BY [% от общ] DESC";

            DataTable tbl = MainClass.Bdc.GetDataSet(query + groupby, sl).Tables[0];
            foreach (DataRow rw in tbl.Rows)
            {
                rw.SetField<double>("Сред сумма ЕГЭ", Math.Round(rw.Field<double>("Сред сумма ЕГЭ"), 2));
                rw.SetField<double>("% от общ", Math.Round(rw.Field<double>("% от общ"), 2));
            }
            
            dgv.DataSource = tbl;
            dgv.Columns["Факультет"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //ReChart();
        }

        //private void ReChart()
        //{
        //    //chart.DataBindCrossTable((System.Collections.IEnumerable)dgv.DataSource, "FacultyName", "FacultyName", "CNT_Region, CNT_Faculty, MinSUM, AvgSUM, MaxSUM", "");
        //    //chart.DataBindTable((System.Collections.IEnumerable)dgv.DataSource, "FacultyName");
        //    //var chArea = new System.Windows.Forms.DataVisualization.Charting.ChartArea("Факультет");
        //    //chart.ChartAreas.Add(chArea);

        //    chart.DataSource = dgv.DataSource;

        //    List<string> lstFacs = (from DataRow rw in ((DataTable)chart.DataSource).Rows
        //                           select rw.Field<string>("FacultyName")).ToList();

        //    //if (!chart.Series.Select(x => x.Name).Contains("Ср.балл"))
        //    //    chart.Series.Add("Ср.балл");
        //    chart.Series[0].IsValueShownAsLabel = true;
        //    chart.Series[0].LegendToolTip = "FacultyName";
        //    chart.Series[0].XValueMember = "FacultyName";
        //    chart.Series[0].YValueMembers = "AvgSUM";
        //    //chart.Series[0].LegendToolTip = "FacultyName";
        //    //chart.Series["Series 1"].
        //    //chart.Series["Series 1"].Points.DataBindY((System.Collections.IEnumerable)dgv.DataSource, "CNT_Region");
        //}

        private void cbRegion_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void cbEntered_CheckedChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
    }
}
