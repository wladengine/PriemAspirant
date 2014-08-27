using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WordOut;
using PriemLib;
namespace Priem
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            FillForm();
        }

        private void FillForm()
        {
            string query = "SELECT SUM(KCP) AS CNT FROM ed.qEntry WHERE StudyLevelGroupId=1 AND StudyFormId='1'";
            int iKCP_1k = (int)MainClass.Bdc.GetValue(query);
            tbKCP_1kurs.Text = iKCP_1k.ToString();

            //Подано заявлений на 1 курс, всего
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT 
FROM ed.qAbiturient WHERE qAbiturient.StudyLevelGroupId=1 AND StudyFormId='1' GROUP BY StudyBasisId";
            DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            tbCnt_Abit_1K_B.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tbCnt_Abit_1K_P.Text = tbl.Rows[1].Field<int>("CNT").ToString();

            //Подано заявлений на 1 курс, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT 
FROM ed.qAbiturient INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.RegionId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            tbCnt_Abit_1K_B_SPB.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tbCnt_Abit_1K_P_SPB.Text = tbl.Rows[1].Field<int>("CNT").ToString();

            //зачислено на 1 курс, всего
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var EntView = from DataRow rw in tbl.Rows
                              select new
                              {
                                  StudyBasisId = rw.Field<int>("StudyBasisId"),
                                  CNT = rw.Field<int>("CNT")
                              };
            int iCntStud1K_B = EntView.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            int iCntStud1K_P = EntView.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            
            tbCnt_Stud_1K_All.Text = (iCntStud1K_B + iCntStud1K_P).ToString();
            tbCnt_Stud_1K_B.Text = iCntStud1K_B.ToString();
            tbCnt_Stud_1K_P.Text = iCntStud1K_P.ToString();

            //зачислено на 1 курс, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT 
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.RegionId=1 GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var EntView_SPB = from DataRow rw in tbl.Rows
                              select new
                              {
                                  StudyBasisId = rw.Field<int>("StudyBasisId"),
                                  CNT = rw.Field<int>("CNT")
                              };
            int iCntStud1K_B_SPB = EntView_SPB.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            int iCntStud1K_P_SPB = EntView_SPB.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum();

            tbCnt_Stud_1K_B_SPB.Text = iCntStud1K_B_SPB.ToString();
            tbCnt_Stud_1K_P_SPB.Text = iCntStud1K_P_SPB.ToString();
            tbCnt_Stud_1K_All_SPB.Text = (iCntStud1K_B_SPB + iCntStud1K_P_SPB).ToString();

            //зачислено на 1 курс мужчин, всего
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.Sex='True' GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var Male = from DataRow rw in tbl.Rows
                       select new
                       {
                           StudyBasisId = rw.Field<int>("StudyBasisId"),
                           CNT = rw.Field<int>("CNT")
                       };
            tbCnt_Stud_1K_Male_B.Text = Male.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_Male_P.Text = Male.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //зачислено на 1 курс мужчин, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.Sex='True' AND extPersonAspirant.RegionId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var Male_SPB = from DataRow rw in tbl.Rows
                       select new
                       {
                           StudyBasisId = rw.Field<int>("StudyBasisId"),
                           CNT = rw.Field<int>("CNT")
                       };
            tbCnt_Stud_1K_Male_B_SPB.Text = Male_SPB.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_Male_P_SPB.Text = Male_SPB.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //имеющих среднее (полное) общее образование, всего
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.SchoolTypeId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var School = from DataRow rw in tbl.Rows
                         select new
                         {
                             StudyBasisId = rw.Field<int>("StudyBasisId"),
                             CNT = rw.Field<int>("CNT")
                         };
            int iCntStud1K_School_B_Priem = School.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            int iCntStud1K_School_P_Priem = School.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            tbCnt_Stud_1K_School_B.Text = iCntStud1K_School_B_Priem.ToString();
            tbCnt_Stud_1K_School_P.Text = iCntStud1K_School_P_Priem.ToString();

            //имеющих среднее (полное) общее образование, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.SchoolTypeId=1 AND extPersonAspirant.RegionId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var School_SPB = from DataRow rw in tbl.Rows
                         select new
                         {
                             StudyBasisId = rw.Field<int>("StudyBasisId"),
                             CNT = rw.Field<int>("CNT")
                         };

            tbCnt_Stud_1K_School_B_SPB.Text = School_SPB.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_School_P_SPB.Text = School_SPB.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //имеющих среднее и высшее профессиональное образование, всего
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.SchoolTypeId NOT IN (1,3)
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var Prof = from DataRow rw in tbl.Rows
                           select new
                           {
                               StudyBasisId = rw.Field<int>("StudyBasisId"),
                               CNT = rw.Field<int>("CNT")
                           };
            int iCntStud1K_Prof_B_Priem = Prof.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            int iCntStud1K_Prof_P_Priem = Prof.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            tbCnt_Stud_1K_Prof_B.Text = iCntStud1K_Prof_B_Priem.ToString();
            tbCnt_Stud_1K_Prof_P.Text = iCntStud1K_Prof_P_Priem.ToString();

            //имеющих среднее и высшее профессиональное образование, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.SchoolTypeId NOT IN (1,3) AND extPersonAspirant.RegionId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var Prof_SPB = from DataRow rw in tbl.Rows
                       select new
                        {
                            StudyBasisId = rw.Field<int>("StudyBasisId"),
                            CNT = rw.Field<int>("CNT")
                        };
            tbCnt_Stud_1K_Prof_B_SPB.Text = Prof_SPB.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_Prof_P_SPB.Text = Prof_SPB.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //имеющих НПО, всего
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.SchoolTypeId=3
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var NPOs = (from DataRow rw in tbl.Rows
                        select new
                        {
                            StudyBasisId = rw.Field<int>("StudyBasisId"),
                            CNT = rw.Field<int>("CNT")
                        });
            tbCnt_Stud_1K_NPO_B.Text = NPOs.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_NPO_P.Text = NPOs.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //имеющих НПО, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.SchoolTypeId=3 AND extPersonAspirant.RegionId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var NPO_SPB = (from DataRow rw in tbl.Rows
                        select new
                        {
                            StudyBasisId = rw.Field<int>("StudyBasisId"),
                            CNT = rw.Field<int>("CNT")
                        });
            tbCnt_Stud_1K_NPO_B_SPB.Text = NPO_SPB.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_NPO_P_SPB.Text = NPO_SPB.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //зачисленных в/к, всего
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND qAbiturient.CompetitionId IN (2,7) 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var VK = from DataRow rw in tbl.Rows
                     select new
                     {
                         StudyBasisId = rw.Field<int>("StudyBasisId"),
                         CNT = rw.Field<int>("CNT")
                     };
            tbCnt_Stud_1K_VK_B.Text = VK.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_VK_P.Text = VK.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //зачисленных в/к, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND qAbiturient.CompetitionId IN (2,7) AND extPersonAspirant.RegionId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var VK_SPB = from DataRow rw in tbl.Rows
                         select new
                         {
                             StudyBasisId = rw.Field<int>("StudyBasisId"),
                             CNT = rw.Field<int>("CNT")
                         };
            tbCnt_Stud_1K_VK_B_SPB.Text = VK_SPB.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_VK_P_SPB.Text = VK_SPB.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //зачисленных олимпиадников, всего
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT 
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.Olympiads ON Olympiads.AbiturientId = qAbiturient.Id 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND Olympiads.OlympValueId IN (5,6) AND Olympiads.OlympTypeId IN (3,4)
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var Olymp = from DataRow rw in tbl.Rows
                        select new
                        {
                            StudyBasisId = rw.Field<int>("StudyBasisId"),
                            CNT = rw.Field<int>("CNT")
                        };
            tbCnt_Stud_1K_Olymp_B.Text = Olymp.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_Olymp_P.Text = Olymp.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //зачисленных олимпиадников, СПб
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT 
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
INNER JOIN ed.Olympiads ON Olympiads.AbiturientId = qAbiturient.Id 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND Olympiads.OlympValueId IN (5,6) AND Olympiads.OlympTypeId IN (3,4) AND extPersonAspirant.RegionId=1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var Olymp_SPB = from DataRow rw in tbl.Rows
                        select new
                        {
                            StudyBasisId = rw.Field<int>("StudyBasisId"),
                            CNT = rw.Field<int>("CNT")
                        };
            tbCnt_Stud_1K_Olymp_B_SPB.Text = Olymp_SPB.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();
            tbCnt_Stud_1K_Olymp_P_SPB.Text = Olymp_SPB.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum().ToString();

            //иностранцев
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.CountryId <> 1 
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var Foreigners = from DataRow rw in tbl.Rows
                             select new
                             {
                                 StudyBasisId = rw.Field<int>("StudyBasisId"),
                                 CNT = rw.Field<int>("CNT")
                             };
            int iForeigners1K_B = Foreigners.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            int iForeigners1K_P = Foreigners.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum();

            tbCnt_Stud_1K_Foreign_B.Text = iForeigners1K_B.ToString();
            tbCnt_Stud_1K_Foreign_P.Text = iForeigners1K_P.ToString();

            //CCCP
            query = @"SELECT qAbiturient.StudyBasisId, COUNT(qAbiturient.Id) AS CNT
FROM ed.qAbiturient 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=qAbiturient.Id 
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=qAbiturient.PersonId 
WHERE qAbiturient.StudyLevelGroupId=1 AND qAbiturient.StudyFormId='1' AND extPersonAspirant.CountryId NOT IN (1,6)
GROUP BY qAbiturient.StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            var USSR = from DataRow rw in tbl.Rows
                             select new
                             {
                                 StudyBasisId = rw.Field<int>("StudyBasisId"),
                                 CNT = rw.Field<int>("CNT")
                             };
            int iUSSR1K_B = USSR.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).Sum();
            int iUSSR1K_P = USSR.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).Sum();

            tbCnt_Stud_1K_Foreign_B.Text = iUSSR1K_B.ToString();
            tbCnt_Stud_1K_Foreign_P.Text = iUSSR1K_P.ToString();

            //КЦ Магистратура
            query = "SELECT SUM(KCP) AS CNT FROM ed.qEntry WHERE StudyLevelGroupId=2 AND StudyFormId='1'";
            int iKCP_Mag = (int)MainClass.Bdc.GetValue(query);
            tbKCP_Mag.Text = iKCP_Mag.ToString();

            //Зачислено бюджет магистратура
            query = @"SELECT COUNT(extAbitAspirant.Id) 
FROM ed.extAbitAspirant 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=extAbitAspirant.Id 
WHERE extAbitAspirant.StudyLevelGroupId=2 AND extAbitAspirant.StudyFormId=1 AND extAbitAspirant.StudyBasisId=1";
            int iCNT_Mag = (int)MainClass.Bdc.GetValue(query);
            tbCnt_Stud_MAG_All.Text = iCNT_Mag.ToString();

            //Зачислено бюджет магистратура СПб
            query = @"SELECT COUNT(extAbitAspirant.Id) 
FROM ed.extAbitAspirant 
INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=extAbitAspirant.Id
INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=extAbitAspirant.PersonId
WHERE extAbitAspirant.StudyLevelGroupId=2 AND extAbitAspirant.StudyFormId=1 AND extAbitAspirant.StudyBasisId=1 AND extPersonAspirant.RegionId = 1";
            int iCNT_Mag_SPB = (int)MainClass.Bdc.GetValue(query);
            tbCnt_Stud_MAG_All_SPB.Text = iCNT_Mag_SPB.ToString();
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            try
            {
                WordDoc doc = new WordDoc(MainClass.dirTemplates + "\\Form2.dot", true);
                doc.SetFields("", tbKCP_1kurs.Text);
            }
            catch
            {

            }
        }
    }
}
