using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using WordOut;
using EducServLib;
using BaseFormsLib;
using PriemLib;

namespace Priem
{
    public partial class FormA : BaseForm
    {
        private DateTime _One = new DateTime(2013, 07, 8, 0, 0, 0);
        private DateTime _Two = new DateTime(2013, 07, 16, 0, 0, 0);
        private DateTime _Three = new DateTime(2013, 07, 26, 0, 0, 0);

        private DBPriem bdcInet = new DBPriem();
        NewWatch wc;

        public FormA()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            bdcInet.OpenDatabase(MainClass.connStringOnline);
            GetKC();
            try
            {
                wc = new NewWatch(DateTime.Now.Date > _One ? (DateTime.Now.Date > _Two ? 33 : 22) : 11);
                wc.Show();
                Fill_1();
                if (DateTime.Now.Date > _One)
                    Fill_2();
                if (DateTime.Now.Date > _Two)
                    Fill_3();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            wc.Close();
            wc = null;
            bdcInet.CloseDataBase();
        }

        private void GetKC()
        {
            string query = "SELECT SUM(KCP) AS KC FROM ed.qEntry WHERE StudyLevelGroupId=@SLGId AND StudyFormId='1'";
            MainClass.Bdc.GetValue(query, new SortedList<string, object>() { { "@SLGId", MainClass.studyLevelGroupId } });
            tbKC.Text = MainClass.Bdc.GetValue(query, new SortedList<string, object>() { { "@SLGId", MainClass.studyLevelGroupId } }).ToString();
        }

        #region Fills
        private void Fill_1()
        {
            using (PriemEntities context = new PriemEntities())
            {
                DateTime _day = _One;

                wc.SetText("подано заявлений всего");
                //подано заявлений всего
                //база "Приём"
                string query = "SELECT StudyBasisId, SUM(CNT) AS CNT FROM ed.hlpStatistics WHERE Date<=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
                DataTable tbl = MainClass.Bdc.GetDataSet(query,
                    new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];

                int iPriemCountB = context.hlpStatistics.Where(x => x.StudyBasisId == 1 && x.Date <= _day && x.StudyLevelGroupId == MainClass.studyLevelGroupId).Select(x => x.CNT).Sum() ?? 0; //tbl.Rows[0].Field<int>("CNT");
                int iPriemCountP = context.hlpStatistics.Where(x => x.StudyBasisId == 2 && x.Date <= _day && x.StudyLevelGroupId == MainClass.studyLevelGroupId).Select(x => x.CNT).Sum() ?? 0; //tbl.Rows[1].Field<int>("CNT");
                tb1_Budzh.Text = iPriemCountB.ToString();
                tb1_Platn.Text = iPriemCountP.ToString();
                tb1_FullCount.Text = (iPriemCountB + iPriemCountP).ToString();
                wc.PerformStep();

                wc.SetText("петербуржцев");
                //петербуржцев
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
    //WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.RegionId='1' GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];

                var spb = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId 
                    && x.DocInsertDate <= _day && x.Person.Person_Contacts.RegionId == 1);
                tb1_SpbBudzh.Text = spb.Where(x => x.Entry.StudyBasisId == 1).Count().ToString();
                tb1_SpbPlatn.Text = spb.Where(x => x.Entry.StudyBasisId == 2).Count().ToString();
                wc.PerformStep();

                var mens = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId 
                    && x.DocInsertDate <= _day && x.Person.Sex == true);
                wc.SetText("мужчин");
                //мужчин
                tb1_MaleBudzh.Text = mens.Where(x => x.Entry.StudyBasisId == 1).Count().ToString();
                tb1_MalePlatn.Text = mens.Where(x => x.Entry.StudyBasisId == 2).Count().ToString();
                wc.PerformStep();

                wc.SetText("Имеющих среднее(полное) среднее образование");
    //            //школьники
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
    //WHERE convert(date, DocInsertDate)<=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId='1' GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
    //            var School = from DataRow rw in tbl.Rows
    //                         select new
    //                         {
    //                             StudyBasisId = rw.Field<int>("StudyBasisId"),
    //                             CNT = rw.Field<int>("CNT")
    //                         };

                var School = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                    && x.DocInsertDate <= _day && x.Person.Person_EducationInfo != null && x.Person.Person_EducationInfo.SchoolTypeId == 1);
                tb1_SchoolBudzh.Text = School.Where(x => x.Entry.StudyBasisId == 1).Count().ToString();
                tb1_SchoolPlatn.Text = School.Where(x => x.Entry.StudyBasisId == 2).Count().ToString();
                wc.PerformStep();

    //            wc.SetText("Имеющих среднее и высшее профессиональное образование");
    //            //СПО+ВПО
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
    //WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId NOT IN ('1', '3') GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];

                var SPO_VPO = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                    && x.DocInsertDate <= _day && x.Person.Person_EducationInfo != null && x.Person.Person_EducationInfo.SchoolTypeId != 1 
                    && x.Person.Person_EducationInfo.SchoolTypeId != 3);
                tb1_ProfBudzh.Text = SPO_VPO.Where(x => x.Entry.StudyBasisId == 1).Count().ToString();
                tb1_ProfPlatn.Text = SPO_VPO.Where(x => x.Entry.StudyBasisId == 2).Count().ToString();
                wc.PerformStep();

                wc.SetText("Имеющих начальное профессиональное образование");
                //НПО
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
    //WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId='3' GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
                //var NPO = from DataRow rw in tbl.Rows
                //          select new
                //          {
                //              StudyBasisId = rw.Field<int>("StudyBasisId"),
                //              CNT = rw.Field<int>("CNT")
                //          };

                var NPO = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                    && x.DocInsertDate <= _day && x.Person.Person_EducationInfo != null && x.Person.Person_EducationInfo.SchoolTypeId == 3);
                tb1_NPOBudzh.Text = NPO.Where(x => x.Entry.StudyBasisId == 1).Count().ToString();
                tb1_NPOPlatn.Text = NPO.Where(x => x.Entry.StudyBasisId == 2).Count().ToString();
                wc.PerformStep();

                wc.SetText("Победителей и призёров олимпиад");
    //            //Олимпиадники
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient WHERE qAbiturient.Id IN 
    //(SELECT AbiturientId FROM ed.extOlympiadsAll WHERE OlympTypeId IN ('3', '4') AND OlympValueId IN ('5','6')) 
    //AND convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
    //            var Olymps = from DataRow rw in tbl.Rows
    //                         select new
    //                         {
    //                             StudyBasisId = rw.Field<int>("StudyBasisId"),
    //                             CNT = rw.Field<int>("CNT")
    //                         };

                var Olymps = context.Olympiads.Where(x => x.Abiturient.Entry.StudyFormId == 1 && x.Abiturient.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                    && x.Abiturient.DocInsertDate <= _day && (x.OlympTypeId == 3 || x.OlympTypeId == 4) && (x.OlympValueId == 5 || x.OlympValueId == 6));
                tb1_OlympBudzh.Text = Olymps.Where(x => x.Abiturient.Entry.StudyBasisId == 1).Count().ToString();
                tb1_OlympPlatn.Text = Olymps.Where(x => x.Abiturient.Entry.StudyBasisId == 2).Count().ToString();
                wc.PerformStep();

                wc.SetText("в/к");
                //в\к
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient WHERE qAbiturient.CompetitionId IN ('2', '7') 
    //AND convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
    //            var VK = from DataRow rw in tbl.Rows
    //                     select new
    //                     {
    //                         StudyBasisId = rw.Field<int>("StudyBasisId"),
    //                         CNT = rw.Field<int>("CNT")
    //                     };

                var VK = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                    && x.DocInsertDate <= _day && (x.CompetitionId == 2 || x.CompetitionId == 7));
                tb1_VKBudzh.Text = VK.Where(x => x.Entry.StudyBasisId == 1).Count().ToString();
                tb1_VKPlatn.Text = VK.Where(x => x.Entry.StudyBasisId == 2).Count().ToString();
                wc.PerformStep();

                wc.SetText("иностранцев");
                //иностранцы
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
    //WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.CountryId<>'1' GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];

                var Foreigners = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                    && x.DocInsertDate <= _day && x.Person.NationalityId != 1);
                int iPriemForeignersB = Foreigners.Where(x => x.Entry.StudyBasisId == 1).Count();
                int iPriemForeignersP = Foreigners.Where(x => x.Entry.StudyBasisId == 2).Count();
                tb1_ForeignBudzh.Text = iPriemForeignersB.ToString();
                tb1_ForeignPlatn.Text = iPriemForeignersP.ToString();
                wc.PerformStep();

                wc.SetText("Граждане б. СССР, кр. России");
                //USSR
    //            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
    //WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.CountryId NOT IN ('1', '6') GROUP BY StudyBasisId";
    //            tbl = MainClass.Bdc.GetDataSet(query,
    //                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];

                var USSR = context.Abiturient.Where(x => x.Entry.StudyFormId == 1 && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                    && x.DocInsertDate <= _day && x.Person.NationalityId != 1 && x.Person.NationalityId != 6);
                tb1_USSRBudzh.Text = USSR.Where(x => x.Entry.StudyBasisId == 1).Count().ToString();
                tb1_USSRPlatn.Text = USSR.Where(x => x.Entry.StudyBasisId == 2).Count().ToString();
            
                wc.PerformStep();

                wc.SetText("Макс. конкурс " + _One.ToShortDateString());
                //max concurs
                query = @"SELECT TOP 1
      t.[LicenseProgramCode] + ' ' + t.[LicenseProgramName] AS Name, convert(float,t.[_cnt]) / (convert(float,t.[_sum])) AS conk
      FROM 
    (
    SELECT     StudyLevelGroupId, LicenseProgramId, LicenseProgramCode, LicenseProgramName, StudyFormId, StudyBasisId, SUM(KCP) AS _sum,
                              (SELECT     COUNT(Id) AS Expr1
                                FROM          ed.qAbiturient
                                WHERE      (StudyLevelGroupId = ed.qEntry.StudyLevelGroupId) AND (LicenseProgramId = ed.qEntry.LicenseProgramId) AND 
                                                       (StudyFormId = ed.qEntry.StudyFormId) AND (StudyBasisId = ed.qEntry.StudyBasisId) AND convert(date, DocInsertDate) <=@Date) AS _cnt
    FROM         ed.qEntry
    GROUP BY StudyLevelGroupId, LicenseProgramId, LicenseProgramCode, LicenseProgramName, StudyFormId, StudyBasisId
    ) t
      WHERE t.[StudyLevelGroupId] = @SLGId AND t.StudyFormId = 1 AND t.StudyBasisId = 1 AND t._sum > 0
      ORDER BY conk DESC";
                tbl = MainClass.Bdc.GetDataSet(query,
                    new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
                tb1_MaxConcurs.Text = Math.Round(tbl.Rows[0].Field<double>("conk"), 2).ToString() + " " + tbl.Rows[0].Field<string>("Name");
                wc.PerformStep();

                tb1_.Text = DateTime.Now < _day ? DateTime.Now.ToShortDateString() : _day.ToShortDateString();
            }
        }
        private void Fill_2()
        {
            DateTime _day = _Two;

            wc.SetText("Подано заявлений всего");
            //подано заявлений всего
            //база "Приём"
            string query = "SELECT StudyBasisId, SUM(CNT) AS CNT FROM ed.hlpStatistics WHERE Date<=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
            DataTable tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            int iPriemCountB = tbl.Rows[0].Field<int>("CNT");
            int iPriemCountP = tbl.Rows[1].Field<int>("CNT");
            //онлайн-база иностранцев
            query = string.Format(@"SELECT Entry.StudyBasisId, COUNT(ForeignApplication.Id) AS CNT FROM ForeignApplication 
INNER JOIN Entry ON Entry.Id=ForeignApplication.EntryId 
WHERE Entry.StudyLevelId IN ({0}) AND convert(date, DateOfStart)<=@Date AND Entry.StudyFormId='1' GROUP BY Entry.StudyBasisId", MainClass.studyLevelGroupId == 1 ? "'16', '18'" : "17");
            tbl = bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@Date", _day } }).Tables[0];
            int iOnlineForeignersB = tbl.Rows[0].Field<int>("CNT");
            int iOnlineForeignersP = tbl.Rows.Count > 1 ? tbl.Rows[1].Field<int>("CNT") : 0;
            tb2_Budzh.Text = (iOnlineForeignersB + iPriemCountB).ToString();
            tb2_Platn.Text = (iPriemCountP + iOnlineForeignersP).ToString();
            tb2_FullCount.Text = (iPriemCountB + iPriemCountP + iOnlineForeignersB + iOnlineForeignersP).ToString();
            wc.PerformStep();

            wc.SetText("петербуржцев");
            //петербуржцев
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.RegionId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb2_SpbBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb2_SpbPlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("мужчин");
            //мужчин
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.Person ON Person.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND Person.Sex='True' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb2_MaleBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb2_MalePlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("Имеющих среднее(полное) среднее образование");
            //школьники
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var School = from DataRow rw in tbl.Rows
                         select new
                         {
                             StudyBasisId = rw.Field<int>("StudyBasisId"),
                             CNT = rw.Field<int>("CNT")
                         };
            tb2_SchoolBudzh.Text = School.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb2_SchoolPlatn.Text = School.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("Имеющих среднее и высшее профессиональное образование");
            //СПО+ВПО
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId NOT IN ('1', '3') GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb2_ProfBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb2_ProfPlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("Имеющих начальное профессиональное образование");
            //НПО
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON Person.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId='3' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var NPO = from DataRow rw in tbl.Rows
                      select new
                      {
                          StudyBasisId = rw.Field<int>("StudyBasisId"),
                          CNT = rw.Field<int>("CNT")
                      };
            tb2_NPOBudzh.Text = NPO.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb2_NPOPlatn.Text = NPO.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("Победителей и призёров олимпиад");
            //Олимпиадники
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient WHERE qAbiturient.Id IN 
(SELECT AbiturientId FROM ed.extOlympiadsAll WHERE OlympTypeId IN ('3', '4') AND OlympValueId IN ('5','6')) 
AND convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var Olymps = from DataRow rw in tbl.Rows
                         select new
                         {
                             StudyBasisId = rw.Field<int>("StudyBasisId"),
                             CNT = rw.Field<int>("CNT")
                         };
            tb2_OlympBudzh.Text = Olymps.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb2_OlympPlatn.Text = Olymps.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("в/к");
            //в\к
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient WHERE qAbiturient.CompetitionId IN ('2', '7') 
AND convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var VK = from DataRow rw in tbl.Rows
                     select new
                     {
                         StudyBasisId = rw.Field<int>("StudyBasisId"),
                         CNT = rw.Field<int>("CNT")
                     };
            tb2_VKBudzh.Text = VK.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb2_VKPlatn.Text = VK.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("Иностранцев");
            //иностранцы
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.Person ON Person.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.CountryId<>'1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];

            int iPriemForeignersB = tbl.Rows[0].Field<int>("CNT");
            int iPriemForeignersP = tbl.Rows.Count > 1 ? tbl.Rows[1].Field<int>("CNT") : 0;
            tb2_ForeignBudzh.Text = (iPriemForeignersB + iOnlineForeignersB).ToString();
            tb2_ForeignPlatn.Text = (iPriemForeignersP + iOnlineForeignersP).ToString();
            wc.PerformStep();

            wc.SetText("Граждане б. СССР, кр. России");
            //USSR
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.Person ON Person.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.CountryId NOT IN ('1', '6') GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb2_USSRBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb2_USSRPlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("Макс. конкурс " + _Two.ToShortDateString());
            //max concurs
            query = @"SELECT TOP 1
  t.[LicenseProgramCode] + ' ' + t.[LicenseProgramName] AS Name, convert(float,t.[_cnt]) / (convert(float,t.[_sum])) AS conk
  FROM 
(
SELECT     StudyLevelGroupId, LicenseProgramId, LicenseProgramCode, LicenseProgramName, StudyFormId, StudyBasisId, SUM(KCP) AS _sum,
                          (SELECT     COUNT(Id) AS Expr1
                            FROM          ed.qAbiturient
                            WHERE      (StudyLevelGroupId = ed.qEntry.StudyLevelGroupId) AND (LicenseProgramId = ed.qEntry.LicenseProgramId) AND 
                                                   (StudyFormId = ed.qEntry.StudyFormId) AND (StudyBasisId = ed.qEntry.StudyBasisId) AND convert(date, DocInsertDate) <=@Date) AS _cnt
FROM         ed.qEntry
GROUP BY StudyLevelGroupId, LicenseProgramId, LicenseProgramCode, LicenseProgramName, StudyFormId, StudyBasisId
) t
  WHERE t.[StudyLevelGroupId] = @SLGId AND t.StudyFormId = 1 AND t.StudyBasisId = 1
  ORDER BY conk DESC";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb2_MaxConcurs.Text = Math.Round(tbl.Rows[0].Field<double>("conk"), 2).ToString() + " " + tbl.Rows[0].Field<string>("Name");
            wc.PerformStep();

            tb2_.Text = DateTime.Now < _day ? DateTime.Now.ToShortDateString() : _day.ToShortDateString();
        }
        private void Fill_3()
        {
            DateTime _day = _Three;

            wc.SetText("подано заявлений всего");
            //подано заявлений всего
            //база "Приём"
            string query = "SELECT StudyBasisId, SUM(CNT) AS CNT FROM ed.hlpStatistics WHERE Date<=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
            DataTable tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            int iPriemCountB = tbl.Rows[0].Field<int>("CNT");
            int iPriemCountP = tbl.Rows[1].Field<int>("CNT");
            //онлайн-база иностранцев
            query = string.Format(@"SELECT Entry.StudyBasisId, COUNT(ForeignApplication.Id) AS CNT FROM ForeignApplication 
INNER JOIN Entry ON Entry.Id=ForeignApplication.EntryId 
WHERE Entry.StudyLevelId IN ({0}) AND convert(date, DateOfStart)<=@Date AND Entry.StudyFormId='1' GROUP BY Entry.StudyBasisId", MainClass.studyLevelGroupId == 1 ? "'16', '18'" : "17");
            tbl = bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@Date", _day } }).Tables[0];
            int iOnlineForeignersB = tbl.Rows[0].Field<int>("CNT");
            int iOnlineForeignersP = tbl.Rows.Count > 1 ? tbl.Rows[1].Field<int>("CNT") : 0;
            tb3_Budzh.Text = (iOnlineForeignersB + iPriemCountB).ToString();
            tb3_Platn.Text = (iPriemCountP + iOnlineForeignersP).ToString();
            tb3_FullCount.Text = (iPriemCountB + iPriemCountP + iOnlineForeignersB + iOnlineForeignersP).ToString();
            wc.PerformStep();

            wc.SetText("петербуржцев");
            //петербуржцев
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.RegionId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb3_SpbBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb3_SpbPlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("мужчин");
            //мужчин
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.Person ON Person.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND Person.Sex='True' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb3_MaleBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb3_MalePlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("Имеющих среднее(полное) среднее образование");
            //школьники
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var School = from DataRow rw in tbl.Rows
                         select new
                         {
                             StudyBasisId = rw.Field<int>("StudyBasisId"),
                             CNT = rw.Field<int>("CNT")
                         };
            tb3_SchoolBudzh.Text = School.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb3_SchoolPlatn.Text = School.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("Имеющих среднее и высшее профессиональное образование");
            //СПО+ВПО
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId NOT IN ('1', '3') GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb3_ProfBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb3_ProfPlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("Имеющих начальное профессиональное образование");
            //НПО
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.extPersonAspirantAll ON extPersonAspirantAll.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.SchoolTypeId='3' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var NPO = from DataRow rw in tbl.Rows
                      select new
                      {
                          StudyBasisId = rw.Field<int>("StudyBasisId"),
                          CNT = rw.Field<int>("CNT")
                      };
            tb3_NPOBudzh.Text = NPO.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb3_NPOPlatn.Text = NPO.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("Победителей и призёров олимпиад");
            //Олимпиадники
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient WHERE qAbiturient.Id IN 
(SELECT AbiturientId FROM ed.extOlympiadsAll WHERE OlympTypeId IN ('3', '4') AND OlympValueId IN ('5','6')) 
AND convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var Olymps = from DataRow rw in tbl.Rows
                         select new
                         {
                             StudyBasisId = rw.Field<int>("StudyBasisId"),
                             CNT = rw.Field<int>("CNT")
                         };
            tb3_OlympBudzh.Text = Olymps.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb3_OlympPlatn.Text = Olymps.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("в/к");
            //в\к
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient WHERE qAbiturient.CompetitionId IN ('2', '7') 
AND convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            var VK = from DataRow rw in tbl.Rows
                     select new
                     {
                         StudyBasisId = rw.Field<int>("StudyBasisId"),
                         CNT = rw.Field<int>("CNT")
                     };
            tb3_VKBudzh.Text = VK.Where(x => x.StudyBasisId == 1).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            tb3_VKPlatn.Text = VK.Where(x => x.StudyBasisId == 2).Select(x => x.CNT).DefaultIfEmpty(0).First().ToString();
            wc.PerformStep();

            wc.SetText("Иностранцев");
            //иностранцы
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.Person ON Person.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.CountryId<>'1' GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            int iPriemForeignersB = tbl.Rows[0].Field<int>("CNT");
            int iPriemForeignersP = tbl.Rows.Count > 1 ? tbl.Rows[1].Field<int>("CNT") : 0;
            tb3_ForeignBudzh.Text = (iPriemForeignersB + iOnlineForeignersB).ToString();
            tb3_ForeignPlatn.Text = (iPriemForeignersP + iOnlineForeignersP).ToString();
            wc.PerformStep();

            wc.SetText("Граждане б. СССР, кр. России");
            //USSR
            query = @"SELECT StudyBasisId, COUNT(qAbiturient.Id) AS CNT FROM ed.qAbiturient INNER JOIN ed.Person ON Person.Id=qAbiturient.PersonId 
WHERE convert(date, DocInsertDate) <=@Date AND StudyLevelGroupId=@SLGId AND StudyFormId='1' AND extPersonAspirantAll.CountryId NOT IN ('1', '6') GROUP BY StudyBasisId";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb3_USSRBudzh.Text = tbl.Rows[0].Field<int>("CNT").ToString();
            tb3_USSRPlatn.Text = tbl.Rows[1].Field<int>("CNT").ToString();
            wc.PerformStep();

            wc.SetText("Макс. конкурс " + _Three.ToShortDateString());
            //max concurs
            query = @"SELECT TOP 1
  t.[LicenseProgramCode] + ' ' + t.[LicenseProgramName] AS Name, convert(float,t.[_cnt]) / (convert(float,t.[_sum])) AS conk
  FROM 
(
SELECT     StudyLevelGroupId, LicenseProgramId, LicenseProgramCode, LicenseProgramName, StudyFormId, StudyBasisId, SUM(KCP) AS _sum,
                          (SELECT     COUNT(Id) AS Expr1
                            FROM          ed.qAbiturient
                            WHERE      (StudyLevelGroupId = ed.qEntry.StudyLevelGroupId) AND (LicenseProgramId = ed.qEntry.LicenseProgramId) AND 
                                                   (StudyFormId = ed.qEntry.StudyFormId) AND (StudyBasisId = ed.qEntry.StudyBasisId) AND convert(date, DocInsertDate) <=@Date) AS _cnt
FROM         ed.qEntry
GROUP BY StudyLevelGroupId, LicenseProgramId, LicenseProgramCode, LicenseProgramName, StudyFormId, StudyBasisId
) t
  WHERE t.[StudyLevelGroupId] = @SLGId AND t.StudyFormId = 1 AND t.StudyBasisId = 1
  ORDER BY conk DESC";
            tbl = MainClass.Bdc.GetDataSet(query,
                new SortedList<string, object>() { { "@Date", _day }, { "@SLGId", MainClass.studyLevelGroupId } }).Tables[0];
            tb3_MaxConcurs.Text = Math.Round(tbl.Rows[0].Field<double>("conk"), 2).ToString() + " " + tbl.Rows[0].Field<string>("Name");
            wc.PerformStep();

            tb3_.Text = DateTime.Now < _day ? DateTime.Now.ToShortDateString() : _day.ToShortDateString();
        }
        #endregion

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            bdcInet.OpenDatabase(MainClass.connStringOnline);
            wc = new NewWatch(11);
            wc.Show();
            try
            {
                if (DateTime.Now < _One)
                    Fill_1();
                else if (DateTime.Now.Date > _One && DateTime.Now < _Two)
                    Fill_2();
                else if (DateTime.Now.Date > _Two && DateTime.Now < _Three)
                    Fill_3();
                wc.Close();
            }
            catch { }
            wc.Close();
            wc = null;
            bdcInet.CloseDataBase();
        }
        private void btnWord_Click(object sender, EventArgs e)
        {
            try
            {
                //MessageBox.Show("NotRealizedYet");
                WordDoc doc = new WordDoc(MainClass.dirTemplates + "\\FormA.dot", true);

                doc.SetFields("VUZNAME", "");
                doc.SetFields("KCP", tbKC.Text);

                //First part
                doc.SetFields("CNT1", tb1_FullCount.Text);
                doc.SetFields("CNT1_B", tb1_Budzh.Text);
                doc.SetFields("CNT1_P", tb1_Platn.Text);
                doc.SetFields("CNT1_SPB_B", tb1_SpbBudzh.Text);
                doc.SetFields("CNT1_SPB_P", tb1_SpbPlatn.Text);
                doc.SetFields("CNT1_MALE_B", tb1_MaleBudzh.Text);
                doc.SetFields("CNT1_MALE_P", tb1_MalePlatn.Text);
                doc.SetFields("CNT1_SCHOOL_B", tb1_SchoolBudzh.Text);
                doc.SetFields("CNT1_SCHOOL_P", tb1_SchoolPlatn.Text);
                doc.SetFields("CNT1_PROF_B", tb1_ProfBudzh.Text);
                doc.SetFields("CNT1_PROF_P", tb1_ProfPlatn.Text);
                doc.SetFields("CNT1_NPO_B", tb1_NPOBudzh.Text);
                doc.SetFields("CNT1_NPO_P", tb1_NPOPlatn.Text);
                doc.SetFields("CNT1_OLYMP_B", tb1_OlympBudzh.Text);
                doc.SetFields("CNT1_OLYMP_P", tb1_OlympPlatn.Text);
                doc.SetFields("CNT1_VK_B", tb1_VKBudzh.Text);
                doc.SetFields("CNT1_VK_P", tb1_VKPlatn.Text);
                doc.SetFields("CNT1_FOREIGN_B", tb1_ForeignBudzh.Text);
                doc.SetFields("CNT1_FOREIGN_P", tb1_ForeignPlatn.Text);
                doc.SetFields("CNT1_USSR_B", tb1_USSRBudzh.Text);
                doc.SetFields("CNT1_USSR_P", tb1_USSRPlatn.Text);
                doc.SetFields("CNT1_MAX_CONCURS", tb1_MaxConcurs.Text);

                //Second part
                doc.SetFields("CNT2", tb2_FullCount.Text);
                doc.SetFields("CNT2_B", tb2_Budzh.Text);
                doc.SetFields("CNT2_P", tb2_Platn.Text);
                doc.SetFields("CNT2_SPB_B", tb2_SpbBudzh.Text);
                doc.SetFields("CNT2_SPB_P", tb2_SpbPlatn.Text);
                doc.SetFields("CNT2_MALE_B", tb2_MaleBudzh.Text);
                doc.SetFields("CNT2_MALE_P", tb2_MalePlatn.Text);
                doc.SetFields("CNT2_SCHOOL_B", tb2_SchoolBudzh.Text);
                doc.SetFields("CNT2_SCHOOL_P", tb2_SchoolPlatn.Text);
                doc.SetFields("CNT2_PROF_B", tb2_ProfBudzh.Text);
                doc.SetFields("CNT2_PROF_P", tb2_ProfPlatn.Text);
                doc.SetFields("CNT2_NPO_B", tb2_NPOBudzh.Text);
                doc.SetFields("CNT2_NPO_P", tb2_NPOPlatn.Text);
                doc.SetFields("CNT2_OLYMP_B", tb2_OlympBudzh.Text);
                doc.SetFields("CNT2_OLYMP_P", tb2_OlympPlatn.Text);
                doc.SetFields("CNT2_VK_B", tb2_VKBudzh.Text);
                doc.SetFields("CNT2_VK_P", tb2_VKPlatn.Text);
                doc.SetFields("CNT2_FOREIGN_B", tb2_ForeignBudzh.Text);
                doc.SetFields("CNT2_FOREIGN_P", tb2_ForeignPlatn.Text);
                doc.SetFields("CNT2_USSR_B", tb2_USSRBudzh.Text);
                doc.SetFields("CNT2_USSR_P", tb2_USSRPlatn.Text);
                doc.SetFields("CNT2_MAX_CONCURS", tb2_MaxConcurs.Text);

                //Third part
                doc.SetFields("CNT3", tb3_FullCount.Text);
                doc.SetFields("CNT3_B", tb3_Budzh.Text);
                doc.SetFields("CNT3_P", tb3_Platn.Text);
                doc.SetFields("CNT3_SPB_B", tb3_SpbBudzh.Text);
                doc.SetFields("CNT3_SPB_P", tb3_SpbPlatn.Text);
                doc.SetFields("CNT3_MALE_B", tb3_MaleBudzh.Text);
                doc.SetFields("CNT3_MALE_P", tb3_MalePlatn.Text);
                doc.SetFields("CNT3_SCHOOL_B", tb3_SchoolBudzh.Text);
                doc.SetFields("CNT3_SCHOOL_P", tb3_SchoolPlatn.Text);
                doc.SetFields("CNT3_PROF_B", tb3_ProfBudzh.Text);
                doc.SetFields("CNT3_PROF_P", tb3_ProfPlatn.Text);
                doc.SetFields("CNT3_NPO_B", tb3_NPOBudzh.Text);
                doc.SetFields("CNT3_NPO_P", tb3_NPOPlatn.Text);
                doc.SetFields("CNT3_OLYMP_B", tb3_OlympBudzh.Text);
                doc.SetFields("CNT3_OLYMP_P", tb3_OlympPlatn.Text);
                doc.SetFields("CNT3_VK_B", tb3_VKBudzh.Text);
                doc.SetFields("CNT3_VK_P", tb3_VKPlatn.Text);
                doc.SetFields("CNT3_FOREIGN_B", tb3_ForeignBudzh.Text);
                doc.SetFields("CNT3_FOREIGN_P", tb3_ForeignPlatn.Text);
                doc.SetFields("CNT3_USSR_B", tb3_USSRBudzh.Text);
                doc.SetFields("CNT3_USSR_P", tb3_USSRPlatn.Text);
                doc.SetFields("CNT3_MAX_CONCURS", tb3_MaxConcurs.Text);
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
