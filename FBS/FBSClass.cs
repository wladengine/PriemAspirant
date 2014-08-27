using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Data.SqlClient;

using EducServLib;
using PriemLib;
namespace Priem
{
    public static class FBSClass
    {        
        //создает запрос
        public static void MakeFBS(int iType)
        {
            string sQuery1 = @"SELECT ed.EgeCertificate.Number as nomer, ed.EgeCertificate.PrintNumber as tipograf,
                ed.EgeCertificate.Id,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=5) as rus,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=4) as mat,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=2) as fiz,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=8) as him,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=3) as bio,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=1) as ist,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=7) as geo,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=11) as ang,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=12) as nem,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=13) as fra,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=9) as obs,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=6) as lit,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=14) as isp,
                (select ed.EgeMark.value as mark FROM ed.EgeMark WHERE EgeCertificateId = ed.EgeCertificate.Id AND EgeExamNameId=10) as inf 
                 FROM ed.EgeMark inner join ed.EgeCertificate on ed.EgeMark.EgeCertificateId = ed.EgeCertificate.id                
                 WHERE ed.EgeCertificate.FBSStatusId IN (0,2) 
                 GROUP BY ed.EgeCertificate.Number , ed.EgeCertificate.PrintNumber,
                 ed.EgeCertificate.Id                                  
                 having COUNT(ed.EgeMark.id) > 1 "; //AND Person.NationalityId > 1";
            
            // уточнить как подгружать ЕГЭ
            string sQuery2 = @"SELECT ed.Person.Surname, ed.Person.Name, ed.Person.SecondName, ed.Person.PassportSeries, ed.Person.PassportNumber FROM ed.Person  
                               WHERE Person.Id IN (SELECT ed.qAbiturient.PersonId FROM ed.qAbiturient WHERE StudyLevelGroupId = 1 AND IsSecond = 0 AND IsReduced = 0 AND IsParallel = 0) 
                              AND Person.PassportSeries <> '' AND NOT EXISTS (SELECT Id FROM ed.EgeCertificate WHERE Year=2013 AND FBSStatusId IN (1,4) AND PersonId = Person.Id)";

            string sQuery3 = @"SELECT ed.Person.Surname, ed.Person.Name, ed.Person.SecondName, ed.Person.PassportSeries, ed.Person.PassportNumber FROM ed.Person  
                               WHERE Person.Id IN (SELECT ed.qAbiturient.PersonId FROM ed.qAbiturient WHERE StudyLevelGroupId = 1 AND IsSecond = 0 AND IsReduced = 0 AND IsParallel = 0) AND 
                               Person.PassportSeries <> '' AND EXISTS (SELECT Id FROM ed.EgeCertificate WHERE FBSStatusId IN (0,2) AND PersonId = Person.Id)";

//            string sQuery2 = @"SELECT ed.Person.Surname, ed.Person.Name, ed.Person.SecondName, ed.Person.PassportSeries, ed.Person.PassportNumber FROM ed.Person  
//                               WHERE Person.Id IN (SELECT ed.qAbiturient.PersonId FROM ed.qAbiturient WHERE StudyLevelGroupId = 1 AND IsSecond = 0 AND IsReduced = 0 AND IsParallel = 0) AND 
//                               Person.PassportSeries <> '' 
//                 AND EXISTS (SELECT ed.EgeCertificate.Id FROM ed.EgeMark inner join ed.EgeCertificate on ed.EgeMark.EgeCertificateId = ed.EgeCertificate.id                
//                 WHERE ed.EgeCertificate.FBSStatusId IN (0,2) AND PersonId = Person.Id
//                 GROUP BY ed.EgeCertificate.Id                                  
//                 having COUNT(ed.EgeMark.id) = 1)";
                      

            string sQuery = string.Empty;
            switch (iType)
            {
                case 1:
                    sQuery = sQuery1; break;
                case 2:
                    sQuery = sQuery2; break;
                case 3:
                    sQuery = sQuery3; break;               
            }

            StreamWriter sw = null;
            try
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    sw = new StreamWriter(saveFileDialog1.OpenFile(), Encoding.GetEncoding(1251));
                    DataSet ds = MainClass.Bdc.GetDataSet(sQuery);

                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        string s = string.Empty;

                        if (iType == 2 || iType == 3)//по ФИО
                        {
                            s = string.Format("{0}%{1}%{2}%{3}%{4}", dr["Surname"].ToString().Trim().ToUpper(), dr["Name"].ToString().Trim().ToUpper(), 
                                dr["SecondName"].ToString().Trim().ToUpper(), dr["PassportSeries"].ToString().Replace(" ", string.Empty), dr["PassportNumber"].ToString().Trim());
                        }
                        else if (iType == 1)//по номеру и баллам
                        {
                            s = string.Format("{0}%{1}%{2}%{3}%{4}%{5}%{6}%{7}%{8}%{9}%{10}%{11}%{12}%{13}%{14}", dr["nomer"], dr["rus"], dr["mat"], dr["fiz"], dr["him"], dr["bio"], dr["ist"], dr["geo"], dr["ang"], dr["nem"], dr["fra"], dr["obs"], dr["lit"], dr["isp"], dr["inf"]);
                        }
                        sw.WriteLine(s);
                    }
                    MessageBox.Show("Выполнено");
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при формировании запроса для ФБС: " + ex.Message);
            }
            finally
            {
                if (sw != null)
                    sw.Close();
            }
        }
    }
}
