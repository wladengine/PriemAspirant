using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Data.Objects;
using System.Transactions;

using EducServLib;
using BDClassLib;
using WordOut;
using PriemLib;

namespace Priem
{
    public partial class LoadFBS : Form
    {
        const string TEMPLATE_MARKS = "Номер свидетельства%Типографский номер%Фамилия%Имя%Отчество%Серия документа%Номер документа%"+
            "Регион%Год%Статус"+
            "%Русский язык%Апелляция%Математика%Апелляция%Физика%Апелляция%Химия%Апелляция%Биология%Апелляция%История России%Апелляция%География%Апелляция%Английский язык%Апелляция%Немецкий язык%Апелляция%Французский язык%Апелляция%Обществознание%Апелляция%Литература%Апелляция%Испанский язык%Апелляция%Информатика%Апелляция%Проверок ВУЗами и их филиалами";
        const string TEMPLATE_NUMBER = @"Номер свидетельства%Типографский номер%Серия документа%Номер документа%Регион%Год%Статус%
            Русский язык%Апелляция%Математика%Апелляция%Физика%Апелляция%Химия%Апелляция%Биология%Апелляция%История России%Апелляция%География%Апелляция%
            Английский язык%Апелляция%Немецкий язык%Апелляция%Французский язык%Апелляция%Обществознание%Апелляция%Литература%Апелляция%Испанский язык%Апелляция%
            Информатика%Апелляция%Проверок ВУЗами и их филиалами";

        const int COLUMNS_NUMBER = 38;
        DateTime dtProtocol;
        DataTable dt;     
        DBPriem bdc;
        string yearBeforeSuffix = "-" + (DateTime.Now.AddYears(-1).Year % 2000).ToString();
        string yearNowSuffix = "-" + (DateTime.Now.Year % 2000).ToString();

        public LoadFBS()
        {
            InitializeComponent();           
            btnLoad.Enabled = true;
            rbEgeAnswerType1.CheckedChanged += new EventHandler(EnableLoadButton);
            rbEgeAnswerType2.CheckedChanged += new EventHandler(EnableLoadButton);
            bdc = MainClass.Bdc;
        }

        void EnableLoadButton(object sender, EventArgs e)
        {
            btnLoad.Enabled = true;
        }

        //открытие файла
        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsPasha())
                return;

            ParseAndAction();
        }

        //открытие файла и запуск
        private void ParseAndAction()
        {
            if (ofdFile.ShowDialog() == DialogResult.OK)
            {
                string filename = ofdFile.FileName;
                tbFile.Text = filename;

                FileInfo fi = new FileInfo(filename);
                dtProtocol = fi.CreationTime;

                StreamReader sr = null;

                if (MessageBox.Show("Начать загрузку данных?", "Внимание!", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
                
                try
                {
                    sr = new StreamReader(ofdFile.OpenFile(), Encoding.GetEncoding(1251));
                    string line = string.Empty;
                    string egenum = string.Empty;

                    dt = new DataTable();

                    line = sr.ReadLine();

                    //запрос по ФИО

                    if (rbEgeAnswerType1.Checked)
                        FBSAnswer1(sr);
                    if (rbEgeAnswerType2.Checked)
                        FBSAnswer2(sr);
                    
                    //теперь формат ответа ФБС одинаковый
                    /*
                    if (line.CompareTo(TEMPLATE_MARKS) == 0)
                    {                      
                        FBSAnswer1(sr);
                    }
                    else if (line.CompareTo(TEMPLATE_NUMBER) == 0)
                    {                       
                        FBSAnswer2(sr);
                    }
                    else
                        throw new Exception("Файл не соответствует формату ФБС");
                     */
                }
                catch (Exception ex)
                {
                    WinFormsServ.Error("Ошибка при загрузке файла ФБС, ответа на пакетный запрос:" + ex.Message);
                }
                finally
                {
                    sr.Close();
                }
            }

            MessageBox.Show("DONE!");
        }

        //build collection fbsnumber-egeexamid
        private SortedList<int, string> GetEgeSubjectsList()
        {
            SortedList<int, string> sl = new SortedList<int, string>();

            DataSet ds = bdc.GetDataSet("SELECT * FROM ed.EgeExamName");

            foreach (DataRow row in ds.Tables[0].Rows)
            {
                sl.Add((int)row["FBSnumber"], row["Id"].ToString());
            }

            return sl;
        }

        //reads fbs answer and updates ege cert status
        private void FBSAnswer1(StreamReader sr)
        {
            List<string> goodEGE = new List<string>();
            List<string> badEGE = new List<string>();
            SortedList<string, string> badEgeComment = new SortedList<string, string>();

            string line = string.Empty;
            
            //main loop
            while (!sr.EndOfStream)
            {
                //read line
                line = sr.ReadLine();

                //check string
                if (line.Length == 0)
                    continue;
                else if (line.StartsWith("Номер", StringComparison.InvariantCultureIgnoreCase))
                    continue;                
                
                //get ege cert number
                string egeNum = line.Substring(line.IndexOf('%') - 15, 15);

                //parse line and store in collection either good or bad
                if (line.ToLower().StartsWith("не найдено", StringComparison.InvariantCultureIgnoreCase))
                {
                    badEGE.Add(egeNum);
                    badEgeComment.Add(egeNum, "Не найдено");
                }
                else if (line.StartsWith("Аннулировано", StringComparison.InvariantCultureIgnoreCase))
                {
                    egeNum = line.Substring(line.IndexOf(":") - 15, 15);

                    badEGE.Add(egeNum);
                    badEgeComment.Add(egeNum, "Аннулировано");
                }
                else if (line.ToLower().Contains("истек срок"))
                {
                    badEGE.Add(egeNum);
                    badEgeComment.Add(egeNum, "Истек срок");
                } 
                else if (line.Contains(",0 ("))
                {
                    badEGE.Add(egeNum);
                    badEgeComment.Add(egeNum, "Ошибка в баллах");
                }                
                else
                {
                    goodEGE.Add(egeNum);
                }
            }

            //update status for ege certs
            using (PriemEntities context = new PriemEntities())
            {
                using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                {
                    try
                    {
                        foreach (string num in goodEGE)
                        {
                            Guid? egecertId = (from cert in context.EgeCertificate
                                               where cert.Number == num
                                               select cert.Id).FirstOrDefault();

                            context.EgeCertificate_UpdateFBSStatus(1, "", egecertId);
                        }

                        foreach (string num in badEGE)
                        {
                            Guid? egecertId = (from cert in context.EgeCertificate
                                               where cert.Number == num
                                               select cert.Id).FirstOrDefault();

                            context.EgeCertificate_UpdateFBSStatus(2, badEgeComment[num], egecertId);
                        
                        }

                        transaction.Complete();                          
                        
                    }
                    catch (Exception exc)
                    {
                        throw new Exception("Ошибка при сохранении данных: " + exc.Message);
                    }
                }
            }            
        }

        //reads fbs answer and saves ege certificates 2011 and 2012
        private void FBSAnswer2(StreamReader sr)
        {
            
                string line = string.Empty;
                string[] arr;
                char[] splitChars = { '%' };

                SortedList<int, string> slEges = GetEgeSubjectsList();
                //SortedList<string, EgeInstance> slCerts = new SortedList<string, EgeInstance>();

                List<ObjListItem> lCerts = new List<ObjListItem>();

                using (PriemEntities context = new PriemEntities())
                {
                    //main loop
                    while (!sr.EndOfStream)
                    {
                        //read line
                        line = sr.ReadLine();
                        try
                        {
                            //check strings when need to skip
                            if (line.Length == 0)
                                continue;
                            else if (line.ToLower().StartsWith("комментарий", StringComparison.InvariantCultureIgnoreCase))
                                continue;
                            else if (line.ToLower().StartsWith("не найдено", StringComparison.InvariantCultureIgnoreCase))
                                continue;
                            else if (line.ToLower().StartsWith("аннулировано", StringComparison.InvariantCultureIgnoreCase))
                                continue;
                            else
                            {
                                //get ege cert number
                                string egeNum = line.Substring(line.IndexOf('%') - 15, 15);

                                //skip year < 2012
                                if (!egeNum.EndsWith(yearBeforeSuffix) && !egeNum.EndsWith(yearNowSuffix))
                                    continue;

                                //split string
                                arr = line.Split(splitChars/*, COLUMNS_NUMBER*/);

                                //check ege cert status
                                if (arr[9].ToLower().CompareTo("действительно") != 0)
                                    continue;

                                //create ege
                                FBSEgeCert sert = new FBSEgeCert(arr[0], arr[1], arr[8]);

                                //get ege marks
                                int FBSnumber = 1;
                                for (int i = 10; i < 34; i = i + 2, FBSnumber++)
                                {
                                    if (arr[i].Length <= 0)
                                        continue;

                                    int mrk = (int)double.Parse(arr[i].Replace("Ошибка  (", "").Replace(",0)", "").Replace("!",""));
                                    if (mrk != 0)
                                        sert.AddMark(new FBSEgeMark(slEges[FBSnumber], mrk, arr[i + 1].CompareTo("0") != 0));
                                }

                                string pspSer = arr[5];
                                string pspNum = arr[6];

                                //get person id by document's data
                                Guid? personid = (from pers in context.Person
                                                  where pers.PassportSeries == pspSer && pers.PassportNumber == pspNum
                                                  select pers.Id).FirstOrDefault();

                                if (personid != null && personid.Value != Guid.Empty)
                                    lCerts.Add(new ObjListItem(personid, sert));
                            }

                        }
                        catch (Exception exc)
                        {
                            throw new Exception("Ошибка при сохранении данных: " + exc.Message);
                        }
                    }

                    //save to database
                    foreach (ObjListItem item in lCerts)
                    {
                        Guid person = (Guid)item.Key;
                        FBSEgeCert egecert = (FBSEgeCert)item.Value;
                        Guid certId;

                        //using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                        //{
                            try
                            {
                                int cnt = (from cert in context.EgeCertificate
                                           where cert.Number == egecert.Name && cert.Year == egecert.Year
                                           select cert).Count();

                                if (cnt > 0)
                                {
                                    var crt = (from cert in context.EgeCertificate
                                               where cert.Number == egecert.Name && cert.Year == egecert.Year
                                               select new { cert.Id, cert.FBSStatusId }).FirstOrDefault();

                                    if (crt.FBSStatusId == 1 || crt.FBSStatusId == 4)
                                        continue;

                                    certId = crt.Id;
                                    context.EgeMark_DeleteByCertId(certId);
                                }
                                else
                                {
                                    ObjectParameter entId = new ObjectParameter("id", typeof(Guid));
                                    context.EgeCertificate_Insert(egecert.Name, egecert.Tipograf, egecert.Year, person, "", true, entId);
                                    certId = (Guid)entId.Value;
                                }

                                foreach (FBSEgeMark mark in egecert.Marks)
                                {
                                    int examId = int.Parse(mark.ExamId);

                                    if (mark.Value > 0 && mark.Value < 101)
                                        context.EgeMark_Insert(mark.Value, examId, certId, mark.isApl, false);
                                }

                                context.EgeCertificate_UpdateFBSStatus(1, "", certId);

                                //transaction.Complete();
                            }
                            catch (Exception exc)
                            {
                                throw new Exception("Ошибка при сохранении данных: " + exc.Message);
                            }
                        //}
                    }
                }
                return;            
        } 
    }
}