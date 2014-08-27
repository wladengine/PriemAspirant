using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data;
using System.Data.Objects;
using WordOut;
using iTextSharp.text;
using iTextSharp.text.pdf;

using EducServLib;
using PriemLib;

namespace Priem
{
    public class Print
    {
        public static void PrintHostelDirection(Guid? persId, bool forPrint, string savePath)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extPersonAspirant person = (from per in context.extPersonAspirant
                                        where per.Id == persId
                                        select per).FirstOrDefault();                   

                    FileStream fileS = null;
                    using (FileStream fs = new FileStream(string.Format(@"{0}\HostelDirection.pdf", MainClass.dirTemplates), FileMode.Open, FileAccess.Read))
                    {

                        byte[] bytes = new byte[fs.Length];
                        fs.Read(bytes, 0, bytes.Length);
                        fs.Close();

                        PdfReader pdfRd = new PdfReader(bytes);

                        try
                        {
                            fileS = new FileStream(string.Format(savePath), FileMode.Create);
                        }
                        catch
                        {
                            if (fileS != null)
                                fileS.Dispose();
                            WinFormsServ.Error("Пожалуйста, закройте открытые файлы pdf");
                            return;
                        }


                        PdfStamper pdfStm = new PdfStamper(pdfRd, fileS);
                        pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "",
        PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING |
        PdfWriter.AllowPrinting);
                        AcroFields acrFlds = pdfStm.AcroFields;

                        acrFlds.SetField("Surname", person.Surname);
                        acrFlds.SetField("Name", person.Name);
                        acrFlds.SetField("LastName", person.SecondName);

                        acrFlds.SetField("Faculty", person.HostelFacultyAcr);
                        acrFlds.SetField("Nationality", person.NationalityName);
                        acrFlds.SetField("Country", person.CountryName);

                        acrFlds.SetField("Male", person.Sex ? "0" : "1");
                        acrFlds.SetField("Female", person.Sex ? "1" : "0");

                        pdfStm.FormFlattening = true;
                        pdfStm.Close();
                        pdfRd.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintExamPass(Guid? persId, string savePath, bool forPrint)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extPersonAspirant person = (from per in context.extPersonAspirant
                                        where per.Id == persId
                                        select per).FirstOrDefault();
                    
                    FileStream fileS = null;

                    using (FileStream fs = new FileStream(string.Format(@"{0}\ExamPass.pdf", MainClass.dirTemplates), FileMode.Open, FileAccess.Read))
                    {
                        byte[] bytes = new byte[fs.Length];
                        fs.Read(bytes, 0, bytes.Length);
                        fs.Close();

                        PdfReader pdfRd = new PdfReader(bytes);

                        try
                        {
                            fileS = new FileStream(string.Format(savePath), FileMode.Create);
                        }
                        catch
                        {
                            if (fileS != null)
                                fileS.Dispose();
                            WinFormsServ.Error("Пожалуйста, закройте открытые файлы pdf");
                            return;
                        }


                        PdfStamper pdfStm = new PdfStamper(pdfRd, fileS);
                        pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "",
        PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING |
        PdfWriter.AllowPrinting);
                        AcroFields acrFlds = pdfStm.AcroFields;

                        Barcode128 barcode = new Barcode128();
                        barcode.Code = person.PersonNum;

                        PdfContentByte cb = pdfStm.GetOverContent(1);

                        iTextSharp.text.Image img = barcode.CreateImageWithBarcode(cb, null, null);
                        img.SetAbsolutePosition(135, 565);
                        cb.AddImage(img);

                        acrFlds.SetField("Surname", person.Surname);
                        acrFlds.SetField("Name", person.Name);
                        acrFlds.SetField("LastName", person.SecondName);

                        acrFlds.SetField("Birth", person.BirthDate.ToShortDateString());
                        acrFlds.SetField("PassportSeries", person.PassportSeries + " " + person.PassportNumber);

                        acrFlds.SetField("chbMale", person.Sex ? "0" : "1");
                        acrFlds.SetField("chbFemale", person.Sex ? "1" : "0");


                        pdfStm.FormFlattening = true;
                        pdfStm.Close();
                        pdfRd.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintExamListWord(Guid? abitId, bool forPrint)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extAbitAspirant abit = (from ab in context.extAbitAspirant
                                    where ab.Id == abitId
                                    select ab).FirstOrDefault();

                    extPersonAspirant person = (from per in context.extPersonAspirant
                                        where per.Id == abit.PersonId
                                        select per).FirstOrDefault();                   

                    WordDoc wd = new WordDoc(string.Format(@"{0}\ExamSheet.dot", MainClass.dirTemplates), !forPrint);
                    TableDoc td = wd.Tables[0];

                    td[0, 0] = abit.FacultyName;
                    td[0, 1] = abit.LicenseProgramName;
                    td[0, 2] = abit.ProfileName;
                    td[1, 1] = MainClass.PriemYear;                   
                    td[1, 0] = abit.StudyBasisName.Substring(0, 1).ToUpper() + abit.StudyFormOldName.Substring(0, 1).ToUpper();
                    td[0, 10] = person.Surname;
                    td[0, 11] = person.Name;
                    td[0, 12] = person.SecondName;

                    td[2, 13] = abit.RegNum;
                    td[1, 14] = abit.FacultyAcr;    
                    td[1, 10] = person.PassportSeries + "   " + person.PassportNumber;

                    // экзамены!!! 
                    int row = 4;
                    IEnumerable<extExamInEntry> exams = from ex in context.extExamInEntry
                                                        where ex.EntryId == abit.EntryId
                                                        orderby ex.ExamName                     
                                                        select ex;

                    foreach (extExamInEntry ex in exams)
                    {
                        string sItem = ex.ExamName;
                        if (sItem.Contains("ностран") && MainClass.IsFilologFac())
                            sItem += string.Format(" ({0})", abit.LanguageName);

                        string mark = (from mrk in context.qMark
                                       where mrk.AbiturientId == abit.Id && mrk.ExamInEntryId == ex.Id
                                       select mrk.Value).FirstOrDefault().ToString();

                        td[0, row] = sItem;
                        td[1, row] = mark;
                        row++; 
                    }
                    
                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
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

        public static void PrintExamList(Guid? abitId, bool forPrint, string savePath)
        {
            FileStream fileS = null;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extAbitAspirant abit = (from ab in context.extAbitAspirant
                                    where ab.Id == abitId
                                    select ab).FirstOrDefault();

                    extPersonAspirant person = (from per in context.extPersonAspirant
                                        where per.Id == abit.PersonId
                                        select per).FirstOrDefault();                    

                    using (FileStream fs = new FileStream(string.Format(@"{0}\ExamList.pdf", MainClass.dirTemplates), FileMode.Open, FileAccess.Read))
                    {

                        byte[] bytes = new byte[fs.Length];
                        fs.Read(bytes, 0, bytes.Length);
                        fs.Close();

                        PdfReader pdfRd = new PdfReader(bytes);

                        try
                        {
                            fileS = new FileStream(string.Format(savePath), FileMode.Create);
                        }
                        catch
                        {
                            if (fileS != null)
                                fileS.Dispose();
                            WinFormsServ.Error("Пожалуйста, закройте открытые файлы pdf");
                            return;
                        }

                        PdfStamper pdfStm = new PdfStamper(pdfRd, fileS);
                        AcroFields acrFlds = pdfStm.AcroFields;

                        Barcode128 barcode = new Barcode128();
                        barcode.Code = abit.PersonNum + @"\" + abit.RegNum;

                        PdfContentByte cb = pdfStm.GetOverContent(1);

                        iTextSharp.text.Image img = barcode.CreateImageWithBarcode(cb, null, null);
                        img.SetAbsolutePosition(15, 65);
                        cb.AddImage(img);

                        acrFlds.SetField("Faculty", abit.FacultyName);
                        acrFlds.SetField("Profession", abit.LicenseProgramName);
                        acrFlds.SetField("Specialization", abit.ProfileName);
                        acrFlds.SetField("Year", MainClass.PriemYear);
                        acrFlds.SetField("Study", abit.StudyBasisName.Substring(0, 1).ToUpper() + abit.StudyFormOldName.Substring(0, 1).ToUpper());

                        acrFlds.SetField("Surname", person.Surname);
                        acrFlds.SetField("Name", person.Name);
                        acrFlds.SetField("SecondName", person.SecondName);
                        acrFlds.SetField("RegNumber", abit.RegNum);
                                               
                        acrFlds.SetField("FacultyAcr", abit.FacultyAcr);
                        acrFlds.SetField("Passport", person.PassportSeries + "   " + person.PassportNumber);

                        // экзамены!!! 
                        int i = 1;
                        IEnumerable<extExamInEntry> exams = from ex in context.extExamInEntry
                                                        where ex.EntryId == abit.EntryId
                                                        orderby ex.ExamName                     
                                                        select ex;

                        foreach (extExamInEntry ex in exams)
                        {
                            string sItem = ex.ExamName;
                            if (sItem.Contains("ностран") && MainClass.IsFilologFac())
                                sItem += string.Format(" ({0})", abit.LanguageName);

                            string mark = (from mrk in context.qMark
                                           where mrk.AbiturientId == abit.Id && mrk.ExamInEntryId == ex.Id
                                           select mrk.Value).FirstOrDefault().ToString();

                            acrFlds.SetField("Exam" + i, sItem);
                            acrFlds.SetField("Mark" + i, mark);
                            i++;
                        }
                        
                        pdfStm.FormFlattening = true;
                        pdfStm.Close();
                        pdfRd.Close();

                        fileS.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }

            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintSprav(Guid? abitId, bool forPrint)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extAbitAspirant abit = (from ab in context.extAbitAspirant
                                    where ab.Id == abitId
                                    select ab).FirstOrDefault();

                    extPersonAspirant person = (from per in context.extPersonAspirant
                                        where per.Id == abit.PersonId
                                        select per).FirstOrDefault();

                    WordDoc wd = new WordDoc(string.Format(@"{0}\Spravka.dot", MainClass.dirTemplates), !forPrint);
                    TableDoc td = wd.Tables[0];

                    string sFac;
                    string sForm;

                    if (abit.StudyFormId == 1)
                        sForm = "дневную форму обучения";
                    else if (abit.StudyFormId == 2)
                        sForm = "вечернюю форму обучения";
                    else
                        sForm = "заочную форму обучения";

                    wd.Fields["Section"].Text = sForm;

                    string vinFac = (from f in context.qFaculty
                                     where f.Id == abit.FacultyId
                                     select (f.VinName == null ? "на " + f.Name : f.VinName)).FirstOrDefault().ToLower();

                    wd.SetFields("Faculty", vinFac);
                    wd.SetFields("FIO", person.FIO);
                    wd.SetFields("Profession", abit.LicenseProgramName);

                    // оценки!!

                    IEnumerable<qMark> marks = from mrk in context.qMark
                                               where mrk.AbiturientId == abit.Id
                                               select mrk;
                   

                    string query = string.Format("SELECT qMark.Value, qMark.PassDate, extExamInProgram.ExamName as Name FROM (qMark INNER JOIN extExamInProgram ON qMark.ExamInProgramId = extExamInProgram.Id) INNER JOIN qAbiturient ON qMark.AbiturientId = qAbiturient.Id WHERE qAbiturient.Id = '{0}'", abitId);
                  
                    int i = 1;
                    foreach (qMark m in marks)
                    {
                        td[0, i] = i.ToString();
                        td[1, i] = m.ExamName;
                        td[2, i] = m.PassDate.Value.ToShortDateString();
                        if (m.Value == 0 || m.Value == 1)
                            td[3, i] = MarkClass.MarkProp(m.Value.ToString());
                        else
                            td[3, i] = m.Value.ToString();
                        td.AddRow(1);
                        i++;
                    }
                    td.DeleteLastRow();

                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
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

        public static void PrintStikerOne(Guid? abitId, bool forPrint)
        {
            string dotName;

            if (MainClass.dbType == PriemType.PriemAspirant)
                dotName = "StikerOneMag";
            else
                dotName = "StikerOne";

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    //AbiturientClass abit = AbiturientClass.GetInstanceFromDBForPrint(abitId);
                    var abit = context.extAbitAspirant.Where(x => x.Id == abitId).First();
                    //PersonClass person = PersonClass.GetInstanceFromDBForPrint(abit.PersonId);
                    var person = context.extPersonAspirant.Where(x => x.Id == abit.PersonId).First();
                    
                    
                    WordDoc wd = new WordDoc(string.Format(@"{0}\{1}.dot", MainClass.dirTemplates, dotName), !forPrint);

                    wd.SetFields("Faculty", abit.FacultyName);
                    wd.SetFields("Num", abit.PersonNum + @"\" + abit.RegNum);
                    wd.SetFields("Surname", person.Surname);
                    wd.SetFields("Name", person.Name);
                    wd.SetFields("SecondName", person.SecondName);
                    wd.SetFields("Profession", "("+ abit.LicenseProgramCode + ") " + abit.LicenseProgramName + ", " + abit.ObrazProgramName);
                    wd.SetFields("Specialization", abit.ProfileName);
                    wd.SetFields("Citizen", person.NationalityName);
                    wd.SetFields("Phone", person.Phone + "; " + person.Mobiles);
                    wd.SetFields("Email", person.Email);

                    for (int i = 1; i < 3; i++)
                    {
                        if (i != abit.StudyFormId)
                            wd.Shapes["StudyForm" + i].Delete();
                    }

                    for (int i = 1; i < 3; i++)
                    {
                        if (i != abit.StudyBasisId)
                            wd.Shapes["StudyBasis" + i].Delete();
                    }

                    wd.Shapes["Comp1"].Visible = false;
                    wd.Shapes["Comp2"].Visible = false;
                    wd.Shapes["Comp3"].Visible = false;
                    wd.Shapes["Comp4"].Visible = false;
                    wd.Shapes["Comp5"].Visible = false;
                    wd.Shapes["Comp6"].Visible = false;

                    wd.Shapes["Comp" + abit.CompetitionId.ToString()].Visible = true;

                    wd.Shapes["HasAssignToHostel"].Visible = person.HasAssignToHostel ?? false;

                    if (abit.CompetitionId == 6 && abit.OtherCompetitionId.HasValue)
                        wd.Shapes["Comp" + abit.CompetitionId.ToString()].Visible = true;

                    if (MainClass.dbType != PriemType.PriemAspirant)
                    {
                        string sPrevYear = DateTime.Now.AddYears(-1).Year.ToString();
                        string sCurrYear = DateTime.Now.Year.ToString();
                        string egePrevYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sPrevYear).Select(x => x.Number).FirstOrDefault();
                            //_bdc.GetStringValue(string.Format("SELECT TOP 1 EgeCertificate.Number FROM EgeCertificate WHERE EgeCertificate.Year = '{1}' AND PersonId = '{0}' ", abit.PersonId, DateTime.Now.Year - 1));
                        string egeCurYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sCurrYear).Select(x => x.Number).FirstOrDefault();
                            //_bdc.GetStringValue(string.Format("SELECT TOP 1 EgeCertificate.Number FROM EgeCertificate WHERE EgeCertificate.Year = '{1}' AND PersonId = '{0}' ", abit.PersonId, DateTime.Now.Year));

                        wd.SetFields("EgeNamePrevYear", egePrevYear);
                        wd.SetFields("EgeNameCurYear", egeCurYear);

                        int j = 1;

                        DataSet dsOlymps = MainClass.Bdc.GetDataSet(string.Format(@"
                                SELECT Olympiads.Id, OlympType.Name as Тип, OlympSubject.Name as Предмет, OlympValue.Id AS OlympValueId, 
                                OlympValue.Name as Степень FROM ed.Olympiads 
                                LEFT JOIN ed.OlympValue ON Olympiads.OlympValueId = OlympValue.Id 
                                LEFT JOIN ed.OlympSubject On OlympSubject.Id = Olympiads.OlympSubjectId 
                                LEFT JOIN ed.OlympType ON OlympType.Id=Olympiads.OlympTypeId 
                                WHERE Olympiads.AbiturientId = '{0}'", abitId));
                        foreach (DataRow dsRow in dsOlymps.Tables[0].Rows)
                        {
                            wd.SetFields("Level" + j, dsRow["Тип"].ToString());
                            wd.SetFields("Value" + j, dsRow["Степень"].ToString());
                            wd.SetFields("Subject" + j, dsRow["Предмет"].ToString());
                            j++;
                        }
                    }
                    else
                        if (person.DiplomSeries != "" || person.DiplomNum != "")
                            wd.SetFields("DocEduc", string.Format("диплом серия {0} № {1}", person.DiplomSeries, person.DiplomNum));

                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
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

        public static void PrintStikerAll(Guid? personId, Guid? abitId, bool forPrint)
        {
            string dotName;

            if (MainClass.dbType == PriemType.PriemAspirant)
                dotName = "StikerAllMag";
            else
                dotName = "StikerAll";

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    //PersonClass person = PersonClass.GetInstanceFromDBForPrint(personId);
                    var person = context.extPersonAspirant.Where(x => x.Id == personId).First();

                    WordDoc wd = new WordDoc(string.Format(@"{0}\{1}.dot", MainClass.dirTemplates, dotName), !forPrint);

                    //wd.SetFields("Faculty", _bdc.GetStringValue(string.Format("SELECT Faculty.Name FROM Faculty WHERE Faculty.Id = {0}", _bdc.GetFacultyId())));
                    wd.SetFields("Num", context.extAbitAspirant.Where(x => x.PersonId == person.Id).Select(x => x.PersonNum).First());
                    wd.SetFields("Surname", person.Surname);
                    wd.SetFields("Name", person.Name);
                    wd.SetFields("SecondName", person.SecondName);
                    wd.SetFields("Citizen", person.NationalityName);
                    wd.SetFields("Phone", person.Phone + "; " + person.Mobiles);
                    wd.SetFields("Email", person.Email);

                    wd.Shapes["Comp1"].Visible = false;
                    wd.Shapes["Comp2"].Visible = false;
                    wd.Shapes["Comp3"].Visible = false;
                    wd.Shapes["Comp4"].Visible = false;
                    wd.Shapes["Comp5"].Visible = false;
                    wd.Shapes["Comp6"].Visible = false;

                    wd.Shapes["HasAssignToHostel"].Visible = person.HasAssignToHostel ?? false;

                    if (MainClass.dbType != PriemType.PriemAspirant)
                    {
                        string sPrevYear = DateTime.Now.AddYears(-1).Year.ToString();
                        string sCurrYear = DateTime.Now.Year.ToString();
                        string egePrevYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sPrevYear).Select(x => x.Number).FirstOrDefault();
                        //_bdc.GetStringValue(string.Format("SELECT TOP 1 EgeCertificate.Number FROM EgeCertificate WHERE EgeCertificate.Year = '{1}' AND PersonId = '{0}' ", abit.PersonId, DateTime.Now.Year - 1));
                        string egeCurYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sCurrYear).Select(x => x.Number).FirstOrDefault();
                        //_bdc.GetStringValue(string.Format("SELECT TOP 1 EgeCertificate.Number FROM EgeCertificate WHERE EgeCertificate.Year = '{1}' AND PersonId = '{0}' ", abit.PersonId, DateTime.Now.Year));

                        wd.SetFields("EgeNamePrevYear", egePrevYear);
                        wd.SetFields("EgeNameCurYear", egeCurYear);

                        int j = 1;

                        DataSet dsOlymps = MainClass.Bdc.GetDataSet(string.Format(@"
                            SELECT Olympiads.Id, OlympType.Name as Тип, OlympSubject.Name as Предмет, OlympValue.Id AS OlympValueId, OlympValue.Name as Степень 
                            FROM ed.Olympiads 
                            LEFT JOIN ed.OlympValue ON Olympiads.OlympValueId = OlympValue.Id 
                            LEFT JOIN ed.OlympSubject On OlympSubject.Id = Olympiads.OlympSubjectId 
                            LEFT JOIN ed.OlympType ON OlympType.Id=Olympiads.OlympTypeId 
                            WHERE Olympiads.AbiturientId = '{0}'", abitId));
                        foreach (DataRow dsRow in dsOlymps.Tables[0].Rows)
                        {
                            wd.SetFields("Level" + j, dsRow["Тип"].ToString());
                            wd.SetFields("Value" + j, dsRow["Степень"].ToString());
                            wd.SetFields("Subject" + j, dsRow["Предмет"].ToString());
                            j++;
                        }
                    }
                    else
                        if (person.DiplomSeries != "" || person.DiplomNum != "")
                            wd.SetFields("DocEduc", string.Format("диплом серия {0} № {1}", person.DiplomSeries, person.DiplomNum));


                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
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
        public static void PrintApplication(Guid? abitId, bool forPrint, string savePath)
        {
            FileStream fileS = null;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    var PersonId = (from ab in context.Abiturient
                                    where ab.Id == abitId
                                  select ab.PersonId).FirstOrDefault();

                    var abitList = (from x in context.Abiturient
                                    join Entry in context.Entry on x.EntryId equals Entry.Id
                                    where Entry.StudyLevel.StudyLevelGroup.Id == 4
                                    && x.PersonId == PersonId
                                    && x.BackDoc == false
                                    select new
                                    {
                                        x.Id,
                                        x.PersonId,
                                        x.Barcode,
                                        Faculty = Entry.SP_Faculty.Name,
                                        Profession = Entry.SP_LicenseProgram.Name,
                                        ProfessionCode = Entry.SP_LicenseProgram.Code,
                                        ObrazProgram = Entry.StudyLevel.Acronym + "." + Entry.SP_ObrazProgram.Number + "." + MainClass.PriemYear + " " + Entry.SP_ObrazProgram.Name,
                                        Specialization = Entry.ProfileName,
                                        Entry.StudyFormId,
                                        Entry.StudyForm.Name,
                                        Entry.StudyBasisId,
                                        EntryType = (Entry.StudyLevelId == 17 ? 2 : 1),
                                        Entry.StudyLevelId,
                                        x.Priority,
                                        x.IsGosLine,
                                        Entry.CommissionId,
                                        ComissionAddress = Entry.CommissionId
                                    }).OrderBy(x => x.Priority).ToList();

                    var person = (from x in context.Person
                                  where x.Id == PersonId
                                  select new
                                  {
                                      x.Surname,
                                      x.Name,
                                      x.SecondName,
                                      x.Barcode,
                                      x.Person_AdditionalInfo.HostelAbit,
                                      x.BirthDate,
                                      BirthPlace = x.BirthPlace ?? "",
                                      Sex = x.Sex,
                                      ForeignNationality = x.ForeignNationalityId,
                                      Nationality = x.Nationality.Name,
                                      //NationalityId = x.NationalityId,
                                      Country = x.Person_Contacts.Country.Name,
                                      ForeignCountryName = x.Person_Contacts.ForeignCountry.Name,
                                      PassportType = x.PassportType.Name,
                                      x.PassportSeries,
                                      x.PassportNumber,
                                      x.PassportAuthor,
                                      x.PassportDate,
                                      x.Person_Contacts.City,
                                      Region = x.Person_Contacts.Region.Name,
                                      ProgramName = x.Person_EducationInfo.HEProfession,
                                      x.Person_Contacts.Code,
                                      x.Person_Contacts.Street,
                                      x.Person_Contacts.House,
                                      x.Person_Contacts.Korpus,
                                      x.Person_Contacts.Flat,
                                      x.Person_Contacts.Phone,
                                      x.Person_Contacts.Email,
                                      x.Person_Contacts.Mobiles,
                                      x.Person_EducationInfo.SchoolExitYear,
                                      x.Person_EducationInfo.SchoolName,
                                      AddInfo = x.Person_AdditionalInfo.ExtraInfo,
                                      Parents = x.Person_AdditionalInfo.PersonInfo,
                                      x.Person_EducationInfo.StartEnglish,
                                      x.Person_EducationInfo.EnglishMark,
                                      x.Person_EducationInfo.IsEqual,
                                      x.Person_EducationInfo.EqualDocumentNumber,
                                      CountryEduc = x.Person_EducationInfo.CountryEducId != null ? x.Person_EducationInfo.Country.Name : "",
                                      x.Person_EducationInfo.CountryEducId,
                                      x.Person_EducationInfo.ForeignCountryEducId,
                                      Qualification = x.Person_EducationInfo.HEQualification,
                                      x.Person_EducationInfo.SchoolTypeId,
                                      EducationDocumentSeries = x.Person_EducationInfo.DiplomSeries,
                                      EducationDocumentNumber = x.Person_EducationInfo.DiplomNum,
                                      x.Person_EducationInfo.AttestatRegion,
                                      x.Person_EducationInfo.AttestatSeries,
                                      x.Person_EducationInfo.AttestatNum,
                                      Language = x.Person_EducationInfo.Language.Name,
                                      HasPrivileges = (x.Person_AdditionalInfo.Privileges ?? 0) > 0,
                                      x.Person_EducationInfo.HasTRKI,
                                      x.Person_EducationInfo.TRKICertificateNumber,
                                      x.Person_AdditionalInfo.HostelEduc,
                                      IsRussia = (x.Person_Contacts.CountryId == 1),
                                      x.HasRussianNationality,
                                      x.Person_AdditionalInfo.Stag,
                                      x.Person_AdditionalInfo.WorkPlace,
                                      x.Num
                                  }).FirstOrDefault();

                    string tmp;
                    string dotName;
                    
                    dotName = "ApplicationAsp_2014";
                    

                    using (FileStream fs = new FileStream(string.Format(@"{0}\{1}.pdf", MainClass.dirTemplates, dotName), FileMode.Open, FileAccess.Read))
                    {

                        byte[] bytes = new byte[fs.Length];
                        fs.Read(bytes, 0, bytes.Length);
                        fs.Close();

                        PdfReader pdfRd = new PdfReader(bytes);

                        try
                        {
                            fileS = new FileStream(string.Format(savePath), FileMode.Create);
                        }
                        catch
                        {
                            if (fileS != null)
                                fileS.Dispose();
                            WinFormsServ.Error("Пожалуйста, закройте открытые файлы pdf");
                            return;
                        }


                        PdfStamper pdfStm = new PdfStamper(pdfRd, fileS);
                        pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "",
        PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING |
        PdfWriter.AllowPrinting);
                        AcroFields acrFlds = pdfStm.AcroFields;

                        //чей мы рисуем штрих-код??
                        //upd 03.06.2011 - уже не рисуем
                        //Barcode128 barcode = new Barcode128();
                        //barcode.Code = abit.RegNum;
                        //barcode.Code = "0008456";

                        PdfContentByte cb = pdfStm.GetOverContent(1);

                        //iTextSharp.text.Image img = barcode.CreateImageWithBarcode(cb, null, null);
                        //img.SetAbsolutePosition(420, 720);
                        //cb.AddImage(img);

                        //acrFlds.SetField("RegNum", abit.PersonNum + @"\" + abit.RegNum);

                        string FIO = ((person.Surname ?? "") + " " + (person.Name ?? "") + " " + (person.SecondName ?? "")).Trim();
                        acrFlds.SetField("FIO", FIO);

                        for (int ii = 0; ii < abitList.Count; ii++)
                        {
                            acrFlds.SetField("Priority" + (ii + 1).ToString(), abitList[ii].Priority.ToString());
                            acrFlds.SetField("Profession" + (ii + 1).ToString(), "(" + abitList[ii].ProfessionCode + ") " + abitList[ii].Profession);
                            acrFlds.SetField("Specialization" + (ii + 1).ToString(), abitList[ii].Specialization);
                            acrFlds.SetField("ObrazProgram" + (ii + 1).ToString(), abitList[ii].ObrazProgram);
                            acrFlds.SetField("StudyBasis" + abitList[ii].StudyBasisId.ToString() + (ii + 1).ToString(), "1");
                            acrFlds.SetField("StudyForm" + abitList[ii].StudyFormId.ToString() + (ii + 1).ToString(), "1");
                        }

                        //acrFlds.SetField("StudyForm1", "1");
                        acrFlds.SetField("ExitYear", person.SchoolExitYear.ToString());

                        string[] splitStr = GetSplittedStrings(person.SchoolName ?? "", 50, 70, 2);
                        for (int i = 1; i <= 2; i++)
                            acrFlds.SetField("School" + i, splitStr[i - 1]);

                        string attestat = (person.EducationDocumentSeries ?? "") + (person.EducationDocumentNumber ?? "");
                        string DiplomaNumber = !String.IsNullOrEmpty(person.EducationDocumentNumber) ? (" №" + person.EducationDocumentNumber) : "";
                        string DiplomaSeries = !String.IsNullOrEmpty(person.EducationDocumentSeries) ? ("серия " + person.EducationDocumentSeries + " ") : "";
                        acrFlds.SetField("Attestat", String.IsNullOrEmpty(attestat) ? "" : "диплом " + DiplomaSeries + DiplomaNumber);

                        acrFlds.SetField("HEProfession", person.ProgramName ?? "");
                        acrFlds.SetField("Qualification", person.Qualification);
                        if (person.SchoolTypeId.HasValue)
                        {
                            if ((person.SchoolTypeId != 4) || (person.SchoolTypeId == 4 && (person.Qualification).ToLower().IndexOf("аспирант") < 0))
                                acrFlds.SetField("NoEduc", "1");
                            else
                            {
                                acrFlds.SetField("HasEduc", "1");
                                acrFlds.SetField("HighEducation", person.SchoolName);
                            }
                        }

                        acrFlds.SetField("HostelEducYes", (person.HostelEduc) ? "1" : "0");
                        acrFlds.SetField("HostelEducNo", (person.HostelEduc) ? "0" : "1");
                        acrFlds.SetField("HostelAbitYes", (person.HostelAbit) ? "1" : "0");
                        acrFlds.SetField("HostelAbitNo", (person.HostelAbit) ? "0" : "1");
                        if (person.IsEqual && (person.ForeignCountryEducId.HasValue || (person.CountryEducId.HasValue && person.CountryEducId.Value != 193)))
                        {
                            acrFlds.SetField("IsEqual", "1");
                            acrFlds.SetField("EqualSertificateNumber", person.EqualDocumentNumber);
                        }
                        else
                        {
                            acrFlds.SetField("NoEqual", "1");
                        }

                        //дробилка даты и места рождения
                        tmp = person.BirthDate.ToShortDateString() + " " + person.BirthPlace;
                        string[] birthFieldsTmp = tmp.Split(' ');
                        string[] birthFields = new string[2];
                        int index = 0;
                        string strb = "";
                        foreach (string str in birthFieldsTmp)
                        {
                            if (index > 1)
                                break;
                            if (strb.Length + str.Length < 45 && index == 0 || strb.Length < 80 && index != 0)
                                strb += str + " ";
                            else
                            {
                                birthFields[index] = strb + str + " ";
                                index++;
                                strb = "";
                                continue;
                            }
                            birthFields[index] = strb;
                        }

                        acrFlds.SetField("BirthDateYear", person.BirthDate.Year.ToString("D2"));
                        acrFlds.SetField("BirthDateMonth", person.BirthDate.Month.ToString("D2"));
                        acrFlds.SetField("BirthDateDay", person.BirthDate.Day.ToString());
                        acrFlds.SetField("BirthPlace", person.BirthPlace);

                        acrFlds.SetField("Male", person.Sex ? "1" : "0");
                        acrFlds.SetField("Female", person.Sex ? "0" : "1");

                        //string Nationality = context.Country.Where(x => x.Id == person.NationalityId).Select(x => x.Name).FirstOrDefault();
                        if (person.Nationality.Contains("зарубеж"))
                        {
                            string ForeignNationality = context.ForeignCountry.Where(x => x.Id == person.ForeignNationality).Select(x => x.Name).FirstOrDefault();
                            acrFlds.SetField("Nationality", ForeignNationality);
                        }
                        else
                            acrFlds.SetField("Nationality", person.Nationality);

                        acrFlds.SetField("PassportSeries", person.PassportSeries);
                        acrFlds.SetField("PassportNumber", person.PassportNumber);

                        splitStr = GetSplittedStrings(person.PassportAuthor + " " + person.PassportDate.Value.ToString("dd.MM.yyyy"), 60, 70, 2);
                        for (int i = 1; i <= 2; i++)
                            acrFlds.SetField("PassportAuthor" + i, splitStr[i - 1]);

                        string country = person.Country;
                        string region = person.Region;
                        if (person.Country.Contains("зарубеж"))
                        {
                            country = person.ForeignCountryName;
                            region = "";
                        }

                        acrFlds.SetField("Address1", string.Format("{0} {1} {2}, {3}, ", person.Code, country, region, person.City));
                        acrFlds.SetField("Address2", string.Format("{0} дом {1} {2} кв. {3}", person.Street, person.House, (person.Korpus == string.Empty || person.Korpus == "-") ? "" : "корп. " + person.Korpus, person.Flat));

                        string addInfo = person.Mobiles.Replace('\r', ',').Replace('\n', ' ').Trim();//если начнут вбивать построчно, то хотя бы в одну строку сведём
                        if (addInfo.Length > 100)
                        {
                            int cutpos = 0;
                            cutpos = addInfo.Substring(0, 100).LastIndexOf(',');
                            addInfo = addInfo.Substring(0, cutpos) + "; ";
                        }


                        acrFlds.SetField("Phone", person.Phone);
                        acrFlds.SetField("Email", person.Email);
                        acrFlds.SetField("Mobiles", person.Mobiles);

                        acrFlds.SetField("Orig", "0");
                        acrFlds.SetField("Copy", "0");

                        string CountryEducName = context.Country.Where(x => x.Id == person.CountryEducId).Select(x => x.Name).FirstOrDefault();
                        string ForeignCountryEducName = context.ForeignCountry.Where(x => x.Id == person.ForeignCountryEducId).Select(x => x.Name).FirstOrDefault();

                        acrFlds.SetField("CountryEduc", CountryEducName ?? (ForeignCountryEducName ?? ""));
                        acrFlds.SetField("Language", person.Language ?? "");

                        if (person.Stag != string.Empty)
                        {
                            acrFlds.SetField("HasStag", "1");
                            acrFlds.SetField("Stag", person.Stag);
                            acrFlds.SetField("WorkPlace", person.WorkPlace);
                        }
                        else
                            acrFlds.SetField("NoStag", "1");

                        if (person.HasPrivileges)
                            acrFlds.SetField("HasPrivileges", "1");


                        // олимпиады
                        acrFlds.SetField("Extra", person.AddInfo ?? "");

                        //экстр. случаи
                        tmp = person.Parents.Replace('\r', ';').Replace('\n', ' ').Trim();
                        string[] mamaPapaWords = tmp.Split(' ');

                        string[] mamaPapa = new string[3];
                        strb = "";
                        index = 0;
                        foreach (string str in mamaPapaWords)
                        {
                            if (index >= 2)
                                break;
                            if (strb.Length + str.Length < 40 && index == 0 || strb.Length + str.Length < 80 && index != 0)
                                strb += str + " ";
                            else
                            {
                                mamaPapa[index] = strb + str + " ";
                                index++;
                                strb = "";
                                continue;
                            }
                            mamaPapa[index] = strb;
                        }
                        acrFlds.SetField("Parents1", mamaPapa[0]);
                        acrFlds.SetField("Parents2", mamaPapa[1]);
                        acrFlds.SetField("Parents3", mamaPapa[2]);

                        pdfStm.FormFlattening = true;
                        pdfStm.Close();
                        pdfRd.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }

            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintEnableProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                string query = 
                    string.Format(@"SELECT DISTINCT extAbitAspirant.Id as Id,
                                    extAbitAspirant.RegNum as Рег_Номер, extPersonAspirant.Surname + ' '+ extPersonAspirant.[Name] + ' ' + extPersonAspirant.SecondName as ФИО, 
                                    (case when extPersonAspirant.SchoolTypeId = 1 then extPersonAspirant.AttestatRegion + ' ' + extPersonAspirant.AttestatSeries + '  №' + extPersonAspirant.AttestatNum else extPersonAspirant.DiplomSeries + '  №' + extPersonAspirant.DiplomNum end) as Аттестат, 
                                    qEntry.LicenseProgramCode + ' ' + qEntry.LicenseProgramName + ', ' + qEntry.ObrazProgramName + ', ' + ( Case when qEntry.ProfileId IS NOT NULL then qEntry.ProfileName else '' end) as Направление,
                                    qEntry.LicenseProgramCode as Код, Competition.NAme as Конкурс, 
                                    extAbitAspirant.PersonId, extAbitAspirant.EntryId,
                                    (CASE WHEN extAbitAspirant.BackDoc > 0 THEN 'Забрал док.' ELSE (CASE WHEN extAbitAspirant.NotEnabled > 0 THEN 'Не допущен'ELSE '' END) END) as Примечания 
                                    FROM ((ed.extAbitAspirant 
                                    INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=extAbitAspirant.PersonId 
                                    INNER JOIN ed.qEntry ON qEntry.Id = extAbitAspirant.EntryId)
                                    LEFT JOIN ed.Competition ON Competition.Id = extAbitAspirant.CompetitionId) 
                                    LEFT JOIN ed.extProtocol ON extProtocol.AbiturientId = extAbitAspirant.Id ", MainClass.GetStringAbitNumber("qAbiturient"));

                string where = string.Format(" WHERE extProtocol.Id= '{0}' ", protocolId);
                string orderby = " ORDER BY Направление, Рег_Номер ";

                DataSet ds = MainClass.Bdc.GetDataSet(query + where + orderby);

                Guid ProtocolId = Guid.Parse(protocolId);

                using (PriemEntities context = new PriemEntities())
                {
                    var info = 
                        (from protocol in context.extEnableProtocol
                         join sf in context.StudyForm
                         on protocol.StudyFormId equals sf.Id
                         where protocol.Id == ProtocolId
                         select new
                         {
                             StudyFormName = sf.Name,
                             protocol.StudyBasisId,
                             protocol.Date,
                             protocol.Number
                         }).FirstOrDefault();

                    string basis = string.Empty;
                    switch (info.StudyBasisId)
                    {
                        case 1:
                            basis = "Бюджетные места";
                            break;
                        case 2:
                            basis = "Места по договорам с оплатой стоимости обучения";
                            break;
                    }

                    Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);

                    using (fileS = new FileStream(savePath, FileMode.Create))
                    {

                        BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font font = new Font(bfTimes, 10);

                        PdfWriter.GetInstance(document, fileS);
                        document.Open();

                        //HEADER
                        string header = string.Format(@"Форма обучения: {0}
    Условия обучения: {1}", info.StudyFormName, basis);

                        Paragraph p = new Paragraph(header, font);
                        document.Add(p);

                        float midStr = 13f;
                        p = new Paragraph(20f);
                        p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                        p.Add(new Phrase(info.Number, new Font(bfTimes, 18, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph(midStr);
                        p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
    о допуске к участию в конкурсе на основные образовательные программы ", new Font(bfTimes, 10, Font.BOLD)));

                        /*
                        p.Add(new Phrase(string.Format("{0} {1} {2}", "KODOKSO", "PROFESSION", "(SPECIALIZATION)"),
                            new Font(bfTimes, 10, Font.UNDERLINE + Font.BOLD)));*/
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        //date
                        p = new Paragraph(midStr);
                        p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(info.Date.HasValue ? info.Date.Value : DateTime.Now, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);


                        string spec = "";
                        PdfPTable curT = null;
                        int cnt = 0;
                        string currSpec = null;
                        string napravlenie = null;

                        foreach (DataRow row in ds.Tables[0].Rows)
                        {
                            cnt++;

                            currSpec = row.Field<string>("Направление");
                            //currSpec = abit.LicenseProgramCode ?? "" + " " + abit.LicenseProgramName ?? "";
                            string code = row.Field<string>("Код");
                            //string code = abit.LicenseProgramCode ?? "";
                            napravlenie = "направлению";

                            if (spec != currSpec)
                            {
                                spec = currSpec;
                                cnt = 1;

                                if (curT != null)
                                {
                                    document.Add(curT);
                                }

                                //Table

                                Table table = new Table(7);
                                table.Padding = 3;
                                table.Spacing = 0;
                                float[] headerwidths = { 5, 10, 30, 15, 20, 10, 10 };
                                table.Widths = headerwidths;
                                table.Width = 100;

                                PdfPTable t = new PdfPTable(7);
                                t.SetWidthPercentage(headerwidths, document.PageSize);
                                t.WidthPercentage = 100f;
                                t.SpacingBefore = 10f;
                                t.SpacingAfter = 10f;

                                t.HeaderRows = 2;

                                Phrase pra = new Phrase(string.Format("По {0} {1} ", napravlenie, currSpec), new Font(bfTimes, 10));

                                PdfPCell pcell = new PdfPCell(pra);
                                pcell.BorderWidth = 0;
                                pcell.Colspan = 7;
                                t.AddCell(pcell);

                                string[] headers = new string[]
                            {
                                "№ п/п",
                                "Рег.номер",
                                "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                                "Номер аттестата или диплома",
                                "Номер сертификата ЕГЭ по профильному предмету",
                                "Вид конкурса",
                                "Примечания"
                            };
                                foreach (string h in headers)
                                {
                                    PdfPCell cell = new PdfPCell();
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                    t.AddCell(cell);
                                }

                                curT = t;
                            }

                            string quer = string.Format(@"
                                    SELECT TOP 1 EgeCertificate.Number FROM ed.EgeCertificate 
                                    INNER JOIN ed.EgeMark ON EgeMark.EgeCertificateId= EgeCertificateId
                                    INNER JOIN ed.EgeToExam ON EgeToExam.EgeExamNameId = EgeMark.EgeExamNameId
                                    WHERE EgeCertificate.PersonId='{0}' AND EgeToExam.ExamId = 
                                    (SELECT TOP 1 ExamId FROM ed.ExamInEntry WHERE ExamInEntry.EntryId='{1}' AND IsProfil>0)",
                                        row["PersonId"].ToString(),
                                        row["EntryId"].ToString());

                            string egecert = MainClass.Bdc.GetStringValue(quer);

                            curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Рег_Номер"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("ФИО"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Аттестат"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(egecert, new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Конкурс"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Примечания"), new Font(bfTimes, 10)));
                        }

                        if (curT != null)
                        {
                            document.Add(curT);
                        }

                        //FOOTER
                        p = new Paragraph(30f);
                        p.KeepTogether = true;
                        p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ____________________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Заместитель начальника Управления по организации приема – советник проректора по направлениям___________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Ответственный секретарь комиссии по приему документов_______________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        document.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }

            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintDisEnableProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                string query = string.Format(@"SELECT DISTINCT extAbitAspirant.Id as Id,
                                    extAbitAspirant.RegNum as Рег_Номер, extPersonAspirant.Surname + ' '+extPersonAspirant.[Name] + ' ' + extPersonAspirant.SecondName as ФИО, 
                                    (case when extPersonAspirant.SchoolTypeId = 1 then extPersonAspirant.AttestatRegion + ' ' + extPersonAspirant.AttestatSeries + '  №' + extPersonAspirant.AttestatNum else extPersonAspirant.DiplomSeries + '  №' + extPersonAspirant.DiplomNum end) as Аттестат, 
                                    qEntry.LicenseProgramCode + ' ' + qEntry.LicenseProgramName + ', ' + qEntry.ObrazProgramName + ', ' + ( Case when qEntry.ProfileId IS NOT NULL then qEntry.ProfileName else '' end) as Направление,
                                    qEntry.LicenseProgramCode as Код, Competition.NAme as Конкурс, 
                                    extAbitAspirant.PersonId, extAbitAspirant.EntryId,
                                    (CASE WHEN extAbitAspirant.BackDoc > 0 THEN 'Забрал док.' ELSE (CASE WHEN extAbitAspirant.NotEnabled > 0 THEN 'Не допущен'ELSE '' END) END) as Примечания 
                                    FROM ((ed.extAbitAspirant 
                                    INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=extAbitAspirant.PersonId 
                                    INNER JOIN ed.qEntry ON qEntry.Id = extAbitAspirant.EntryId)
                                    LEFT JOIN ed.Competition ON Competition.Id = extAbitAspirant.CompetitionId) 
                                    LEFT JOIN ed.extProtocol ON extProtocol.AbiturientId = extAbitAspirant.Id  ", MainClass.GetStringAbitNumber("qAbiturient"));

                string where = string.Format(" WHERE extProtocol.Id = '{0}' ", protocolId);
                string orderby = " ORDER BY Направление, Рег_Номер ";

                DataSet ds = MainClass.Bdc.GetDataSet(query + where + orderby);

                using (PriemEntities context = new PriemEntities())
                {
                    Guid ProtocolId = Guid.Parse(protocolId);

                    var info =
                        (from protocol in context.extProtocol
                         join sf in context.StudyForm
                         on protocol.StudyFormId equals sf.Id

                         where protocol.Id == ProtocolId 
                         && protocol.ProtocolTypeId == 2 && protocol.IsOld == false && protocol.Excluded == false//disEnable
                         select new
                         {
                             StudyFormName = sf.Name,
                             protocol.StudyBasisId,
                             protocol.Date,
                             protocol.Number
                         }).FirstOrDefault();

                    //string form = MainClass.Bdc.GetStringValue(string.Format("SELECT StudyForm.Acronym FROM StudyForm INNER JOIN Protocol ON Protocol.StudyFormId = StudyForm.Id WHERE Protocol.Id='{0}'", protocolId));
                    //string basisId = MainClass.Bdc.GetStringValue(string.Format("SELECT StudyBasis.Id FROM StudyBasis INNER JOIN Protocol ON Protocol.StudyBasisId = StudyBasis.Id WHERE Protocol.Id='{0}'", protocolId));
                    //DateTime protocolDate = (DateTime)MainClass.Bdc.GetValue(string.Format("SELECT Protocol.Date FROM Protocol WHERE Protocol.Id='{0}'", protocolId));
                    //string protocolNum = MainClass.Bdc.GetStringValue(string.Format("SELECT Protocol.Number FROM Protocol WHERE Protocol.Id='{0}'", protocolId));

                    string form = info.StudyFormName;
                    string basisId = info.StudyBasisId.ToString();
                    DateTime protocolDate = info.Date.HasValue ? info.Date.Value : DateTime.Now;
                    string protocolNum = info.Number;


                    string basis = string.Empty;

                    switch (basisId)
                    {
                        case "1":
                            basis = "Бюджетные места";
                            break;
                        case "2":
                            basis = "Места по договорам с оплатой стоимости обучения";
                            break;
                    }

                    Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);

                    using (fileS = new FileStream(savePath, FileMode.Create))
                    {

                        BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font font = new Font(bfTimes, 10);

                        PdfWriter.GetInstance(document, fileS);
                        document.Open();

                        //HEADER
                        string header = string.Format(@"Форма обучения: {0}
Условия обучения: {1}", form, basis);

                        Paragraph p = new Paragraph(header, font);
                        document.Add(p);

                        float midStr = 13f;
                        p = new Paragraph(20f);
                        p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                        p.Add(new Phrase(protocolNum, new Font(bfTimes, 18, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph(midStr);
                        p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
об исключении из участия в конкурсе на основные образовательные программы ", new Font(bfTimes, 10, Font.BOLD)));

                        /*
                        p.Add(new Phrase(string.Format("{0} {1} {2}", "KODOKSO", "PROFESSION", "(SPECIALIZATION)"),
                            new Font(bfTimes, 10, Font.UNDERLINE + Font.BOLD)));*/
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        //date
                        p = new Paragraph(midStr);
                        p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(protocolDate, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);


                        string spec = "";
                        PdfPTable curT = null;
                        int cnt = 0;
                        string currSpec = null;
                        string napravlenie = null;
                        foreach (DataRow row in ds.Tables[0].Rows)
                        {
                            cnt++;

                            currSpec = row.Field<string>("Направление");
                            string code = row.Field<string>("Код");
                            napravlenie = "направлению";

                            if (spec != currSpec)
                            {
                                spec = currSpec;
                                cnt = 1;

                                if (curT != null)
                                {
                                    document.Add(curT);
                                }

                                //Table

                                Table table = new Table(7);
                                table.Padding = 3;
                                table.Spacing = 0;
                                float[] headerwidths = { 5, 10, 30, 15, 20, 10, 10 };
                                table.Widths = headerwidths;
                                table.Width = 100;

                                PdfPTable t = new PdfPTable(7);
                                t.SetWidthPercentage(headerwidths, document.PageSize);
                                t.WidthPercentage = 100f;
                                t.SpacingBefore = 10f;
                                t.SpacingAfter = 10f;

                                t.HeaderRows = 2;

                                Phrase pra = new Phrase(string.Format("По {0} {1} ", napravlenie, currSpec), new Font(bfTimes, 10));

                                PdfPCell pcell = new PdfPCell(pra);
                                pcell.BorderWidth = 0;
                                pcell.Colspan = 7;
                                t.AddCell(pcell);

                                string[] headers = new string[]
                        {
                            "№ п/п",
                            "Рег.номер",
                            "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                            "Номер аттестата или диплома",
                            "Номер сертификата ЕГЭ по профильному предмету",
                            "Вид конкурса",
                            "Примечания"
                        };
                                foreach (string h in headers)
                                {
                                    PdfPCell cell = new PdfPCell();
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                    t.AddCell(cell);
                                }

                                curT = t;
                            }

                            string quer = string.Format(@"
                                    SELECT TOP 1 EgeCertificate.Number FROM ed.EgeCertificate 
                                    INNER JOIN ed.EgeMark ON EgeMark.EgeCertificateId= EgeCertificateId
                                    INNER JOIN ed.EgeToExam ON EgeToExam.EgeExamNameId = EgeMark.EgeExamNameId
                                    WHERE EgeCertificate.PersonId='{0}' AND EgeToExam.ExamId = 
                                    (SELECT TOP 1 ExamId FROM ed.ExamInEntry WHERE ExamInEntry.EntryId='{1}' AND IsProfil>0)",
                                         row["PersonId"].ToString(),
                                         row["EntryId"].ToString());

                            string egecert = MainClass.Bdc.GetStringValue(quer);

                            curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Рег_Номер"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("ФИО"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Аттестат"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(egecert, new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Конкурс"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Примечания"), new Font(bfTimes, 10)));
                        }

                        if (curT != null)
                        {
                            document.Add(curT);
                        }

                        //FOOTER
                        p = new Paragraph(30f);
                        p.KeepTogether = true;
                        p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ_______________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase(@"Заместитель Ответственного секретаря Приемной 
комиссии  СПбГУ по группе основных образовательных программ_____________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Ответственный по приему на основную образовательную программу___________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        document.Close();



                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintChangeCompCelProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                string query =
                    string.Format(@"SELECT DISTINCT extAbitAspirant.Id as Id,
                                    extAbitAspirant.RegNum as Рег_Номер, extPersonAspirant.Surname + ' '+extPersonAspirant.[Name] + ' ' + extPersonAspirant.SecondName as ФИО, 
                                    (case when extPersonAspirant.SchoolTypeId = 1 then extPersonAspirant.AttestatRegion + ' ' + extPersonAspirant.AttestatSeries + '  №' + extPersonAspirant.AttestatNum else extPersonAspirant.DiplomSeries + '  №' + extPersonAspirant.DiplomNum end) as Аттестат, 
                                    qEntry.LicenseProgramCode + ' ' + qEntry.LicenseProgramName + ', ' + qEntry.ObrazProgramName + ', ' + ( Case when qEntry.ProfileId IS NOT NULL then qEntry.ProfileName else '' end) as Направление,
                                    qEntry.LicenseProgramCode as Код, Competition.NAme as Конкурс, 
                                    extAbitAspirant.PersonId, extAbitAspirant.EntryId,
                                    (CASE WHEN extAbitAspirant.BackDoc > 0 THEN 'Забрал док.' ELSE (CASE WHEN extAbitAspirant.NotEnabled > 0 THEN 'Не допущен'ELSE '' END) END) as Примечания 
                                    FROM ((ed.extAbitAspirant 
                                    INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=extAbitAspirant.PersonId 
                                    INNER JOIN ed.qEntry ON qEntry.Id = extAbitAspirant.EntryId)
                                    LEFT JOIN ed.Competition ON Competition.Id = extAbitAspirant.CompetitionId) 
                                    LEFT JOIN ed.extProtocol ON extProtocol.AbiturientId = extAbitAspirant.Id ", MainClass.GetStringAbitNumber("qAbiturient"));

                string where = string.Format(" WHERE extProtocol.Id = '{0}' ", protocolId);
                string orderby = " ORDER BY Направление, Рег_Номер ";

                DataSet ds = MainClass.Bdc.GetDataSet(query + where + orderby);

                using (PriemEntities context = new PriemEntities())
                {
                    Guid ProtocolId = Guid.Parse(protocolId);

                    var info =
                        (from protocol in context.extProtocol
                         join sf in context.StudyForm
                         on protocol.StudyFormId equals sf.Id

                         where protocol.Id == ProtocolId
                         && protocol.ProtocolTypeId == 3 && protocol.IsOld == false && protocol.Excluded == false//ChangeCompCel
                         select new
                         {
                             StudyFormName = sf.Name,
                             protocol.StudyBasisId,
                             protocol.Date,
                             protocol.Number
                         }).FirstOrDefault();

                    string form = info.StudyFormName;
                    string basisId = info.StudyBasisId.ToString();
                    DateTime protocolDate = info.Date.HasValue ? info.Date.Value : DateTime.Now;
                    string protocolNum = info.Number;

                    //string form = MainClass.Bdc.GetStringValue(string.Format("SELECT StudyForm.Acronym FROM StudyForm INNER JOIN Protocol ON Protocol.StudyFormId = StudyForm.Id WHERE Protocol.Id='{0}'", protocolId));
                    //string basisId = MainClass.Bdc.GetStringValue(string.Format("SELECT StudyBasis.Id FROM StudyBasis INNER JOIN Protocol ON Protocol.StudyBasisId = StudyBasis.Id WHERE Protocol.Id='{0}'", protocolId));
                    //DateTime protocolDate = (DateTime)MainClass.Bdc.GetValue(string.Format("SELECT Protocol.Date FROM Protocol WHERE Protocol.Id='{0}'", protocolId));
                    //string protocolNum = MainClass.Bdc.GetStringValue(string.Format("SELECT Protocol.Number FROM Protocol WHERE Protocol.Id='{0}'", protocolId));
                    
                    string basis = string.Empty;

                    switch (basisId)
                    {
                        case "1":
                            basis = "Бюджетные места";
                            break;
                        case "2":
                            basis = "Места по договорам с оплатой стоимости обучения";
                            break;
                    }

                    Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);

                    using (fileS = new FileStream(savePath, FileMode.Create))
                    {

                        BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font font = new Font(bfTimes, 10);

                        PdfWriter.GetInstance(document, fileS);
                        document.Open();

                        //HEADER
                        string header = string.Format(@"Форма обучения: {0}
Условия обучения: {1}", form, basis);

                        Paragraph p = new Paragraph(header, font);
                        document.Add(p);

                        float midStr = 13f;
                        p = new Paragraph(20f);
                        p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                        p.Add(new Phrase(protocolNum, new Font(bfTimes, 18, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph(midStr);
                        p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
об изменении типа конкурса целевикам ", new Font(bfTimes, 10, Font.BOLD)));

                        /*
                        p.Add(new Phrase(string.Format("{0} {1} {2}", "KODOKSO", "PROFESSION", "(SPECIALIZATION)"),
                            new Font(bfTimes, 10, Font.UNDERLINE + Font.BOLD)));*/
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        //date
                        p = new Paragraph(midStr);
                        p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(protocolDate, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);


                        string spec = "";
                        PdfPTable curT = null;
                        int cnt = 0;
                        string currSpec = null;
                        string napravlenie = null;
                        foreach (DataRow row in ds.Tables[0].Rows)
                        {
                            cnt++;

                            currSpec = row.Field<string>("Направление");
                            string code = row.Field<string>("Код");
                            napravlenie = "направлению";

                            if (spec != currSpec)
                            {
                                spec = currSpec;
                                cnt = 1;

                                if (curT != null)
                                {
                                    document.Add(curT);
                                }

                                //Table

                                Table table = new Table(6);
                                table.Padding = 3;
                                table.Spacing = 0;
                                float[] headerwidths = { 5, 10, 30, 15, 10, 10 };
                                table.Widths = headerwidths;
                                table.Width = 100;

                                PdfPTable t = new PdfPTable(6);
                                t.SetWidthPercentage(headerwidths, document.PageSize);
                                t.WidthPercentage = 100f;
                                t.SpacingBefore = 10f;
                                t.SpacingAfter = 10f;

                                t.HeaderRows = 2;

                                Phrase pra = new Phrase(string.Format("По {0} {1} ", napravlenie, currSpec), new Font(bfTimes, 10));

                                PdfPCell pcell = new PdfPCell(pra);
                                pcell.BorderWidth = 0;
                                pcell.Colspan = 7;
                                t.AddCell(pcell);

                                string[] headers = new string[]
                        {
                            "№ п/п",
                            "Рег.номер",
                            "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                            "Номер аттестата или диплома",                            
                            "Новый вид конкурса",
                            "Примечания"
                        };
                                foreach (string h in headers)
                                {
                                    PdfPCell cell = new PdfPCell();
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                    t.AddCell(cell);
                                }

                                curT = t;
                            }

                            curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Рег_Номер"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("ФИО"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Аттестат"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Конкурс"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Примечания"), new Font(bfTimes, 10)));
                        }

                        if (curT != null)
                        {
                            document.Add(curT);
                        }

                        //FOOTER
                        p = new Paragraph(30f);
                        p.KeepTogether = true;
                        p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ____________________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Заместитель начальника Управления по организации приема – советник проректора по направлениям___________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Ответственный секретарь комиссии по приему документов_______________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        document.Close();


                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintChangeCompBEProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                string query =
                    string.Format(@"SELECT DISTINCT extAbitAspirant.Id as Id,
                                    extAbitAspirant.RegNum as Рег_Номер, extPersonAspirant.Surname + ' '+extPersonAspirant.[Name] + ' ' + extPersonAspirant.SecondName as ФИО, 
                                    (case when extPersonAspirant.SchoolTypeId = 1 then extPersonAspirant.AttestatRegion + ' ' + extPersonAspirant.AttestatSeries + '  №' + extPersonAspirant.AttestatNum else extPersonAspirant.DiplomSeries + '  №' + extPersonAspirant.DiplomNum end) as Аттестат, 
                                    qEntry.LicenseProgramCode + ' ' + qEntry.LicenseProgramName + ', ' + qEntry.ObrazProgramName + ', ' + ( Case when qEntry.ProfileId IS NOT NULL then qEntry.ProfileName else '' end) as Направление,
                                    qEntry.LicenseProgramCode as Код, Competition.NAme as Конкурс, 
                                    extAbitAspirant.PersonId, extAbitAspirant.EntryId,
                                    (CASE WHEN extAbitAspirant.BackDoc > 0 THEN 'Забрал док.' ELSE (CASE WHEN extAbitAspirant.NotEnabled > 0 THEN 'Не допущен'ELSE '' END) END) as Примечания 
                                    FROM ((ed.extAbitAspirant 
                                    INNER JOIN ed.extPersonAspirant ON extPersonAspirant.Id=extAbitAspirant.PersonId 
                                    INNER JOIN ed.qEntry ON qEntry.Id = extAbitAspirant.EntryId)
                                    LEFT JOIN ed.Competition ON Competition.Id = extAbitAspirant.CompetitionId) 
                                    LEFT JOIN ed.extProtocol ON extProtocol.AbiturientId = extAbitAspirant.Id ", MainClass.GetStringAbitNumber("qAbiturient"));

                string where = string.Format(" WHERE extProtocol.Id = '{0}' ", protocolId);
                string orderby = " ORDER BY Направление, Рег_Номер ";

                DataSet ds = MainClass.Bdc.GetDataSet(query + where + orderby);

                using (PriemEntities context = new PriemEntities())
                {
                    Guid ProtocolId = Guid.Parse(protocolId);

                    var info =
                        (from protocol in context.extProtocol
                         join sf in context.StudyForm
                         on protocol.StudyFormId equals sf.Id
                         
                         where protocol.Id == ProtocolId
                         && protocol.ProtocolTypeId == 6 && protocol.IsOld == false && protocol.Excluded == false//ChangeCompBE
                         select new
                         {
                             StudyFormName = sf.Name,
                             protocol.StudyBasisId,
                             protocol.Date,
                             protocol.Number
                         }).FirstOrDefault();

                    string form = info.StudyFormName;
                    string basisId = info.StudyBasisId.ToString();
                    DateTime protocolDate = info.Date.HasValue ? info.Date.Value : DateTime.Now;
                    string protocolNum = info.Number;

                    //string form = MainClass.Bdc.GetStringValue(string.Format("SELECT StudyForm.Acronym FROM StudyForm INNER JOIN Protocol ON Protocol.StudyFormId = StudyForm.Id WHERE Protocol.Id='{0}'", protocolId));
                    //string basisId = MainClass.Bdc.GetStringValue(string.Format("SELECT StudyBasis.Id FROM StudyBasis INNER JOIN Protocol ON Protocol.StudyBasisId = StudyBasis.Id WHERE Protocol.Id='{0}'", protocolId));
                    //DateTime protocolDate = (DateTime)MainClass.Bdc.GetValue(string.Format("SELECT Protocol.Date FROM Protocol WHERE Protocol.Id='{0}'", protocolId));
                    //string protocolNum = MainClass.Bdc.GetStringValue(string.Format("SELECT Protocol.Number FROM Protocol WHERE Protocol.Id='{0}'", protocolId));

                    string basis = string.Empty;

                    switch (basisId)
                    {
                        case "1":
                            basis = "Бюджетные места";
                            break;
                        case "2":
                            basis = "Места по договорам с оплатой стоимости обучения";
                            break;
                    }

                    Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);

                    using (fileS = new FileStream(savePath, FileMode.Create))
                    {

                        BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font font = new Font(bfTimes, 10);

                        PdfWriter.GetInstance(document, fileS);
                        document.Open();

                        //HEADER
                        string header = string.Format(@"Форма обучения: {0}
Условия обучения: {1}", form, basis);

                        Paragraph p = new Paragraph(header, font);
                        document.Add(p);

                        float midStr = 13f;
                        p = new Paragraph(20f);
                        p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                        p.Add(new Phrase(protocolNum, new Font(bfTimes, 18, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph(midStr);
                        p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
об изменении типа конкурса на общий ", new Font(bfTimes, 10, Font.BOLD)));

                        /*
                        p.Add(new Phrase(string.Format("{0} {1} {2}", "KODOKSO", "PROFESSION", "(SPECIALIZATION)"),
                            new Font(bfTimes, 10, Font.UNDERLINE + Font.BOLD)));*/
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        //date
                        p = new Paragraph(midStr);
                        p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(protocolDate, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);


                        string spec = "";
                        PdfPTable curT = null;
                        int cnt = 0;
                        string currSpec = null;
                        string napravlenie = null;
                        foreach (DataRow row in ds.Tables[0].Rows)
                        {
                            cnt++;

                            currSpec = row.Field<string>("Направление");
                            string code = row.Field<string>("Код");
                            napravlenie = "направлению";

                            if (spec != currSpec)
                            {
                                spec = currSpec;
                                cnt = 1;

                                if (curT != null)
                                {
                                    document.Add(curT);
                                }

                                //Table

                                Table table = new Table(6);
                                table.Padding = 3;
                                table.Spacing = 0;
                                float[] headerwidths = { 5, 10, 30, 15, 10, 10 };
                                table.Widths = headerwidths;
                                table.Width = 100;

                                PdfPTable t = new PdfPTable(6);
                                t.SetWidthPercentage(headerwidths, document.PageSize);
                                t.WidthPercentage = 100f;
                                t.SpacingBefore = 10f;
                                t.SpacingAfter = 10f;

                                t.HeaderRows = 2;

                                Phrase pra = new Phrase(string.Format("По {0} {1} ", napravlenie, currSpec), new Font(bfTimes, 10));

                                PdfPCell pcell = new PdfPCell(pra);
                                pcell.BorderWidth = 0;
                                pcell.Colspan = 7;
                                t.AddCell(pcell);

                                string[] headers = new string[]
                        {
                            "№ п/п",
                            "Рег.номер",
                            "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                            "Номер аттестата или диплома",                            
                            "Новый вид конкурса",
                            "Примечания"
                        };
                                foreach (string h in headers)
                                {
                                    PdfPCell cell = new PdfPCell();
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                    t.AddCell(cell);
                                }

                                curT = t;
                            }

                            curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Рег_Номер"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("ФИО"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Аттестат"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Конкурс"), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(row.Field<string>("Примечания"), new Font(bfTimes, 10)));
                        }

                        if (curT != null)
                        {
                            document.Add(curT);
                        }

                        //FOOTER
                        p = new Paragraph(30f);
                        p.KeepTogether = true;
                        p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ____________________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Заместитель начальника Управления по организации приема – советник проректора по направлениям___________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Ответственный секретарь комиссии по приему документов_______________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        document.Close();


                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        //public static void PrintDogovor(Guid dogId, Guid abitId, bool forPrint)
        //{
        //    using (PriemEntities context = new PriemEntities())
        //    {
        //        var abit = context.extAbitAspirant.Where(x => x.Id == abitId).FirstOrDefault();
        //        if (abit == null)
        //        {
        //            WinFormsServ.Error("Не удалось загрузить данные заявления");
        //            return;
        //        }

        //        var person = context.extPersonAll.Where(x => x.Id == abit.PersonId).FirstOrDefault();
        //        if (person == null)
        //        {
        //            WinFormsServ.Error("Не удалось загрузить данные абитуриента");
        //            return;
        //        }

        //        var dogovorInfo =
        //            (from pd in context.PaidData
        //             join pi in context.PayDataEntry on pd.Abiturient.EntryId equals pi.EntryId into pi2
        //             from pi in pi2.DefaultIfEmpty()
        //             where pd.Id == dogId
        //             select new
        //             {
        //                 pd.DogovorNum,
        //                 DogovorTypeName = pd.DogovorType.Name,
        //                 pd.DogovorDate,
        //                 pd.Qualification,
        //                 pd.Srok,
        //                 pd.DateStart,
        //                 pd.DateFinish,
        //                 pd.SumFirstYear,
        //                 pd.SumFirstPeriod,
        //                 pd.Parent,
        //                 Prorector = pd.Prorektor.NameFull,
        //                 PayPeriodName = pd.PayPeriod.Name,
        //                 pd.AbitFIORod,
        //                 pd.AbiturientId,
        //                 pd.Customer,
        //                 pd.CustomerLico,
        //                 pd.CustomerReason,
        //                 pd.CustomerAddress,
        //                 pd.CustomerPassport,
        //                 pd.CustomerPassportAuthor,
        //                 pd.CustomerINN,
        //                 pd.CustomerRS,
        //                 pd.Prorektor.DateDov,
        //                 pd.Prorektor.NumberDov,
        //                 PayPeriod = pd.PayPeriod.Name,
        //                 PayPeriodPad = pd.PayPeriod.NamePad,
        //                 DogovorTypeId = pd.DogovorTypeId,
        //                 pi.UniverName,
        //                 pi.UniverAddress,
        //                 pi.UniverINN,
        //                 pi.UniverRS,
        //                 pi.Props
        //             }).FirstOrDefault();

        //        string dogType = dogovorInfo.DogovorTypeId.ToString();

        //        WordDoc wd = new WordDoc(string.Format(@"{0}\Dogovor{1}.dot", MainClass.dirTemplates, dogType), !forPrint);

        //        //вступление
        //        wd.SetFields("DogovorNum", dogovorInfo.DogovorNum.ToString());
        //        wd.SetFields("DogovorDate", dogovorInfo.DogovorDate.Value.ToLongDateString());

        //        //проректор и студент
        //        wd.SetFields("Lico", dogovorInfo.Prorector);
        //        wd.SetFields("LicoDateNum", dogovorInfo.DateDov.ToString() + "г. " + dogovorInfo.NumberDov.ToString());
        //        wd.SetFields("FIO", person.FIO);


        //        string programcode = (abit.ObrazProgramCrypt ?? "").Trim();
        //        string profcode = abit.LicenseProgramCode.Trim();

        //        //wd.SetFields("ObrazProgramName", abit.ObrazProgramName.Trim());
        //        //wd.SetFields("ObrazProgramName1", abit.ObrazProgramName.Trim());

        //        //wd.SetFields("ProgramCode", programcode);
        //        wd.SetFields("Profession", abit.LicenseProgramName);

        //        //wd.SetFields("StudyCourse", "1");
        //        //wd.SetFields("StudyFaculty", abit.FacultyName);
        //        string form = context.StudyForm.Where(x => x.Id == abit.StudyFormId).Select(x => x.Name).FirstOrDefault().ToLower();
        //        wd.SetFields("StudyForm", form.ToLower());

        //        //wd.SetFields("Qualification", dogovorInfo.Qualification);

        //        //сроки обучения
        //        wd.SetFields("Srok", dogovorInfo.Srok);
        //        wd.SetFields("Srok1", dogovorInfo.Srok); 
        //        DateTime dStart = dogovorInfo.DateStart.Value; 
        //        wd.SetFields("DateStart", dStart.ToLongDateString());
        //        DateTime dFinish = dogovorInfo.DateFinish.Value; 
        //        wd.SetFields("DateFinish", dFinish.ToLongDateString());

        //        //суммы обучения
        //        wd.SetFields("SumFirstYear", dogovorInfo.SumFirstYear);
        //        wd.SetFields("SumFirstPeriod", dogovorInfo.SumFirstPeriod);

        //        wd.SetFields("PayPeriod", dogovorInfo.PayPeriod);

        //        //wd.SetFields("Parent", dogovorInfo.Parent);

        //        //if (dogovorInfo.Parent.Trim().Length > 0)
        //        //    wd.SetFields("AbitFIORod", dogovorInfo.AbitFIORod);

        //        wd.SetFields("Address1", string.Format("{0} {1} {2}, {3}, ", person.Code, person.CountryName, person.RegionName, person.City));
        //        wd.SetFields("Address2", string.Format("{0} дом {1} {2} кв. {3}", person.Street, person.House, person.Korpus == string.Empty ? "" : "корп. " + person.Korpus, person.Flat));

        //        wd.SetFields("Passport", "серия " + person.PassportSeries + "№ " + person.PassportNumber);
        //        wd.SetFields("PassportAuthor", "выдан " + person.PassportDate.Value.ToShortDateString() + " " + person.PassportAuthor);

        //        wd.SetFields("UniverName", dogovorInfo.UniverName);
        //        wd.SetFields("UniverAddress", dogovorInfo.UniverAddress);
        //        wd.SetFields("UniverINN", dogovorInfo.UniverINN);
        //        wd.SetFields("UniverRS", dogovorInfo.UniverRS);
        //        wd.SetFields("Props", dogovorInfo.Props);

        //        switch (dogType)
        //        {
        //            case "1":
        //                {
        //                    break;
        //                }
        //            case "2":
        //                {
        //                    wd.SetFields("Customer", dogovorInfo.Customer);
        //                    wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);
        //                    wd.SetFields("CustomerINN", dogovorInfo.CustomerPassport);
        //                    wd.SetFields("CustomerRS", dogovorInfo.CustomerPassportAuthor);

        //                    break;
        //                }
        //            case "4":
        //                {
        //                    wd.SetFields("Customer", dogovorInfo.Customer);
        //                    wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);
        //                    wd.SetFields("CustomerINN", dogovorInfo.CustomerPassport);
        //                    wd.SetFields("CustomerRS", dogovorInfo.CustomerPassportAuthor);

        //                    break;
        //                }
        //            case "3":
        //                {
        //                    wd.SetFields("Customer", dogovorInfo.Customer);
        //                    wd.SetFields("CustomerLico", dogovorInfo.CustomerLico);
        //                    wd.SetFields("CustomerReason", dogovorInfo.CustomerReason);
        //                    wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);
        //                    wd.SetFields("CustomerINN", dogovorInfo.CustomerINN);
        //                    wd.SetFields("CustomerRS", dogovorInfo.CustomerRS);

        //                    break;
        //                }
        //        }

        //        if (forPrint)
        //        {
        //            wd.Print();
        //            wd.Close();
        //        }

        //    }
        //}
        public static void PrintDogovor(Guid dogId, Guid abitId, bool forPrint)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var abit = context.extAbitAspirant.Where(x => x.Id == abitId).FirstOrDefault();
                if (abit == null)
                {
                    WinFormsServ.Error("Не удалось загрузить данные заявления");
                    return;
                }

                var person = context.extPersonAll.Where(x => x.Id == abit.PersonId).FirstOrDefault();
                if (person == null)
                {
                    WinFormsServ.Error("Не удалось загрузить данные абитуриента");
                    return;
                }

                var dogovorInfo =
                    (from pd in context.PaidData
                     join pi in context.PayDataEntry on pd.Abiturient.EntryId equals pi.EntryId into pi2
                     from pi in pi2.DefaultIfEmpty()
                     where pd.Id == dogId
                     select new
                     {
                         pd.DogovorNum,
                         DogovorTypeName = pd.DogovorType.Name,
                         pd.DogovorDate,
                         pd.Qualification,
                         pd.Srok,
                         pd.SrokIndividual,
                         pd.DateStart,
                         pd.DateFinish,
                         pd.SumTotal,
                         pd.SumFirstYear,
                         pd.SumFirstPeriod,
                         pd.Parent,
                         Prorector = pd.Prorektor.NameFull,
                         PayPeriodName = pd.PayPeriod.Name,
                         pd.AbitFIORod,
                         pd.AbiturientId,
                         pd.Customer,
                         pd.CustomerLico,
                         pd.CustomerReason,
                         pd.CustomerAddress,
                         pd.CustomerPassport,
                         pd.CustomerPassportAuthor,
                         pd.CustomerINN,
                         pd.CustomerRS,
                         pd.Prorektor.DateDov,
                         pd.Prorektor.NumberDov,
                         PayPeriod = pd.PayPeriod.Name,
                         PayPeriodPad = pd.PayPeriod.NamePad,
                         DogovorTypeId = pd.DogovorTypeId,
                         pi.UniverName,
                         pi.UniverAddress,
                         pi.UniverINN,
                         pi.UniverRS,
                         pi.Props
                     }).FirstOrDefault();

                //bool IsCommonWithParent = false;
                string dogType = dogovorInfo.DogovorTypeId.ToString();
                //if (dogType.Equals("1"))
                //    if (!String.IsNullOrEmpty(dogovorInfo.Parent))
                //    {
                //        dogType = "2";
                //       IsCommonWithParent = true;
                //    }

                WordDoc wd = new WordDoc(string.Format(@"{0}\Dogovor{1}.dot", MainClass.dirTemplates, dogType), !forPrint);

                //вступление
                wd.SetFields("DogovorNum", dogovorInfo.DogovorNum.ToString());
                wd.SetFields("DogovorDate", dogovorInfo.DogovorDate.Value.ToLongDateString());

                //wd.SetFields("DogovorDay", ((DateTime)dsRow["DogovorDate"]).Date.Day.ToString());
                //wd.SetFields("DogovorMonth", ((DateTime)dsRow["DogovorDate"]).Date.Month.ToString());
                //wd.SetFields("DogovorYear", ((DateTime)dsRow["DogovorDate"]).Date.Year.ToString());

                //проректор и студент
                wd.SetFields("Lico", dogovorInfo.Prorector);
                wd.SetFields("LicoDate", dogovorInfo.DateDov.ToString() + "г.");
                wd.SetFields("LicoNum", dogovorInfo.NumberDov.ToString());
                wd.SetFields("FIO", person.FIO);
                wd.SetFields("Sex", (person.Sex) ? "ый" : "ая");



                //                DataSet dsProgram = _bdc.GetDataSet(string.Format(@"SELECT hlpStudyPlan.ProgramCode, hlpStudyPlan.ProfessionCode, hlpStudyPlan.ObrazProgram, hlpStudyPlan.Profession, hlpStudyPlan.Specialization, 
                //            (Case When hlpStudyPlan.FacultyId IN (4,6,7,12,13,14,15,16) then 'факультет ' + hlpStudyPlan.Faculty 
                //            else (case when hlpStudyPlan.FacultyId IN (1,2,5,8,9,10,17,18,19,20,21,22) then hlpStudyPlan.Faculty + ' факультет' 
                //                 else hlpStudyPlan.Faculty end) end) AS FacultyName,
                //            hlpStudyPlan.ProfessionCode, StudyForm.Acronym AS StudyForm FROM qAbiturient 
                //            INNER JOIN hlpStudyPlan ON hlpStudyPlan.StudyPlanId = qAbiturient.StudyPlanId 
                //            INNER JOIN StudyForm ON hlpStudyPlan.StudyFormId = StudyForm.Id WHERE qAbiturient.Id = '{0}' ", abitId));
                //                DataRow drPr = dsProgram.Tables[0].Rows[0];

                string programcode = abit.ObrazProgramCrypt.Trim();// drPr["ProgramCode"].ToString().Trim();
                string profcode = abit.LicenseProgramCode.Trim();// drPr["ProfessionCode"].ToString().Trim();
                string level = "";

                wd.SetFields("ObrazProgramName", "(" + programcode + ") " + abit.ObrazProgramName.Trim());//drPr["ObrazProgram"].ToString().Trim()
                // wd.SetFields("ObrazProgramName1", abit.ObrazProgramName.Trim());//drPr["ObrazProgram"].ToString().Trim()

                // wd.SetFields("ProgramCode", programcode);

                wd.SetFields("Profession", "(" + profcode + ") " + abit.LicenseProgramName);//drPr["Profession"].ToString().Trim()

                wd.SetFields("StudyCourse", "1");
                wd.SetFields("StudyFaculty", abit.FacultyName);
                string form = context.StudyForm.Where(x => x.Id == abit.StudyFormId).Select(x => x.Name).FirstOrDefault().ToLower();
                //_bdc.GetStringValue("SELECT Acronym FROM StudyForm WHERE Id = " + abit.StudyForm);
                wd.SetFields("StudyForm", form.ToLower());
                //wd.SetFields("StudyLevel", level);

                //wd.SetFields("Program", programName + level + ", 1 курс " + drPr["FacultyName"].ToString() + ", " + prof + " " + profcode + " " + programName + ", " + drPr["StudyForm"].ToString().ToLower() + " форма обучения");

                wd.SetFields("Qualification", dogovorInfo.Qualification);//dsRow["Qualification"].ToString()

                //сроки обучения
                wd.SetFields("Srok", dogovorInfo.Srok); //dsRow["Srok"].ToString()
                //wd.SetFields("SrokIndividual", dogovorInfo.SrokIndividual); //dsRow["Srok"].ToString()

                DateTime dStart = dogovorInfo.DateStart.Value; //(DateTime)dsRow["DateStart"];
                //wd.SetFields("DateStart", "\"" + dStart.Date.Day.ToString() + "\" " + dStart.Date.Month.ToString() + " " + dStart.Date.Year.ToString());
                wd.SetFields("DateStart", dStart.ToLongDateString());
                DateTime dFinish = dogovorInfo.DateFinish.Value; //(DateTime)dsRow["DateFinish"];
                //wd.SetFields("DateFinish", "\"" + dFinish.Date.Day.ToString() + "\" " + dFinish.Date.Month.ToString() + " " + dFinish.Date.Year.ToString());
                wd.SetFields("DateFinish", dFinish.ToLongDateString());

                //суммы обучения
                wd.SetFields("SumTotal", dogovorInfo.SumTotal);//dsRow["SumFirstYear"].ToString()

                //wd.SetFields("SumFirstYear", dogovorInfo.SumFirstYear);//dsRow["SumFirstYear"].ToString()
                wd.SetFields("SumFirstPeriod", dogovorInfo.SumFirstPeriod);//dsRow["SumFirstPeriod"].ToString()

                //wd.SetFields("PayPeriod", dogovorInfo.PayPeriod);//dsRow["PayPeriod"].ToString()


                //wd.SetFields("Parent", dogovorInfo.Parent);//dsRow["Parent"].ToString()

                /*if (dogovorInfo.Parent.Trim().Length > 0)
                    wd.SetFields("AbitFIORod", dogovorInfo.AbitFIORod);//dsRow["AbitFIORod"].ToString()
                */
                wd.SetFields("Address1", string.Format("{0} {1}, {2}, {3}, ", person.Code, person.CountryName, person.RegionName, person.City));
                wd.SetFields("Address2", string.Format("{0} дом {1} {2} кв. {3}", person.Street, person.House, person.Korpus == string.Empty ? "" : "корп. " + person.Korpus, person.Flat));

                wd.SetFields("Passport", "серия " + person.PassportSeries + " № " + person.PassportNumber);
                wd.SetFields("PassportAuthorDate", person.PassportDate.Value.ToShortDateString());
                wd.SetFields("PassportAuthor", person.PassportAuthor);

                wd.SetFields("PhoneNumber", person.Phone + (String.IsNullOrEmpty(person.Mobiles) ? "" : ", доп.: " + person.Mobiles));

                //string studyPlanId = _bdc.GetStringValue(string.Format("SELECT qAbiturient.StudyPlanId FROM qAbiturient WHERE Id = '{0}'", abitId));
                //DataRow dr = _bdc.GetDataSet("SELECT * FROM PayDataStudyPlan WHERE StudyPlanId = " + studyPlanId).Tables[0].Rows[0];

                wd.SetFields("UniverName", dogovorInfo.UniverName);//dr["UniverName"].ToString()
                wd.SetFields("UniverAddress", dogovorInfo.UniverAddress);//dr["UniverAddress"].ToString()
                wd.SetFields("UniverINN", dogovorInfo.UniverINN);//dr["UniverINN"].ToString()
                //wd.SetFields("UniverRS", dogovorInfo.UniverRS);//dr["UniverRS"].ToString()
                wd.SetFields("Props", dogovorInfo.Props);//dr["UniverDop"].ToString()

                switch (dogType)
                {
                    // обычный
                    case "1":
                        {
                            break;
                        }
                    // физ лицо
                    case "2":
                        {
                            wd.SetFields("CustomerLico", dogovorInfo.Customer);//dsRow["Customer"].ToString()
                            //wd.SetFields("AbitFIORod2", dsRow["AbitFIORod"].ToString());
                            wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);//dsRow["CustomerAddress"].ToString()
                            wd.SetFields("CustomerINN", "Паспорт: " + dogovorInfo.CustomerPassport);//dsRow["CustomerPassport"].ToString()
                            wd.SetFields("CustomerRS", "Выдан: " + dogovorInfo.CustomerPassportAuthor);//dsRow["CustomerPassportAuthor"].ToString()

                            break;
                        }
                    // мат кап
                    case "4":
                        {
                            wd.SetFields("Customer", dogovorInfo.Customer);//dsRow["Customer"].ToString()
                            //wd.SetFields("AbitFIORod2", dsRow["AbitFIORod"].ToString());
                            wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);//dsRow["CustomerAddress"].ToString()
                            wd.SetFields("CustomerINN", dogovorInfo.CustomerPassport);//dsRow["CustomerPassport"].ToString()
                            wd.SetFields("CustomerRS", dogovorInfo.CustomerPassportAuthor);//dsRow["CustomerPassportAuthor"].ToString()

                            break;
                        }
                    // юридическое лицо
                    case "3":
                        {
                            wd.SetFields("Customer", dogovorInfo.Customer);//dsRow["Customer"].ToString()
                            wd.SetFields("CustomerLico", dogovorInfo.CustomerLico);//dsRow["CustomerLico"].ToString()
                            wd.SetFields("CustomerReason", dogovorInfo.CustomerReason);//dsRow["CustomerReason"].ToString()
                            //wd.SetFields("AbitFIORod2", dsRow["AbitFIORod"].ToString());
                            wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);//dsRow["CustomerAddress"].ToString()
                            wd.SetFields("CustomerINN", "ИНН " + dogovorInfo.CustomerINN);//dsRow["CustomerINN"].ToString()
                            wd.SetFields("CustomerRS", "Р/С " + dogovorInfo.CustomerRS);//dsRow["CustomerRS"].ToString()

                            break;
                        }
                }

                if (forPrint)
                {
                    wd.Print();
                    wd.Close();
                }

            }
            //AbiturientClass abit = AbiturientClass.GetInstanceFromDBForPrint(abitId);
            //PersonClass person = PersonClass.GetInstanceFromDBForPrint(abit.PersonId);

            //DataSet ds = _bdc.GetDataSet(string.Format("SELECT PaidData.DogovorNum, PaidData.DogovorTypeId, PaidData.DogovorDate, " +
            //    "PaidData.Qualification, PaidData.Srok, PaidData.DateStart, PaidData.DateFinish, " +
            //    "PaidData.SumFirstYear, PaidData.SumFirstPeriod, PaidData.Parent, " +
            //    "PaidData.ProrektorId, PaidData.PayPeriodId, PaidData.AbitFIORod, PaidData.AbiturientId, " +
            //    "PaidData.Customer, PaidData.CustomerLico, PaidData.CustomerAddress, PaidData.CustomerPassport, " +
            //    "PaidData.CustomerPassportAuthor, PaidData.CustomerReason, PaidData.CustomerINN, PaidData.CustomerRS, " +
            //    "Prorektor.NameFull AS Prorektor, Prorektor.DateDov, Prorektor.NumberDov, PayPeriod.Name AS PayPeriod, PayPeriod.NamePad AS PayPeriodPad " +
            //    "FROM PaidData LEFT JOIN Prorektor ON PaidData.ProrektorId = Prorektor.Id " +
            //    "LEFT JOIN PayPeriod ON PaidData.PayPeriodId = PayPeriod.Id " +
            //    "WHERE PaidData.Id = '{0}'", dogId));

            //DataRow dsRow = ds.Tables[0].Rows[0];
        }

        public static void PrintDocInventory(IList<int> ids, Guid? _abitId)
        {
            string strIds = Util.BuildStringWithCollection(ids);
            using (PriemEntities context = new PriemEntities())
            {
                var abit = context.extAbitAspirant.Where(x => x.Id == _abitId).FirstOrDefault();
                if (abit == null)
                {
                    WinFormsServ.Error("Не найдены данные по заявлению!");
                    return;
                }
                Guid PersonId = abit.PersonId ?? Guid.Empty;
                var person = context.Person.Where(x => x.Id == PersonId).FirstOrDefault();
                if (person == null)
                {
                    WinFormsServ.Error("Не найдены данные по человеку!");
                    return;
                }
                string FIO = (person.Surname ?? "") + " " + (person.Name ?? "") + " " + (person.SecondName ?? "");
                WordDoc wd = new WordDoc(string.Format(@"{0}\DocInventory.dot", MainClass.dirTemplates), true);

                wd.SetFields("FIO", FIO);

                var docs = context.AbitDoc.Join(ids, x => x.Id, y => y, (x, y) => new { x.Id, x.Name }).Select(x => x.Name);

                int i = 1;
                wd.AddNewTable(docs.Count(), 1);
                foreach (var d in docs)
                {
                    wd.Tables[0][0, i - 1] = i.ToString() + ") " + d + "\n";
                    i++;
                }
            }
        }

        public static void PrintRatingProtocol(int? iStudyFormId, int? iStudyBasisId, int? iFacultyId, int? iLicenseProgramId, int? iObrazProgramId, Guid? gProfileId, bool isCel, int plan, string savePath, bool isSecond, bool isReduced, bool isParallel)
        {
            FileStream fileS = null;
            try
            {
                Guid fixId;
                int? docNum;
                string form;
                string facDat;
                string prof;
                string obProg;
                string spec;

                using (PriemEntities ctx = new PriemEntities())
                {
                    fixId = (from fixierenView in ctx.FixierenView
                             where fixierenView.StudyFormId == iStudyFormId && fixierenView.StudyBasisId == iStudyBasisId && fixierenView.FacultyId == iFacultyId && fixierenView.LicenseProgramId == iLicenseProgramId &&
                             fixierenView.ObrazProgramId == iObrazProgramId && (gProfileId.HasValue ? fixierenView.ProfileId == gProfileId : true) && fixierenView.IsCel == isCel && fixierenView.IsSecond == isSecond && fixierenView.IsParallel == isParallel && fixierenView.IsReduced == isReduced
                             select fixierenView.Id).FirstOrDefault();

                    docNum = (from fixierenView in ctx.FixierenView
                              where fixierenView.Id == fixId
                              select fixierenView.DocNum).FirstOrDefault();


                    form = (from studyForm in ctx.StudyForm
                            where studyForm.Id == iStudyFormId
                            select studyForm.Acronym).FirstOrDefault();

                    facDat = (from sP_Faculty in ctx.SP_Faculty
                              where sP_Faculty.Id == iFacultyId
                              select sP_Faculty.DatName).FirstOrDefault();

                    prof = (from entry in ctx.Entry
                            where entry.LicenseProgramId == iLicenseProgramId
                            select entry.SP_LicenseProgram.Code + " " + entry.SP_LicenseProgram.Name).FirstOrDefault();

                    obProg = (from entry in ctx.Entry
                              where entry.ObrazProgramId == iObrazProgramId
                              select ((entry.StudyLevel.Acronym + "." + entry.SP_ObrazProgram.Number + "." + MainClass.PriemYear) ?? "") + " " + entry.SP_ObrazProgram.Name).FirstOrDefault();

                    spec = (from entry in ctx.Entry
                            where gProfileId.HasValue ? entry.ProfileId == gProfileId : entry.ProfileId == null
                            select entry.ProfileName).FirstOrDefault();
                }

                string basis = string.Empty;

                switch (iStudyBasisId)
                {
                    case 1:
                        basis = "обучение за счет средств федерального бюджета";
                        break;
                    case 2:
                        basis = "обучение по договорам с оплатой стоимости обучения";
                        break;
                }

                Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);

                using (fileS = new FileStream(savePath, FileMode.Create))
                {

                    BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font font = new Font(bfTimes, 12);

                    PdfWriter writer = PdfWriter.GetInstance(document, fileS);
                    document.Open();

                    float firstLineIndent = 30f;
                    //HEADER
                    Paragraph p = new Paragraph("ПРАВИТЕЛЬСТВО РОССИЙСКОЙ ФЕДЕРАЦИИ", new Font(bfTimes, 12, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО", new Font(bfTimes, 10));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("ПРОФЕССИОНАЛЬНОГО ОБРАЗОВАНИЯ", new Font(bfTimes, 10));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("САНКТ-ПЕТЕРБУРГСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ", new Font(bfTimes, 12, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("(СПбГУ)", new Font(bfTimes, 12, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("ПРЕДСТАВЛЕНИЕ", new Font(bfTimes, 20, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph(10f);
                    p.Add(new Paragraph("По " + facDat, font));
                    p.Add(new Paragraph((form + " форма обучения").ToLower(), font));
                    p.Add(new Paragraph(basis, font));
                    p.IndentationLeft = 510;
                    document.Add(p);

                    p = new Paragraph("О зачислении на 1 курс", font);
                    p.SpacingBefore = 10f;
                    document.Add(p);

                    p = new Paragraph(@"В соответствии с Федеральным законом от 22.08.1996 N 125-Ф3 (ред. от 21.12.2009) «О высшем и послевузовском профессиональном образовании», Порядком приема граждан в имеющие государственную аккредитацию образовательные учреждения высшего профессионального образования, утвержденным Приказом Министерства образования и науки Российской Федерации от 21.10.2009 N 442 (ред. от 11.05.2010)", font);
                    p.SpacingBefore = 10f;
                    p.Alignment = Element.ALIGN_JUSTIFIED;
                    p.FirstLineIndent = firstLineIndent;
                    document.Add(p);

                    p = new Paragraph("Представляем на рассмотрение Приемной комиссии СПбГУ полный пофамильный перечень поступающих на 1 курс обучения по основным образовательным программам высшего профессионального образования:", font);
                    p.FirstLineIndent = firstLineIndent;
                    p.Alignment = Element.ALIGN_JUSTIFIED;
                    p.SpacingBefore = 20f;
                    document.Add(p);

                    p = new Paragraph("по направлению " + prof, font);
                    p.FirstLineIndent = firstLineIndent * 2;
                    document.Add(p);

                    p = new Paragraph("по образовательной программе " + obProg, font);
                    p.FirstLineIndent = firstLineIndent * 2;
                    document.Add(p);

                    if (!string.IsNullOrEmpty(spec))
                    {
                        p = new Paragraph("по профилю " + spec, font);
                        p.FirstLineIndent = firstLineIndent * 2;
                        document.Add(p);
                    }

                    //Table

                    float[] headerwidths = { 5, 9, 9, 19, 6, 10, 10, 7, 11, 14 };

                    PdfPTable t = new PdfPTable(10);
                    t.SetWidthPercentage(headerwidths, document.PageSize);
                    t.WidthPercentage = 100f;
                    t.SpacingBefore = 10f;
                    t.SpacingAfter = 10f;

                    t.HeaderRows = 1;

                    string[] headers = new string[]
                    {
                        "№ п/п",
                        "Рег. номер",
                        "Ид. номер",
                        "ФИО",
                        "Cумма баллов",
                        "Подлинники документов",
                        "Рейтинговый коэффициент",
                        "Конкурс",
                        "Профильное вступительное испытание",
                        "Дополнительное вступительное испытание"
                    };
                    foreach (string h in headers)
                    {
                        PdfPCell cell = new PdfPCell();
                        cell.BorderColor = Color.BLACK;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cell.AddElement(new Phrase(h, new Font(bfTimes, 12, Font.BOLD)));

                        t.AddCell(cell);
                    }

                    int counter = 0;

                    using (PriemEntities ctx = new PriemEntities())
                    {
                        var lst = (from extabit in ctx.extAbitAspirant
                                   join fixieren in ctx.Fixieren on extabit.Id equals fixieren.AbiturientId
                                   join fixierenView in ctx.FixierenView on fixieren.FixierenViewId equals fixierenView.Id into fixierenView2
                                   from fixierenView in fixierenView2.DefaultIfEmpty()
                                   join extperson in ctx.extPersonAll on extabit.PersonId equals extperson.Id
                                   join competition in ctx.Competition on extabit.CompetitionId equals competition.Id
                                   join hlpabiturientProfAdd in ctx.hlpAbiturientProfAdd on extabit.Id equals hlpabiturientProfAdd.Id into hlpabiturientProfAdd2
                                   from hlpabiturientProfAdd in hlpabiturientProfAdd2.DefaultIfEmpty()
                                   join hlpabiturientProf in ctx.hlpAbiturientProf on extabit.Id equals hlpabiturientProf.Id into hlpabiturientProf2
                                   from hlpabiturientProf in hlpabiturientProf2.DefaultIfEmpty()
                                   join extabitMarksSum in ctx.extAbitMarksSum on extabit.Id equals extabitMarksSum.Id into extabitMarksSum2
                                   from extabitMarksSum in extabitMarksSum2.DefaultIfEmpty()
                                   where fixierenView.StudyFormId == iStudyFormId && fixierenView.StudyBasisId == iStudyBasisId && fixierenView.FacultyId == iFacultyId && fixierenView.LicenseProgramId == iLicenseProgramId &&
                                   fixierenView.ObrazProgramId == iObrazProgramId && (gProfileId.HasValue ? fixierenView.ProfileId == gProfileId : true) && fixierenView.IsCel == isCel && fixierenView.IsSecond == isSecond && fixierenView.IsParallel == isParallel && fixierenView.IsReduced == isReduced
                                   orderby fixieren.Number
                                   select new
                                   {
                                       Id = extabit.Id,
                                       Рег_Номер = extabit.RegNum,
                                       Ид_номер = extabit.PersonNum,
                                       ФИО = extabit.FIO,
                                       Сумма_баллов = extabitMarksSum.TotalSum,
                                       Кол_во_оценок = extabitMarksSum.TotalCount,
                                       Подлинники_документов = extabit.HasOriginals ? "Да" : "Нет",
                                       Рейтинговый_коэффициент = extabit.Coefficient,
                                       Конкурс = competition.Name,
                                       Проф_экзамен = hlpabiturientProf.Prof,
                                       Доп_экзамен = hlpabiturientProfAdd.ProfAdd,
                                       comp = competition.Id == 1 ? 1 : (competition.Id == 2 || competition.Id == 7) && extperson.Privileges > 0 ? 2 : 3,
                                       noexamssort = competition.Id == 1 ? extabit.Coefficient : 0
                                   }).ToList().Distinct().Select(x =>
                                       new
                                       {
                                           Id = x.Id.ToString(),
                                           Рег_Номер = x.Рег_Номер,
                                           Ид_номер = x.Ид_номер,
                                           ФИО = x.ФИО,
                                           Сумма_баллов = x.Сумма_баллов,
                                           Кол_во_оценок = x.Кол_во_оценок,
                                           Подлинники_документов = x.Подлинники_документов,
                                           Рейтинговый_коэффициент = x.Рейтинговый_коэффициент,
                                           Конкурс = x.Конкурс,
                                           Проф_экзамен = x.Проф_экзамен,
                                           Доп_экзамен = x.Доп_экзамен,
                                           comp = x.comp,
                                           noexamssort = x.noexamssort
                                       }
                                   );

                        foreach (var v in lst)
                        {
                            ++counter;
                            t.AddCell(new Phrase(counter.ToString(), font));
                            t.AddCell(new Phrase(v.Рег_Номер, font));
                            t.AddCell(new Phrase(v.Ид_номер, font));
                            t.AddCell(new Phrase(v.ФИО, font));
                            t.AddCell(new Phrase(v.Сумма_баллов.ToString(), font));
                            t.AddCell(new Phrase(v.Подлинники_документов, font));
                            t.AddCell(new Phrase(v.Рейтинговый_коэффициент.ToString(), font));
                            t.AddCell(new Phrase(v.Конкурс, font));
                            t.AddCell(new Phrase(v.Проф_экзамен.ToString(), font));
                            t.AddCell(new Phrase(v.Доп_экзамен.ToString(), font));
                        }
                    }

                    document.Add(t);

                    //FOOTER
                    p = new Paragraph();
                    p.SpacingBefore = 30f;
                    p.Alignment = Element.ALIGN_JUSTIFIED;
                    p.FirstLineIndent = firstLineIndent;
                    p.Add(new Phrase("Основание:", new Font(bfTimes, 12, Font.BOLD)));
                    p.Add(new Phrase(" личные заявления, результаты вступительных испытаний, документы, подтверждающие право на поступление без вступительных испытаний или внеконкурсное зачисление.", font));
                    document.Add(p);


                    p = new Paragraph(30f);
                    p.KeepTogether = true;
                    p.Add(new Paragraph("Ответственный секретарь по приему документов по группе направлений:", font));
                    p.Add(new Paragraph("Заместитель начальника управления - советник проректора по группе направлений:", font));
                    //p.Add(new Paragraph("Ответственный секретарь приемной комиссии:", font));

                    document.Add(p);


                    p = new Paragraph(30f);
                    p.Add(new Phrase("В." + iFacultyId.ToString() + "." + docNum, font));
                    document.Add(p);
                    document.Close();



                    Process pr = new Process();

                    pr.StartInfo.Verb = "Open";
                    pr.StartInfo.FileName = string.Format(savePath);
                    pr.Start();

                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintEntryView(string protocolId, string savePath)
        {
            FileStream fileS = null;
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    Guid? protId = new Guid(protocolId);
                    var prot = (from pr in context.extProtocol
                                where pr.Id == protId
                                select pr).FirstOrDefault();

                    string docNum = prot.Number.ToString();
                    DateTime docDate = prot.Date.Value.Date;
                    string form = prot.StudyFormRodName;
                    string form2 = prot.StudyFormRodName;
                    string facDat = prot.FacultyDatName;

                    string basisId = prot.StudyBasisId.ToString();
                    string basis = string.Empty;

                    bool? isSec = prot.IsSecond;
                    bool? isReduced = prot.IsReduced;
                    bool? isParallel = prot.IsParallel;
                    bool? isList = prot.IsListener;

                    string profession = (from extabit in context.extAbitAspirant
                                         join extentryView in context.extEntryView on extabit.Id equals extentryView.AbiturientId
                                         where extentryView.Id == protId
                                         select extabit.LicenseProgramName
                                  ).FirstOrDefault();

                    string professionCode = (from extabit in context.extAbitAspirant
                                             join extentryView in context.extEntryView on extabit.Id equals extentryView.AbiturientId
                                             where extentryView.Id == protId
                                             select extabit.LicenseProgramCode
                                  ).FirstOrDefault();

                    switch (basisId)
                    {
                        case "1":
                            basis = "обучение за счет средств федерального бюджета";
                            break;
                        case "2":
                            basis = "обучение по договорам с оплатой стоимости обучения";
                            break;
                    }

                    string list = string.Empty, sec = string.Empty;

                    string copyDoc = "оригиналы";
                    if (isList.HasValue && isList.Value)
                    {
                        list = " в качестве слушателя";
                        copyDoc = "заверенные ксерокопии";
                    }

                    if (isReduced.HasValue && isReduced.Value)
                        sec = " (сокращенной)";
                    if (isParallel.HasValue && isParallel.Value)
                        sec = " (параллельной)";

                    Document document = new Document(PageSize.A4, 50, 50, 50, 50);

                    using (fileS = new FileStream(savePath, FileMode.Create))
                    {

                        BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font font = new Font(bfTimes, 12);

                        PdfWriter writer = PdfWriter.GetInstance(document, fileS);
                        document.Open();

                        float firstLineIndent = 30f;
                        //HEADER
                        Paragraph p = new Paragraph("Правительство Российской Федерации", new Font(bfTimes, 12, Font.BOLD));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("Федеральное государственное бюджетное образовательное учреждение", new Font(bfTimes, 12));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("высшего профессионального образования", new Font(bfTimes, 12));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("САНКТ-ПЕТЕРБУРГСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ", new Font(bfTimes, 12, Font.BOLD));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("ПРЕДСТАВЛЕНИЕ", new Font(bfTimes, 20, Font.BOLD));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph(string.Format("От {0} г. № {1}", Util.GetDateString(docDate, true, true), docNum), font);
                        p.SpacingBefore = 10f;
                        document.Add(p);

                        p = new Paragraph(10f);
                        p.Add(new Paragraph("по " + facDat, font));

                        string bakspec = "", naprspec = "", naprspecRod = "", profspec = "", naprobProgRod = "", educDoc = ""; ;

                        naprobProgRod = "образовательной программе";
                        naprspec = "направление";
                        naprspecRod = "направлению";

                        educDoc = "о высшем профессиональном образовании";
                        
                        p.Add(new Paragraph("по основной образовательной программе послевузовского профессионального образования (аспирантура)", font));
                        p.Add(new Paragraph(("по " + form + " форме обучения,").ToLower(), font));
                        p.Add(new Paragraph(basis, font));
                        p.IndentationLeft = 320;
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Paragraph("О зачислении в аспирантуру", font));
                        //p.Add(new Paragraph("граждан Российской Федерации", font));
                        p.SpacingBefore = 10f;
                        document.Add(p);

                        p = new Paragraph("В соответствии с Положением о подготовке научно-педагогических и научных кадров в системе послевузовского профессионального образования в Российской Федерации, утвержденного приказом Минобразования Российской Федерации от 27.03.1998 № 814, Правилами приема на основные образовательные программы послевузовского профессионального образования (программы подготовки научно-педагогических кадров в аспирантуре) Санкт-Петербургского государственного университета в 2013 году", font);
                        p.SpacingBefore = 10f;
                        p.Alignment = Element.ALIGN_JUSTIFIED;
                        p.FirstLineIndent = firstLineIndent;
                        document.Add(p);

                        p = new Paragraph(string.Format("Представить на рассмотрение Приемной комиссии СПбГУ по вопросу зачисления c 01.09.2013 года на 1 курс{2} с освоением основной{3} образовательной программы подготовки {0} по {1} форме обучения следующих граждан, успешно выдержавших вступительные испытания:", bakspec, form2, list, sec), font);
                        p.FirstLineIndent = firstLineIndent;
                        p.Alignment = Element.ALIGN_JUSTIFIED;
                        p.SpacingBefore = 20f;
                        document.Add(p);


                        string curSpez = "-";
                        string curObProg = "-";
                        string curHeader = "-";

                        int counter = 0;

                        using (PriemEntities ctx = new PriemEntities())
                        {
                            var lst = (from extabit in ctx.extAbitAspirant
                                       join extentryView in ctx.extEntryView on extabit.Id equals extentryView.AbiturientId
                                       join extperson in ctx.extPersonAll on extabit.PersonId equals extperson.Id
                                       join competition in ctx.Competition on extabit.CompetitionId equals competition.Id
                                       join entryHeader in ctx.EntryHeader on extentryView.EntryHeaderId equals entryHeader.Id into entryHeader2
                                       from entryHeader in entryHeader2.DefaultIfEmpty()
                                       join extabitMarksSum in ctx.extAbitMarksSum on extabit.Id equals extabitMarksSum.Id into extabitMarksSum2
                                       from extabitMarksSum in extabitMarksSum2.DefaultIfEmpty()
                                       where extentryView.Id == protId
                                       orderby extabit.ObrazProgramName, extabit.ProfileName, entryHeader.Id, extperson.FIO
                                       select new
                                       {
                                           Id = extabit.Id,
                                           Рег_Номер = extabit.RegNum,
                                           Ид_номер = extperson.PersonNum,
                                           TotalSum = extabitMarksSum.TotalSum,
                                           ФИО = extperson.FIO,
                                           LicenseProgramName = extabit.LicenseProgramName,
                                           LicenseProgramCode = extabit.LicenseProgramCode,
                                           ObrazProgram = extabit.ObrazProgramName,
                                           ObrazProgramId = extabit.ObrazProgramId,
                                           ObrazProgramCrypt = extabit.ObrazProgramCrypt,
                                           ProfileName = extabit.ProfileName,
                                           EntryHeaderId = entryHeader.Id,
                                           EntryHeaderName = entryHeader.Name
                                       }).ToList().Distinct().Select(x =>
                                           new
                                           {
                                               Id = x.Id.ToString(),
                                               Рег_Номер = x.Рег_Номер,
                                               Ид_номер = x.Ид_номер,
                                               TotalSum = x.TotalSum,
                                               ФИО = x.ФИО,
                                               LicenseProgramName = x.LicenseProgramName,
                                               x.LicenseProgramCode,
                                               ObrazProgram = x.ObrazProgram,
                                               ObrazProgramId = x.ObrazProgramId,
                                               ObrazProgramCrypt = x.ObrazProgramCrypt,
                                               ProfileName = x.ProfileName,
                                               EntryHeaderId = x.EntryHeaderId,
                                               EntryHeaderName = x.EntryHeaderName
                                           }
                                       );

                            foreach (var v in lst)
                            {
                                ++counter;
                                string obProg = v.ObrazProgram;
                                string obProgCrypt = v.ObrazProgramCrypt;
                                string obProgId = v.ObrazProgramId.ToString();

                                if (obProgId != curObProg)
                                {
                                    p = new Paragraph();
                                    p.Add(new Paragraph(string.Format("{3}по {0} {1} \"{2}\"", naprspecRod, v.LicenseProgramCode, v.LicenseProgramName, curObProg == "-" ? "" : "\r\n"), font));

                                    if (!string.IsNullOrEmpty(obProg))
                                        p.Add(new Paragraph(string.Format("по {0} {1} \"{2}\"", naprobProgRod, obProgCrypt, obProg), font));

                                    //string spez = v.ProfileName;
                                    //if (spez != curSpez)
                                    //{
                                    //    if (!string.IsNullOrEmpty(spez) && spez != "нет")
                                    //        p.Add(new Paragraph(string.Format("по {0} \"{1}\"", profspec, spez), font));

                                    //    curSpez = spez;
                                    //    curHeader = "NULL";
                                    //}

                                    p.IndentationLeft = 40;
                                    document.Add(p);

                                    curObProg = obProgId;
                                    curHeader = "NULL";
                                }
                                else
                                {
                                    //string spez = v.ProfileName;
                                    //if (spez != curSpez && spez != "нет")
                                    //{
                                    //    p = new Paragraph();
                                    //    p.Add(new Paragraph(string.Format("{3}по {0} {1} \"{2}\"", naprspecRod, professionCode, profession, curObProg == "-" ? "" : "\r\n"), font));

                                    //    if (!string.IsNullOrEmpty(obProg))
                                    //        p.Add(new Paragraph(string.Format("по {0} \"{1}\"", naprobProgRod, obProg), font));

                                    //    if (!string.IsNullOrEmpty(spez))
                                    //        p.Add(new Paragraph(string.Format("по {0} \"{1}\"", profspec, spez), font));

                                    //    p.IndentationLeft = 40;
                                    //    document.Add(p);

                                    //    curSpez = spez;
                                    //    curHeader = "NULL";
                                    //}
                                }

                                string header = v.EntryHeaderName;
                                if (header != curHeader)
                                {
                                    p = new Paragraph();
                                    p.Add(new Paragraph(string.Format("{0}:", header), font));
                                    p.IndentationLeft = 40;
                                    document.Add(p);

                                    curHeader = header;
                                }

                                p = new Paragraph();
                                p.Add(new Paragraph(string.Format("{0}. {1} {2}", counter, v.ФИО, v.TotalSum.ToString()), font));
                                p.IndentationLeft = 60;
                                document.Add(p);
                            }
                        }

                        //FOOTER
                        p = new Paragraph();
                        p.SpacingBefore = 30f;
                        p.Alignment = Element.ALIGN_JUSTIFIED;
                        p.FirstLineIndent = firstLineIndent;
                        p.Add(new Phrase("ОСНОВАНИЕ:", new Font(bfTimes, 12)));
                        p.Add(new Phrase(string.Format(" личные заявления, протоколы вступительных испытаний, {0} документов государственного образца {1}.", copyDoc, educDoc), font));
                        document.Add(p);

                        p = new Paragraph();
                        p.SpacingBefore = 30f;
                        p.KeepTogether = true;
                        p.Add(new Paragraph("Ответственный секретарь", font));
                        p.Add(new Paragraph("комиссии по приему документов СПбГУ                                                                                          ", font));
                        document.Add(p);

                        p = new Paragraph();
                        p.SpacingBefore = 30f;
                        p.Add(new Paragraph("Заместитель начальника управления - ", font));
                        p.Add(new Paragraph("советник проректора по направлениям", font));
                        document.Add(p);

                        document.Close();

                        Process pr = new Process();

                        pr.StartInfo.Verb = "Open";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();

                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintOrder(Guid protocolId, bool isRus, bool isCel)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\EntryOrder1.dot", MainClass.dirTemplates));
                TableDoc td = wd.Tables[0];
                
                string docNum;
                DateTime docDate;
                string formId;
                string facDat;

                bool? isSec;
                bool? isParallel;
                bool? isReduced;
                bool? isList;

                string basisId;
                string basis = string.Empty;
                string form = string.Empty;
                string form2 = string.Empty;

                string LicenseProgramName;
                string LicenseProgramCode;

                string DateStart = "01.09.2014";
                string DateOfFinish = "30.06.2015";

                using (PriemEntities ctx = new PriemEntities())
                {

                    docNum = (from protocol in ctx.Protocol
                              where protocol.Id == protocolId
                              select protocol.Number).FirstOrDefault();

                    docDate = (DateTime)(from protocol in ctx.Protocol
                                         where protocol.Id == protocolId
                                         select protocol.Date).FirstOrDefault();

                    formId = (from protocol in ctx.Protocol
                              join studyForm in ctx.StudyForm on protocol.StudyFormId equals studyForm.Id
                              where protocol.Id == protocolId
                              select studyForm.Id).FirstOrDefault().ToString();

                    facDat = (from protocol in ctx.Protocol
                              join sP_Faculty in ctx.SP_Faculty on protocol.FacultyId equals sP_Faculty.Id
                              where protocol.Id == protocolId
                              select sP_Faculty.DatName).FirstOrDefault();

                    isSec = (from protocol in ctx.Protocol
                             where protocol.Id == protocolId
                             select protocol.IsSecond).FirstOrDefault();

                    isParallel = (from protocol in ctx.Protocol
                                  where protocol.Id == protocolId
                                  select protocol.IsParallel).FirstOrDefault();

                    isReduced = (from protocol in ctx.Protocol
                                 where protocol.Id == protocolId
                                 select protocol.IsReduced).FirstOrDefault();

                    isList = (from protocol in ctx.Protocol
                              where protocol.Id == protocolId
                              select protocol.IsListener).FirstOrDefault();

                    basisId = (from protocol in ctx.Protocol
                               join studyBasis in ctx.StudyBasis on protocol.StudyBasisId equals studyBasis.Id
                               where protocol.Id == protocolId
                               select studyBasis.Id).FirstOrDefault().ToString();

                    LicenseProgramName = (from entry in ctx.Entry
                                          join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                                          where extentryView.Id == protocolId
                                          select entry.SP_LicenseProgram.Name).FirstOrDefault();

                    LicenseProgramCode = (from entry in ctx.Entry
                                          join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                                          where extentryView.Id == protocolId
                                          select entry.SP_LicenseProgram.Code).FirstOrDefault();
                }
                string studySrok = string.Empty, studyFinish = string.Empty;
                switch (formId)
                {
                    case "1":
                        form = "очная форма обучения";
                        form2 = "очной";
                        studySrok = "1";
                        studyFinish = "31.08.2016";
                        break;
                    case "2":
                        form = "заочная форма обучения";
                        form2 = "заочной";
                        studySrok = "1";
                        studyFinish = "31.08.2017";
                        break;
                }

                //string educDoc = "об образовании";
                string copyDoc = "оригиналы";
                if (isList.HasValue && isList.Value)
                {
                    copyDoc = "заверенные ксерокопии";
                }

                string dogovorDoc = "";
                switch (basisId)
                {
                    case "1":
                        basis = "за счет бюджетных ассигнований федерального бюджета";
                        dogovorDoc = ", оригиналы документов установленного образца об образовании";
                        break;
                    case "2":
                        basis = "по договорам об образовании";
                        dogovorDoc = ", договоры об образовании";
                        break;
                }


                wd.SetFields("Граждан", isRus ? "граждан Российской Федерации" : "иностранных граждан");
                wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                wd.SetFields("Стипендия", (basisId == "2" || formId == "2") ? "" : "и назначении стипендии");
                wd.SetFields("Форма2", form2);
                wd.SetFields("Основа2", basis);
               // wd.SetFields("CopyDoc", copyDoc);
                wd.SetFields("DogovorDoc", dogovorDoc);
                wd.SetFields("Srok", studySrok);
                wd.SetFields("DateStart", DateStart);

                int curRow = 4;
                string curObProg = "-";
                string curCountry = "-";
                string curLPHeader = "-";
                string curHeader = "-";
                string curMotivation = "-";
                string Motivation = string.Empty;
                string stipendia = "", naprobProgRod = "";
                naprobProgRod = "образовательной программе";
                string naprspecRod = "направлению подготовки";

                using (PriemEntities ctx = new PriemEntities())
                {
                    var lst = (from extabit in ctx.extAbitAspirant
                               join extentryView in ctx.extEntryView on extabit.Id equals extentryView.AbiturientId
                               join extperson in ctx.extPersonAll on extabit.PersonId equals extperson.Id
                               join country in ctx.Country on extperson.NationalityId equals country.Id
                               join competition in ctx.Competition on extabit.CompetitionId equals competition.Id
                               join extabitMarksSum in ctx.extAbitMarksSum on extabit.Id equals extabitMarksSum.Id into extabitMarksSum2
                               from extabitMarksSum in extabitMarksSum2.DefaultIfEmpty()
                               join entryHeader in ctx.EntryHeader on extentryView.EntryHeaderId equals entryHeader.Id into entryHeader2
                               from entryHeader in entryHeader2.DefaultIfEmpty()
                               join celCompetition in ctx.CelCompetition on extabit.CelCompetitionId equals celCompetition.Id into celCompetition2
                               from celCompetition in celCompetition2.DefaultIfEmpty()
                               where extentryView.Id == protocolId && (isRus ? extperson.NationalityId == 1 : extperson.NationalityId != 1)
                               orderby celCompetition.TvorName, extabit.ObrazProgramName, extabit.ProfileName, country.NameRod, entryHeader.SortNum, extabit.FIO
                               select new
                               {
                                   Id = extabit.Id,
                                   extperson.Sex,
                                   Рег_Номер = extabit.RegNum,
                                   Ид_номер = extabit.PersonNum,
                                   TotalSum = (extabit.CompetitionId == 8 || extabit.CompetitionId == 1) ? null : extabitMarksSum.TotalSumFiveGrade,
                                   ФИО = extabit.FIO,
                                   CelCompName = celCompetition.TvorName,
                                   LicenseProgramName = extabit.LicenseProgramName,
                                   LicenseProgramCode = extabit.LicenseProgramCode,
                                   ProfileName = extabit.ProfileName,
                                   ObrazProgram = extabit.ObrazProgramName,
                                   ObrazProgramId = extabit.ObrazProgramId,
                                   EntryHeaderId = entryHeader.Id,
                                   SortNum = entryHeader.SortNum,
                                   EntryHeaderName = entryHeader.Name,
                                   NameRod = country.NameRod
                               }).ToList().Distinct().Select(x =>
                                   new
                                   {
                                       Id = x.Id.ToString(),
                                       x.Sex,
                                       Рег_Номер = x.Рег_Номер,
                                       Ид_номер = x.Ид_номер,
                                       TotalSum = x.TotalSum.ToString(),
                                       ФИО = x.ФИО,
                                       CelCompName = x.CelCompName,
                                       LicenseProgramName = x.LicenseProgramName,
                                       LicenseProgramCode = x.LicenseProgramCode,
                                       ProfileName = x.ProfileName,
                                       ObrazProgram = x.ObrazProgram.Replace("(очно-заочная)", "").Replace(" ВВ", ""),
                                       ObrazProgramId = x.ObrazProgramId,
                                       EntryHeaderId = x.EntryHeaderId,
                                       SortNum = x.SortNum,
                                       EntryHeaderName = x.EntryHeaderName,
                                       NameRod = x.NameRod
                                   }
                               );
                    //wd.SetFields("CountType", lst.Count() == 1 ? "Нижеуказанное лицо" : "Нижеуказанных лиц");
                    //wd.SetFields("DateFinish", studyFinish);

                    int pos = 0;
                    bool bFirstRun = true;

                    foreach (var v in lst)
                    {
                        pos++;
                        string header = v.EntryHeaderName;

                        if (!isCel && !bFirstRun)
                        {
                            if (header != curHeader)
                            {
                                td.AddRow(1);
                                curRow++;
                                td[0, curRow] = string.Format("\t{0}:", header);

                                curHeader = header;
                            }
                        }

                        bFirstRun = false;

                        string LP = v.LicenseProgramName;
                        string LPCode = v.LicenseProgramCode;
                        if (curLPHeader != LP)
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("{3}\tпо {0} {1} \"{2}\"", naprspecRod, LPCode, LP, curObProg == "-" ? "" : "\r\n");
                            curLPHeader = LP;
                        }

                        int Code = 0;
                        string Num = LPCode.Substring(LPCode.IndexOf('.') + 1);
                        Num = Num.Substring(0, Num.IndexOf('.'));
                        if (int.TryParse(Num, out Code))
                        {
                            if (Code < 38)
                                stipendia = "6330";
                            else
                                stipendia = "2637";
                        }

                        string ObrazProgramId = v.ObrazProgramId.ToString();
                        string obProg = v.ObrazProgram;
                        string obProgCode = (from entry in ctx.Entry
                                             where entry.ObrazProgramId == v.ObrazProgramId
                                             select entry.StudyLevel.Acronym + "." + entry.SP_ObrazProgram.Number + "." + MainClass.PriemYear).FirstOrDefault();

                        if (ObrazProgramId != curObProg)
                        {
                            if (obProg != String.Empty)
                            {   
                                td.AddRow(1);
                                curRow++;
                                td[0, curRow] = string.Format("\tпо {0} {1} \"{2}\"", naprobProgRod, obProgCode, obProg);
                            }
                            curObProg = ObrazProgramId;
                            if (!isCel)
                            {
                                if (header != curHeader)
                                {
                                    td.AddRow(1);
                                    curRow++;
                                    td.AddRow(1);
                                    curRow++;
                                    td[0, curRow] = string.Format("\t{0}:", header);
                                    td.AddRow(1);
                                    curRow++;
                                    curHeader = header;
                                }
                            }
                        }

                        if (!isRus)
                        {
                            string country = v.NameRod;
                            if (country != curCountry)
                            {
                                td.AddRow(1);
                                curRow++;
                                td[0, curRow] = string.Format("\r\n граждан {0}:", country);

                                curCountry = country;
                            }
                        }



                        string balls = v.TotalSum;
                        string ballToStr = " балл";

                        if (balls.Length == 0)
                            ballToStr = "";
                        else if (balls.EndsWith("1"))
                        { 
                            if (balls.EndsWith("11"))
                                ballToStr +="ов";
                            else
                                ballToStr += ""; 
                        }
                        else if (balls.EndsWith("2") || balls.EndsWith("3") || balls.EndsWith("4"))
                        {
                            if ((balls.EndsWith("12") || balls.EndsWith("13") || balls.EndsWith("14")))
                                ballToStr += "ов";
                            else
                                ballToStr += "а";
                        }
                        else
                            ballToStr += "ов";

                        td.AddRow(1);
                        curRow++;
                        //td[0, curRow] = string.Format("1.{0}. {1}, {3} {2}", (pos + "." + subpos), v.ФИО, v.NameRod, v.Sex ? "гражданин" : "гражданка");
                        td[0, curRow] = string.Format("\t\t1.{0}. {1} {2} ", pos, v.ФИО + ',', balls + ballToStr);//, string.IsNullOrEmpty(Motivation) ? "" : ("\n\n\t\t" + tmpMotiv + "\n"));
                    }
                }

                if (!string.IsNullOrEmpty(curMotivation) && isCel)
                    td[0, curRow] += "\n\t\t" + curMotivation + "\n";


                if (basisId != "2" && formId == "1")//платникам и всем очно-заочникам стипендия не платится
                {
                    td.AddRow(1);
                    curRow++;
                    td[0, curRow] = "\r\n2.  Назначить лицам, указанным в п. 1 настоящего приказа, стипендию в размере  " + stipendia + @" рублей ежемесячно с " + DateStart + " по " + DateOfFinish + ".";
                }
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

        public static void PrintOrderReview(Guid protocolId, bool isRus)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\EntryOrderList.dot", MainClass.dirTemplates));

                string query = @"SELECT extAbitAspirant.Id as Id, extAbitAspirant.RegNum as Рег_Номер, 
                     extAbitAspirant.PersonNum as 'Ид. номер', extAbitMarksSum.TotalSum, 
                     extAbitAspirant.FIO as ФИО, 
                     extAbitAspirant.LicenseProgramCode + ' ' + extAbitAspirant.LicenseProgramName, extAbitAspirant.ProfileName, 
                     replace(replace(extAbitAspirant.ObrazProgramName, '(очно-заочная)', ''), ' ВВ', '')  as ObrazProgram, extAbitAspirant.ObrazProgramCrypt, 
                     EntryHeader.Id as EntryHeaderId, EntryHeader.Name as EntryHeaderName, Country.NameRod,
                     extAbitMarksSum.TotalSumFiveGrade,
                     extAbitAspirant.CompetitionId,
                     extEntryView.SignerName,
                     extEntryView.SignerPosition
                     FROM ed.extAbitAspirant 
                     INNER JOIN ed.extEntryView ON extEntryView.AbiturientId=extAbitAspirant.Id 
                     INNER JOIN ed.Person ON Person.Id = extAbitAspirant.PersonId 
                     INNER JOIN ed.Country ON Person.NationalityId = Country.Id 
                     INNER JOIN ed.Competition ON Competition.Id = extAbitAspirant.CompetitionId 
                     LEFT JOIN ed.EntryHeader ON EntryHeader.Id = extEntryView.EntryHeaderId 
                     LEFT JOIN ed.extAbitMarksSum ON extAbitAspirant.Id = extAbitMarksSum.Id";

                string where = " WHERE extEntryView.Id = @protocolId ";
                where += " AND Person.NationalityId" + (isRus ? "=1 " : "<>1 ");
                string orderby = " ORDER BY ObrazProgram, extAbitAspirant.ProfileName, NameRod, EntryHeader.Id, ФИО ";
                SortedList<string, object> slDel = new SortedList<string, object>();

                slDel.Add("@protocolId", protocolId);

                DataSet ds = MainClass.Bdc.GetDataSet(query + where + orderby, slDel);

                DateTime protocolDate = (DateTime)MainClass.Bdc.GetValue(string.Format("SELECT Protocol.Date FROM ed.Protocol WHERE Protocol.Id='{0}'", protocolId));
                string protocolNum = MainClass.Bdc.GetStringValue(string.Format("SELECT Protocol.Number FROM ed.Protocol WHERE Protocol.Id='{0}'", protocolId));

                string DateStart = "01.09.2014";
                string DateOfFinish = "30.06.2015";
                string docNum = "НОМЕР";
                string docDate = "ДАТА";
                DateTime tempDate;
                if (isRus)
                {
                    docNum = MainClass.Bdc.GetStringValue(string.Format("SELECT OrderNum FROM ed.OrderNumbers WHERE ProtocolId='{0}'", protocolId));
                    DateTime.TryParse(MainClass.Bdc.GetStringValue(string.Format("SELECT OrderDate FROM ed.OrderNumbers WHERE ProtocolId='{0}'", protocolId)), out tempDate);

                    docDate = tempDate.ToShortDateString();
                }
                else
                {
                    docNum = MainClass.Bdc.GetStringValue(string.Format("SELECT OrderNumFor FROM ed.OrderNumbers WHERE ProtocolId='{0}'", protocolId));

                    DateTime.TryParse(MainClass.Bdc.GetStringValue(string.Format("SELECT OrderDateFor FROM ed.OrderNumbers WHERE ProtocolId='{0}'", protocolId)), out tempDate);

                    docDate = tempDate.ToShortDateString();
                }

                string formId = MainClass.Bdc.GetStringValue("SELECT StudyForm.Id FROM ed.Protocol INNER JOIN ed.StudyForm ON Protocol.StudyFormId=StudyForm.Id WHERE Protocol.Id= @protocolId", slDel);
                string facDat = MainClass.Bdc.GetStringValue("SELECT SP_Faculty.DatName FROM ed.Protocol INNER JOIN ed.SP_Faculty ON Protocol.FacultyId=SP_Faculty.Id WHERE Protocol.Id= @protocolId", slDel);

                string basisId = MainClass.Bdc.GetStringValue("SELECT StudyBasis.Id FROM ed.Protocol INNER JOIN ed.StudyBasis ON ed.Protocol.StudyBasisId=StudyBasis.Id WHERE Protocol.Id= @protocolId", slDel);
                string basis = string.Empty;
                string form = string.Empty;
                string form2 = string.Empty;

                string profession = MainClass.Bdc.GetStringValue("SELECT LicenseProgramName FROM ed.qEntry INNER JOIN ed.extEntryView ON qEntry.LicenseProgramId=extEntryView.LicenseProgramId WHERE extEntryView.Id= @protocolId", slDel);
                string professionCode = MainClass.Bdc.GetStringValue("SELECT LicenseProgramCode FROM ed.qEntry INNER JOIN ed.extEntryView ON qEntry.LicenseProgramId=extEntryView.LicenseProgramId WHERE extEntryView.Id= @protocolId", slDel);

                string dogovorDoc = "";
                switch (basisId)
                {
                    case "1":
                        basis = "за счет бюджетных ассигнований федерального бюджета";
                        dogovorDoc = ", оригиналы документов установленного образца об образовании";
                        break;
                    case "2":
                        basis = "по договорам об образовании";
                        dogovorDoc = ", договоры об образовании";
                        break;
                }
                string studySrok = string.Empty, studyFinish = string.Empty;
                switch (formId)
                {
                    case "1":
                        form = "очная форма обучения";
                        form2 = "очной ";
                        studySrok = "1";
                        studyFinish = "31.08.2016";
                        break;
                    case "2":
                        form = "заочная форма обучения";
                        form2 = "заочной";
                        studySrok = "1";
                        studyFinish = "31.08.2017";
                        break;
                }

                string  naprspecRod = "", profspec = "";
                string naprobProgRod = "образовательной программе"; ;

                if (MainClass.dbType == PriemType.PriemAspirant)
                {
                    naprspecRod = "направлению";
                    profspec = "магистерской программе";
                }

                int curRow = 5, counter = 0;
                TableDoc td = null;

                int Code = 0;
                string stipendia = "";
                string Num = professionCode.Substring(professionCode.IndexOf('.') + 1);
                Num = Num.Substring(0, Num.IndexOf('.'));
                if (int.TryParse(Num, out Code))
                {
                    if (Code < 38)
                        stipendia = "6330";
                    else
                        stipendia = "2637";
                }

                foreach (DataRow r in ds.Tables[0].Rows)
                {

                    wd.InsertAutoTextInEnd("выписка", true);
                    wd.GetLastFields(12);
                    td = wd.Tables[counter];

                    wd.SetFields("Граждан", isRus ? "граждан РФ" : "иностранных граждан");
                    wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                    wd.SetFields("Стипендия", (basisId == "2" || formId == "2") ? "" : "и назначении стипендии");
                    wd.SetFields("Форма2", form2);
                    wd.SetFields("Основа2", basis);

                    wd.SetFields("ПриказДата", docDate);
                    wd.SetFields("ПриказНомер", "№ " + docNum);

                    wd.SetFields("DogovorDoc", dogovorDoc);
                    wd.SetFields("Srok", studySrok);
                    wd.SetFields("DateStart", DateStart);

                    wd.SetFields("SignerPosition", r["SignerPosition"].ToString());
                    wd.SetFields("SignerName", r["SignerName"].ToString());

                    string curLPHeader = "-";
                    string curSpez = "-";
                    string curObProg = "-";
                    string curHeader = "-";
                    string curCountry = "-";

                    ++counter;

                    string LP = profession;
                    string LPCode = professionCode;
                    if (curLPHeader != LP)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("{3}\tпо {0} {1} \"{2}\"", naprspecRod, LPCode, LP, curObProg == "-" ? "" : "\r\n");
                        curLPHeader = LP;
                    }

                    string obProg = r["ObrazProgram"].ToString();
                    string obProgCode = r["ObrazProgramCrypt"].ToString();
                    if (obProg != curObProg)
                    {
                        if (!string.IsNullOrEmpty(obProg))
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("\tпо {0} {1} \"{2}\"", naprobProgRod, obProgCode, obProg);
                        }

                        string spez = r["ProfileName"].ToString();

                        if (!string.IsNullOrEmpty(spez) && spez != "нет")
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("\t {0} \"{1}\"", profspec, spez);
                        }

                        curSpez = spez;

                        curObProg = obProg;
                    }
                    else
                    {
                        string spez = r["ProfileName"].ToString();
                        if (spez != curSpez)
                        {
                            if (!string.IsNullOrEmpty(spez) && spez != "нет")
                            {
                                td.AddRow(1);
                                curRow++;
                                td[0, curRow] = string.Format("\t {0} \"{1}\"", profspec, spez);
                            }

                            curSpez = spez;
                        }
                    }

                    if (!isRus)
                    {
                        string country = r["NameRod"].ToString();
                        if (country != curCountry)
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("\r\n граждан {0}:", country);

                            curCountry = country;
                        }
                    }

                    string header = r["EntryHeaderName"].ToString();
                    if (header != curHeader)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("\t{0}:", header);

                        curHeader = header;
                    }

                    string balls ="";
                    if ((r["CompetitionId"].ToString() == "8") || (r["CompetitionId"].ToString() == "1"))
                        balls = "";
                    else
                        balls = r["TotalSumFiveGrade"].ToString();
                    string ballToStr = " балл";

                    if (balls.Length == 0)
                        ballToStr = "";
                    else if (balls.EndsWith("1"))
                    {
                        if (balls.EndsWith("1"))
                            ballToStr += "ов";
                        else
                            ballToStr += ""; 

                    }
                    else if (balls.EndsWith("2") || balls.EndsWith("3") || balls.EndsWith("4"))
                    {
                        if (balls.EndsWith("2") || balls.EndsWith("3") || balls.EndsWith("4"))
                            ballToStr += "ов";
                        else
                            ballToStr += "а";
                    }
                    else
                        ballToStr += "ов";

                    td.AddRow(1);
                    curRow++;
                    td[0, curRow] = string.Format("\t\t{0} {1}", r["ФИО"].ToString(), balls + ballToStr);

                    if (basisId != "2" && formId != "2")
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = "\r\n2.      Назначить указанным лицам стипендию в размере 1200 рублей ежемесячно до 31 января 2013 г.";
                        td[0, curRow] = "\r\n2.  Назначить лицам, указанным в п. 1 настоящего приказа, стипендию в размере  " + stipendia + @" рублей ежемесячно с " + DateStart + " по " + DateOfFinish + ".";

                    }
                }

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

        public static void PrintDisEntryOrder(string protocolId, bool isRus)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\DisEntryOrder.dot", MainClass.dirTemplates));
                TableDoc td = wd.Tables[0];

                string query = @"SELECT ed.extAbitAspirant.Id as Id, ed.extAbitAspirant.RegNum as Рег_Номер, 
                    ed.extPersonAspirant.PersonNum as 'Ид. номер', ed.extAbitMarksSum.TotalSum, 
                    ed.extPersonAspirant.FIO  as ФИО, 
                    ed.extAbitAspirant.LicenseProgramName, ed.extAbitAspirant.ProfileName as Specialization, ed.Country.NameRod 
                     FROM ed.extAbitAspirant 
                     INNER JOIN ed.extDisEntryView ON ed.extDisEntryView.AbiturientId=ed.extAbitAspirant.Id 
                     INNER JOIN ed.extPersonAspirant ON ed.extPersonAspirant.Id = ed.extAbitAspirant.PersonId 
                     INNER JOIN ed.Country ON ed.extPersonAspirant.NationalityId = ed.Country.Id
                     INNER JOIN ed.Competition ON ed.Competition.Id = ed.extAbitAspirant.CompetitionId 
                     LEFT JOIN ed.extAbitMarksSum ON ed.extAbitAspirant.Id = ed.extAbitMarksSum.Id";

                string where = " WHERE ed.extDisEntryView.Id = @protocolId AND extDisEntryView.StudyLevelGroupId=@StudyLevelGroupId";
                where += " AND ed.extPersonAspirant.NationalityId" + (isRus ? "=1 " : "<>1 ");
                string orderby = " ORDER BY ed.extAbitAspirant.ProfileName, NameRod ,ФИО ";

                SortedList<string, object> slDel = new SortedList<string, object>();

                slDel.Add("@protocolId", protocolId);
                slDel.Add("@StudyLevelGroupId", MainClass.studyLevelGroupId);

                DataSet ds = MainClass.Bdc.GetDataSet(query + where + orderby, slDel);

                string entryProtocolId = MainClass.Bdc.GetStringValue("SELECT ed.extProtocol.Id FROM ed.extDisEntryView INNER JOIN ed.extProtocol ON ed.extDisEntryView.AbiturientId=ed.extProtocol.AbiturientId WHERE ed.extDisEntryView.Id=@protocolId AND ed.extProtocol.ProtocolTypeId=4 AND ed.extprotocol.isold=0 ", slDel);

                string docNum = "НОМЕР";
                string docDate = "ДАТА";
                DateTime tempDate;
                if (isRus)
                {
                    docNum = MainClass.Bdc.GetStringValue(string.Format("SELECT OrderNum FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId));
                    DateTime.TryParse(MainClass.Bdc.GetStringValue(string.Format("SELECT OrderDate FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId)), out tempDate);

                    docDate = tempDate.ToShortDateString();
                }
                else
                {
                    docNum = MainClass.Bdc.GetStringValue(string.Format("SELECT OrderNumFor FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId));

                    DateTime.TryParse(MainClass.Bdc.GetStringValue(string.Format("SELECT OrderDateFor FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId)), out tempDate);

                    docDate = tempDate.ToShortDateString();
                }

                string formId = MainClass.Bdc.GetStringValue("SELECT StudyForm.Id FROM ed.Protocol INNER JOIN StudyForm ON Protocol.StudyFormId=StudyForm.Id WHERE Protocol.Id= @protocolId", slDel);
                string facDat = MainClass.Bdc.GetStringValue("SELECT SP_Faculty.DatName FROM ed.Protocol INNER JOIN SP_Faculty ON Protocol.FacultyId=SP_Faculty.Id WHERE Protocol.Id= @protocolId", slDel);

                string basisId = MainClass.Bdc.GetStringValue("SELECT StudyBasis.Id FROM ed.Protocol INNER JOIN StudyBasis ON Protocol.StudyBasisId=StudyBasis.Id WHERE Protocol.Id= @protocolId", slDel);
                string basis = string.Empty;
                string form = string.Empty;
                string form2 = string.Empty;

                bool? isSec = (bool?)MainClass.Bdc.GetValue(string.Format("SELECT IsSecond FROM ed.Protocol  WHERE Protocol.Id= '{0}'", protocolId));
                bool? isReduced = (bool?)MainClass.Bdc.GetValue(string.Format("SELECT IsReduced FROM ed.Protocol  WHERE Protocol.Id= '{0}'", protocolId));
                bool? isList = (bool?)MainClass.Bdc.GetValue(string.Format("SELECT IsListener FROM ed.Protocol WHERE Protocol.Id='{0}'", protocolId));

                string list = string.Empty, sec = string.Empty;

                if (isList.HasValue && isList.Value)
                    list = " в качестве слушателя";

                if (isReduced.HasValue && isReduced.Value)
                    sec = " (сокращенной)";

                if (isSec.HasValue && isSec.Value)
                    sec = " (для лиц с высшим образованием)";

                string LicenseProgramName = MainClass.Bdc.GetStringValue("SELECT qEntry.LicenseProgramName FROM ed.qEntry INNER JOIN ed.extDisEntryView ON qEntry.LicenseProgramId=extDisEntryView.LicenseProgramId WHERE extDisEntryView.Id= @protocolId AND extDisEntryView.StudyLevelGroupId=@StudyLevelGroupId", slDel);
                string LicenseProgramCode = MainClass.Bdc.GetStringValue("SELECT qEntry.LicenseProgramCode FROM ed.qEntry INNER JOIN ed.extDisEntryView ON qEntry.LicenseProgramId=extDisEntryView.LicenseProgramId WHERE extDisEntryView.Id= @protocolId AND extDisEntryView.StudyLevelGroupId=@StudyLevelGroupId", slDel);

                switch (basisId)
                {
                    case "1":
                        basis = "обучение за счет средств федерального бюджета";
                        break;
                    case "2":
                        basis = string.Format("по договорам оказания государственной услуги по обучению по основной{0} образовательной программе высшего профессионального образования", sec);
                        break;
                }

                switch (formId)
                {
                    case "1":
                        form = "очная форма обучения";
                        form2 = "по очной форме";
                        break;
                    case "2":
                        form = "очно-заочная (вечерняя) форма обучения";
                        form2 = "по очно-заочной (вечерней) форме";
                        break;
                }

                string bakspec = "", naprspec = "", naprspecRod = "", profspec = "";

                if (MainClass.dbType == PriemType.PriemAspirant)
                {
                    bakspec = "магистра";
                    naprspec = "направление";
                    naprspecRod = "направлению";
                    profspec = "магистерской программе";
                }
                else
                {
                    if (LicenseProgramCode.EndsWith("00"))
                        bakspec = "бакалавра";
                    else
                        bakspec = "подготовки специалиста";

                    naprspec = "направление";
                    naprspecRod = "направлению";
                    profspec = "профилю";
                }
                wd.SetFields("Граждан", isRus ? "граждан РФ" : "иностранных граждан");
                wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                wd.SetFields("Стипендия", (basisId == "2" || formId == "2") ? "" : "\r\nи назначении стипендии");
                wd.SetFields("Стипендия2", (basisId == "2" || formId == "2") ? "" : " и назначении стипендии");
                wd.SetFields("Факультет", facDat);
                wd.SetFields("Форма", form);
                wd.SetFields("Основа", basis);
                wd.SetFields("БакСпец", bakspec);
                wd.SetFields("НапрСпец", string.Format(" {0} {1} «{2}»", naprspecRod, LicenseProgramCode, LicenseProgramName));
                wd.SetFields("ПриказОт", docDate);
                wd.SetFields("ПриказНомер", docNum);
                wd.SetFields("ПриказОт2", docDate);
                wd.SetFields("ПриказНомер2", docNum);
                wd.SetFields("Сокращ", sec);

                int curRow = 4;
                //int counter = 0;
                //string curSpez = "-";
                //string curHeader = "-";
                //string curCountry = "-";

                foreach (DataRow r in ds.Tables[0].Rows)
                {
                    /*
                    ++counter;
                    string spez = r["specialization"].ToString();
                    if (spez != curSpez)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("{3}\tпо {0} {1} \"{2}\"", naprspecRod, professionCode, profession, curSpez == "-" ? "" : "\r\n");

                        if (!string.IsNullOrEmpty(spez))
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("\tпо {0} \"{1}\"", profspec, spez);
                        }
                        
                        curSpez = spez;
                    }

                    if (!isRus)
                    {
                        string country = r["NameRod"].ToString();
                        if (country != curCountry)
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("\r\n граждан {0}:", country);

                            curCountry = country;
                        }
                    }

                    string header = r["EntryHeaderName"].ToString();
                    if (header != curHeader)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("\r\n\t{0}:", header);

                        curHeader = header;
                    }
                    */
                    td.AddRow(1);
                    curRow++;
                    td[0, curRow] = string.Format("\t\tп. № {0} {1} - исключить.", r["ФИО"].ToString(), r["TotalSum"]);

                }
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

        public static void PrintDisEntryView(string protocolId)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\DisEntryView.dot", MainClass.dirTemplates));
                TableDoc td = wd.Tables[0];

                string query = @"SELECT ed.extAbitAspirant.Id as Id, ed.extAbitAspirant.RegNum as Рег_Номер, 
                    ed.extPersonAspirant.PersonNum as 'Ид. номер', ed.extAbitMarksSum.TotalSum, 
                    ed.extPersonAspirant.FIO  as ФИО, 
                    ed.extAbitAspirant.LicenseProgramName, ed.extAbitAspirant.ProfileName as Specialization, ed.Country.NameRod 
                     FROM ed.extAbitAspirant 
                     INNER JOIN ed.extDisEntryView ON ed.extDisEntryView.AbiturientId=ed.extAbitAspirant.Id 
                     INNER JOIN ed.extPersonAspirant ON ed.extPersonAspirant.Id = ed.extAbitAspirant.PersonId 
                     INNER JOIN ed.Country ON ed.extPersonAspirant.NationalityId = ed.Country.Id
                     INNER JOIN ed.Competition ON ed.Competition.Id = ed.extAbitAspirant.CompetitionId 
                     LEFT JOIN ed.extAbitMarksSum ON ed.extAbitAspirant.Id = ed.extAbitMarksSum.Id";

                string where = " WHERE extDisEntryView.Id = @protocolId AND extDisEntryView.StudyLevelGroupId=@StudyLevelGroupId";
                string orderby = " ORDER BY extAbitAspirant.ProfileName, NameRod, ФИО ";

                DateTime protocolDate = (DateTime)MainClass.Bdc.GetValue(string.Format("SELECT Protocol.Date FROM ed.Protocol WHERE Protocol.Id='{0}'", protocolId));
                string protocolNum = MainClass.Bdc.GetStringValue(string.Format("SELECT Protocol.Number FROM ed.Protocol WHERE Protocol.Id='{0}'", protocolId));

                SortedList<string, object> slDel = new SortedList<string, object>();

                slDel.Add("@protocolId", protocolId);
                slDel.Add("@StudyLevelGroupId", MainClass.studyLevelGroupId);

                DataSet ds = MainClass.Bdc.GetDataSet(query + where + orderby, slDel);

                bool isRus = "1" == MainClass.Bdc.GetStringValue(" SELECT NationalityId FROM ed.extPersonAspirant INNER JOIN ed.Abiturient on Abiturient.personid=person.id INNER JOIN ed.extDisEntryView ON extDisEntryView.AbiturientId=Abiturient.Id WHERE extDisEntryView.Id=@protocolId", slDel);

                string entryProtocolId = MainClass.Bdc.GetStringValue("SELECT extProtocol.Id FROM ed.extDisEntryView INNER JOIN ed.extProtocol ON ed.extDisEntryView.AbiturientId=extProtocol.AbiturientId WHERE extProtocol.isOld = 0 and extDisEntryView.Id=@protocolId AND extProtocol.ProtocolTypeId=4 ", slDel);

                string docNum = "НОМЕР";
                string docDate = "ДАТА";
                DateTime tempDate;
                if (isRus)
                {
                    docNum = MainClass.Bdc.GetStringValue(string.Format("SELECT OrderNum FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId));
                    DateTime.TryParse(MainClass.Bdc.GetStringValue(string.Format("SELECT OrderDate FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId)), out tempDate);

                    docDate = tempDate.ToShortDateString();
                }
                else
                {
                    docNum = MainClass.Bdc.GetStringValue(string.Format("SELECT OrderNumFor FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId));

                    DateTime.TryParse(MainClass.Bdc.GetStringValue(string.Format("SELECT OrderDateFor FROM ed.OrderNUmbers WHERE ProtocolId='{0}'", entryProtocolId)), out tempDate);

                    docDate = tempDate.ToShortDateString();
                }

                string formId = MainClass.Bdc.GetStringValue("SELECT StudyForm.Id FROM ed.Protocol INNER JOIN ed.StudyForm ON Protocol.StudyFormId=StudyForm.Id WHERE Protocol.Id= @protocolId", slDel);
                string facDat = MainClass.Bdc.GetStringValue("SELECT SP_Faculty.DatName FROM ed.Protocol INNER JOIN ed.SP_Faculty ON Protocol.FacultyId=SP_Faculty.Id WHERE Protocol.Id= @protocolId", slDel);

                string basisId = MainClass.Bdc.GetStringValue("SELECT StudyBasis.Id FROM ed.Protocol INNER JOIN ed.StudyBasis ON Protocol.StudyBasisId=StudyBasis.Id WHERE Protocol.Id= @protocolId", slDel);
                string basis = string.Empty;
                string form = string.Empty;
                string form2 = string.Empty;

                bool? isSec = (bool?)MainClass.Bdc.GetValue(string.Format("SELECT IsSecond FROM ed.Protocol WHERE Protocol.Id= '{0}'", protocolId));
                bool? isList = (bool?)MainClass.Bdc.GetValue(string.Format("SELECT IsListener FROM ed.Protocol WHERE Protocol.Id='{0}'", protocolId));
                bool? isReduced = (bool?)MainClass.Bdc.GetValue(string.Format("SELECT IsReduced FROM ed.Protocol WHERE Protocol.Id='{0}'", protocolId));

                string list = string.Empty, sec = string.Empty;

                if (isList.HasValue && isList.Value)
                    list = " в качестве слушателя";

                if (isSec.HasValue && isSec.Value)
                    sec = " (для лиц с ВО)";

                if (isReduced.HasValue && isReduced.Value)
                    sec = " (сокращенной)";


                string LicenseProgramName = MainClass.Bdc.GetStringValue("SELECT TOP 1 qEntry.LicenseProgramName FROM ed.qEntry INNER JOIN ed.extDisEntryView ON qEntry.LicenseProgramId=extDisEntryView.LicenseProgramId WHERE extDisEntryView.Id= @protocolId AND extDisEntryView.StudyLevelGroupId=@StudyLevelGroupId", slDel);
                string LicenseProgramCode = MainClass.Bdc.GetStringValue("SELECT TOP 1 qEntry.LicenseProgramCode FROM ed.qEntry INNER JOIN ed.extDisEntryView ON qEntry.LicenseProgramId=extDisEntryView.LicenseProgramId WHERE extDisEntryView.Id= @protocolId AND extDisEntryView.StudyLevelGroupId=@StudyLevelGroupId", slDel);
                
                switch (basisId)
                {
                    case "1":
                        basis = "обучение за счет средств федерального бюджета";
                        break;
                    case "2":
                        basis = string.Format("по договорам оказания государственной услуги по обучению по основной{0} образовательной программе высшего профессионального образования", sec);
                        break;
                }

                switch (formId)
                {
                    case "1":
                        form = "очная форма обучения";
                        form2 = "по очной форме";
                        break;
                    case "2":
                        form = "очно-заочная (вечерняя) форма обучения";
                        form2 = "по очно-заочной (вечерней) форме";
                        break;
                }

                string bakspec = "", naprspec = "", naprspecRod = "", profspec = "";

                if (MainClass.dbType == PriemType.PriemAspirant)
                {
                    bakspec = "магистра";
                    naprspec = "направление";
                    naprspecRod = "направлению";
                    profspec = "магистерской программе";
                }
                else
                {
                    if (LicenseProgramCode.EndsWith("00"))
                        bakspec = "бакалавра";
                    else
                        bakspec = "подготовки специалиста";

                    naprspec = "направление";
                    naprspecRod = "направлению";
                    profspec = "профилю";

                }
                wd.SetFields("Граждан", isRus ? "граждан РФ" : "иностранных граждан");
                wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                wd.SetFields("Стипендия", basisId == "2" ? "" : "и назначении стипендии");
                wd.SetFields("Стипендия2", basisId == "2" ? "" : "и назначении стипендии");
                wd.SetFields("Факультет", facDat);
                wd.SetFields("Форма", form);
                wd.SetFields("Основа", basis);
                wd.SetFields("БакСпец", bakspec);
                wd.SetFields("НапрСпец", string.Format(" {0} {1} «{2}»", naprspecRod, LicenseProgramCode, LicenseProgramName));
                wd.SetFields("ПриказОт", docDate);
                wd.SetFields("ПриказНомер", docNum);
                wd.SetFields("ПриказОт2", docDate);
                wd.SetFields("ПриказНомер2", docNum);
                wd.SetFields("ПредставлениеОт", protocolDate.ToShortDateString());
                wd.SetFields("ПредставлениеНомер", protocolNum);
                wd.SetFields("Сокращ", sec);


                int curRow = 4;
                //int counter = 0;
                //string curSpez = "-";
                //string curHeader = "-";
                //string curCountry = "-";

                foreach (DataRow r in ds.Tables[0].Rows)
                {
                    /*
                    ++counter;
                    string spez = r["specialization"].ToString();
                    if (spez != curSpez)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("{3}\tпо {0} {1} \"{2}\"", naprspecRod, professionCode, profession, curSpez == "-" ? "" : "\r\n");

                        if (!string.IsNullOrEmpty(spez))
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("\tпо {0} \"{1}\"", profspec, spez);
                        }
                        
                        curSpez = spez;
                    }

                    if (!isRus)
                    {
                        string country = r["NameRod"].ToString();
                        if (country != curCountry)
                        {
                            td.AddRow(1);
                            curRow++;
                            td[0, curRow] = string.Format("\r\n граждан {0}:", country);

                            curCountry = country;
                        }
                    }

                    string header = r["EntryHeaderName"].ToString();
                    if (header != curHeader)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("\r\n\t{0}:", header);

                        curHeader = header;
                    }
                    */
                    td.AddRow(1);
                    curRow++;
                    td[0, curRow] = string.Format("\t\tп. № {0}, {1} - исключить.", r["ФИО"].ToString(), r["TotalSum"]);

                }
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
        public static string[] GetSplittedStrings(string sourceStr, int firstStrLen, int strLen, int numOfStrings)
        {
            sourceStr = sourceStr ?? "";
            string[] retStr = new string[numOfStrings];
            int index = 0, startindex = 0;
            for (int i = 0; i < numOfStrings; i++)
            {
                if (sourceStr.Length > startindex && startindex >= 0)
                {
                    int rowLength = firstStrLen;//длина первой строки
                    if (i > 1) //длина остальных строк одинакова
                        rowLength = strLen;
                    index = startindex + rowLength;
                    if (index < sourceStr.Length)
                    {
                        index = sourceStr.IndexOf(" ", index);
                        string val = index > 0 ? sourceStr.Substring(startindex, index - startindex) : sourceStr.Substring(startindex);
                        retStr[i] = val;
                    }
                    else
                        retStr[i] = sourceStr.Substring(startindex);
                }
                startindex = index;
            }

            return retStr;
        }
    }
}
