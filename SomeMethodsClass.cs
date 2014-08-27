using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.Objects;
using System.Data;
using System.Transactions;

using BDClassLib;
using EducServLib;
using PriemLib;

namespace Priem
{
    public static class SomeMethodsClass
    {
        public static void FillOlymps()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "CSV|*.csv";
            ofd.Multiselect = true;

            if (!(ofd.ShowDialog() == DialogResult.OK))
                return;
            foreach (string fileName in ofd.FileNames)
            {
                using (StreamReader sr = new StreamReader(fileName, Encoding.GetEncoding(1251)))
                {
                    List<string[]> list = new List<string[]>();

                    while (!sr.EndOfStream)
                    {
                        string str = sr.ReadLine();
                        char[] split = new char[] { ';' };
                        string[] values = str.Split(split, 4);

                        list.Add(values);
                    }

                    int i = 0;
                    foreach (string[] val in list)
                    {
                        i++;

                        if (val.Length < 4)
                        {
                            MessageBox.Show("Строка " + i.ToString() + ": некорректный формат строки!");
                            continue;
                        }

                        string num = val[0].Trim();
                        int j;
                        int? numInt;
                        if(int.TryParse(num, out j))
                            numInt = j;
                        else
                            numInt = null;
                       
                        string name = val[1].Trim();
                        string olName = num + " - " + name;                       

                        using (PriemEntities context = new PriemEntities())
                        {
                            int? olNameId;
                            var ol = from qq in context.OlympName
                                       where qq.Name == olName
                                       select qq;

                            if (ol != null && ol.Count() > 0)
                                olNameId = (ol.First()).Id;
                            else
                            {
                                ObjectParameter olEnt = new ObjectParameter("id", typeof(Int32));
                                context.OlympName_Insert(olName, numInt, olEnt);
                                olNameId = (int)olEnt.Value;
                            }

                            string subjName = val[2].Trim();
                            
                            int? subjId;
                            var subj = from qq in context.OlympSubject
                                       where qq.Name == subjName
                                       select qq;

                            if (subj != null && subj.Count() > 0)
                                subjId = (subj.First()).Id;
                            else
                            {
                                ObjectParameter sbEnt = new ObjectParameter("id", typeof(Int32));
                                context.OlympSubject_Insert(subjName, "", sbEnt);
                                subjId = (int)sbEnt.Value;
                            } 

                            int? levId;
                            int level;
                            if (!int.TryParse(val[3].Trim(), out level))
                                levId = null;
                            else
                                levId = level;

                            int cnt = (from ob in context.OlympBook
                                       where ob.OlympTypeId == 4 && ob.OlympNameId == olNameId && ob.OlympSubjectId == subjId && ob.OlympLevelId == levId
                                       select ob).Count();

                            if (cnt > 0)
                                continue;
                            
                            ObjectParameter EntId = new ObjectParameter("id", typeof(Int32));
                            context.OlympBook_Insert(4, olNameId, subjId, levId, EntId);                        
                        }
                    }
                }
            }

            MessageBox.Show("Выполнено!");
        }

        // проверка на уникальность абитуриента
        public static bool CheckThreeAbits(PriemEntities context, Guid? personId, int? LicenseProgramId)
        {
            // если прием - то проверяем на три заявления
            if (MainClass.dbType != PriemType.Priem)
                return true;  

            //просто сосчитаем количество созданных конкурсов на человека
            var concurses = (from allab in context.qAbitAll
                             where allab.PersonId == personId
                                 //&& allab.LicenseProgramId != LicenseProgramId
                                 //&& allab.ObrazProgramId != ObrazProgramId
                                 //&& (ProfileId == null ? allab.ProfileId != null : allab.ProfileId != ProfileId)
                             && allab.StudyLevelGroupId == MainClass.studyLevelGroupId
                             && allab.BackDoc != true
                             select new { allab.LicenseProgramId, allab.ObrazProgramId, allab.ProfileId }).Distinct();
            int cntExist = concurses.Count();

            return true;
        }

        // проверка на уникальность абитуриента
        private static bool CheckIdent(extPersonAspirant person)
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

                context.CheckPersonIdent(person.Surname, person.Name, person.SecondName, person.BirthDate, person.PassportSeries, person.PassportNumber,
                    person.AttestatRegion, person.AttestatSeries, person.AttestatNum, boolPar);
               
                return Convert.ToBoolean(boolPar.Value);
            }
        }

        public static void ImportMagAbits()
        {
            List<string> lstPersons = new List<string>();
            List<string> lstAbits = new List<string>();

            LoadFromInet loadClass = new LoadFromInet();
            DBPriem _bdcInet = loadClass.BDCInet;

            using (PriemEntities context = new PriemEntities())
            {
                string _sQuery = @"SELECT DISTINCT qAbiturient.Id, qAbiturient.PersonId, qAbiturient.Barcode AS AbitBarcode, extPersonAspirant.Barcode AS PersonBarcode
                              FROM qAbiturient INNER JOIN extPersonAspirant ON qAbiturient.PersonId = extPersonAspirant.Id WHERE StudyLevelId = 17 AND Enabled = 1 AND IsApprovedByComission = 1 AND IsImported = 0";
                DataSet ds = _bdcInet.GetDataSet(_sQuery);

                Guid? personId;
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    int? abitBarcode = (int?)dr["AbitBarcode"];
                    if (!MainClass.CheckAbitBarcode(abitBarcode))
                        continue;
                    
                    
                    int? persBarcode = (int?)dr["PersonBarcode"]; 
                    if (!MainClass.CheckPersonBarcode(persBarcode))
                    {
                        personId = (from pers in context.extPersonAspirant
                                    where pers.Barcode == persBarcode
                                    select pers.Id).FirstOrDefault();
                    }
                    else
                    {
                        extPersonAspirant person = loadClass.GetPersonByBarcode(persBarcode.Value);

                        if (!CheckIdent(person))
                        {
                            lstPersons.Add(persBarcode.ToString());
                            continue;
                        }

                        ObjectParameter entId = new ObjectParameter("id", typeof(Guid));
                        context.Person_insert(person.Barcode, person.Name, person.SecondName, person.Surname, person.BirthDate, person.BirthPlace, person.PassportTypeId,
                            person.PassportSeries, person.PassportNumber, person.PassportAuthor, person.PassportDate, person.Sex, person.CountryId, person.NationalityId, 
                            person.RegionId, person.Phone, person.Mobiles, person.Email, person.Code, person.City, person.Street, person.House, person.Korpus, 
                            person.Flat, person.CodeReal, person.CityReal, person.StreetReal, person.HouseReal, person.KorpusReal, person.FlatReal,  person.KladrCode,
                            person.HostelAbit, person.HostelEduc, false, null, false, null, person.IsExcellent, 
                            person.LanguageId, person.SchoolCity, person.SchoolTypeId, person.SchoolName, person.SchoolNum, person.SchoolExitYear, person.SchoolAVG, 
                            person.CountryEducId, person.RegionEducId, person.IsEqual, person.AttestatRegion, person.AttestatSeries, person.AttestatNum, person.DiplomSeries, person.DiplomNum, 
                            person.HighEducation, person.HEProfession, person.HEQualification, person.HEEntryYear, person.HEExitYear, person.HEStudyFormId, 
                            person.HEWork, person.Stag, person.WorkPlace, person.MSVuz, person.MSCourse, person.MSStudyFormId, person.Privileges, person.PassportCode,
                            person.PersonalCode, person.PersonInfo, person.ExtraInfo, person.ScienceWork, person.StartEnglish, person.EnglishMark, false, person.SNILS, entId);

                        //_bdcInet.ExecuteQuery("UPDATE Person SET IsImported = 1 WHERE Person.Barcode = " + persBarcode);     
                        
                        personId = (Guid)entId.Value;
                    }

                    qAbiturient abit = loadClass.GetAbitByBarcode(abitBarcode.Value);

                    int cnt = (from en in context.qEntry
                               where en.Id == abit.EntryId && !en.IsClosed
                               select en).Count();

                    if(cnt == 0)
                    {
                        lstAbits.Add(abitBarcode.ToString());
                        continue;
                    }

                    ObjectParameter abEntId = new ObjectParameter("id", typeof(Guid));

                    int competitionId;
                    if (abit.StudyBasisId == 1)
                        competitionId = 4;
                    else
                        competitionId = 3;

                    context.Abiturient_Insert(personId, abit.EntryId, competitionId, false, false, false, false, null, abit.DocDate, DateTime.Now,
                    false, false, null, null, null, null, abit.LanguageId, false,
                    abit.Priority, abit.Barcode, abit.CommitId, abit.CommitNumber, abit.IsGosLine, abEntId);

                    // _bdcInet.ExecuteQuery("UPDATE Application SET IsImported = 1 WHERE Application.Barcode = " + abitBarcode);
                }
            }        
        }

        public static void SetBackDocForBudgetInEntryView()
        {
            if (!MainClass.IsPasha())
                return;

            using (PriemEntities context = new PriemEntities())
            {
                if (MessageBox.Show("Проставить 'Забрал документы' для платных заявлений, поступивших на бесплатное?", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {                    
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew, TimeSpan.FromHours(1)))
                    {
                        var abits = from ev in context.extEntryView
                                    join ab in context.extAbitAspirant
                                    on ev.AbiturientId equals ab.Id
                                    where ab.StudyLevelGroupId == 1 && ab.StudyBasisId == 1 && !ab.BackDoc && ab.HasOriginals
                                    select ab;

                        foreach (extAbitAspirant abit in abits)
                        {
                            var abBackDocks = from ab in context.extAbitAspirant
                                              where ab.StudyLevelGroupId == abit.StudyLevelGroupId
                                              && ab.IsReduced == abit.IsReduced && ab.IsParallel == abit.IsParallel && ab.IsSecond == abit.IsSecond
                                              && ab.FacultyId == abit.FacultyId && ab.LicenseProgramId == abit.LicenseProgramId
                                              && ab.ObrazProgramId == abit.ObrazProgramId
                                              && (abit.ProfileId == null ? ab.ProfileId == null : ab.ProfileId == abit.ProfileId)
                                              && ab.StudyFormId == abit.StudyFormId
                                              && ab.StudyBasisId == 2
                                              select ab;

                            if (abBackDocks.Count() > 0)
                            {
                                foreach(extAbitAspirant abBack in abBackDocks)
                                {
                                    context.Abiturient_UpdateBackDoc(true, DateTime.Now, abBack.Id);
                                }
                            }
                        }


                        transaction.Complete();
                    }
                }
            }
        }

        public static void DeleteDogFromFirstWave()
        {
            if (!MainClass.IsPasha())
                return;

            using (PriemEntities context = new PriemEntities())
            {
                if (MessageBox.Show("Удалить платников, забравших документы или зачисленных на 1 курс, из FirstWave?", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {                    
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew, TimeSpan.FromHours(1)))
                    {
                        context.FirstWave_DELETE_DogEntryBack();

                        transaction.Complete();
                    }
                }
            }
        }
    }
}
