using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

using BDClassLib;
using EducServLib;
using PriemLib;

namespace Priem
{
    public class LoadFromInet
    {
        private DBPriem _bdcInet;

        public LoadFromInet()
        {
            _bdcInet = new DBPriem();
            try
            {
                _bdcInet.OpenDatabase(MainClass.connStringOnline);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);                
            }
        }

        public DBPriem BDCInet
        {
            get { return _bdcInet; }
        }

        public void CloseDB()
        {
            _bdcInet.CloseDataBase();
        }          
        
        public void UpdatePersonData(int personBarc)
        {
            try
            {
                _bdcInet.ExecuteQuery("UPDATE Person SET IsImported = 1 WHERE Person.Barcode = " + personBarc);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public DataTable GetPersonEgeByBarcode(int fileNum)
        {
            string queryEge = "SELECT EgeMark.Id, EgeMark.EgeExamId AS ExamId, EgeMark.Value, EgeCertificate.Number, EgeMark.EgeCertificateId FROM EgeMark LEFT JOIN EgeCertificate ON EgeMark.EgeCertificateId = EgeCertificate.Id LEFT JOIN Person ON EgeCertificate.PersonId = Person.Id";
            DataSet dsEge = _bdcInet.GetDataSet(queryEge + " WHERE Person.Barcode = " + fileNum + " ORDER BY EgeMark.EgeCertificateId ");
            return dsEge.Tables[0];
        }

        public extPersonAspirant GetPersonByBarcode(int fileNum)
        {
            try
            {
                string personQueryInet =
                @"SELECT Id, Barcode, Name, SecondName, Surname, BirthDate, BirthPlace, Sex,
                    PassportTypeId, PassportSeries, PassportNumber, PassportAuthor, PassportDate,
                    PassportCode, '' AS PersonalCode, CountryId, NationalityId, RegionId, Phone, Mobiles, Email,
                    Code, City, Street, House, Korpus, Flat, CodeReal, CityReal, StreetReal, HouseReal, KorpusReal, FlatReal,
                    HostelAbit, HostelEduc, LanguageId, SchoolCity, SchoolTypeId, SchoolName, SchoolNum, SchoolExitYear, IsExcellent,
                    CountryEducId, RegionEducId, AttestatRegion, AttestatSeries, AttestatNum,
                    EducationDocumentSeries AS DiplomSeries, EducationDocumentNumber AS DiplomNum, AvgMark AS SchoolAVG,
                    (case when SchoolTypeId=1 then '' else SchoolName end) AS HighEducation, HEProfession AS HEProfession, 
                    HEQualification AS HEQualification, DiplomaTheme AS HEWork, HEEntryYear, HEExitYear, StudyFormId AS HEStudyFormId, 
                    Parents AS PersonInfo, AddInfo AS ExtraInfo, StartEnglish, EnglishMark, AbiturientTypeId, KladrCode, SNILS
                    FROM extPerson_All
                    WHERE 0=0";

                DataSet ds = _bdcInet.GetDataSet(personQueryInet + " AND extPerson_All.Barcode = " + fileNum);
                if (ds.Tables[0].Rows.Count == 0)
                    throw new Exception("Записей не найдено");

                DataRow row = ds.Tables[0].Rows[0];
                extPersonAspirant pers = new extPersonAspirant();
               
                int iAbitTypeId = (int)row["AbiturientTypeId"];

                //if (iAbitTypeId == 1 && MainClass.dbType == PriemType.PriemMag)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего на 1 курс бакалавриата/специалитета");
                //if (iAbitTypeId == 2 && MainClass.dbType == PriemType.Priem)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего в магистратуру");

                //if (iAbitTypeId == 1)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего на 1 курс. Проверьте штрих-код");
                //if (iAbitTypeId == 2)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего в магистратуру. Проверьте штрих-код");
                //if (iAbitTypeId == 3)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего в колледжи СПбГУ. Проверьте штрих-код");
                //    //WinFormsServ.Error("Выбран профиль человека, переводящегося в СПбГУ. Проверьте штрих-код");
                //if (iAbitTypeId == 5)
                //    WinFormsServ.Error("Выбран профиль человека, восстанавливающегося в СПбГУ. Проверьте штрих-код");
                //if (iAbitTypeId == 6 || iAbitTypeId == 7)
                //    WinFormsServ.Error("Выбран профиль человека, переводящегося внутри СПбГУ. Проверьте штрих-код");
                //if (iAbitTypeId == 8)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего в АГ СПбГУ. Проверьте штрих-код");

                pers.Id = (Guid)row["Id"];
                pers.Barcode = (int?)row["Barcode"];
                pers.FIO = Util.GetFIO(row["Surname"].ToString(), row["Name"].ToString(), row["SecondName"].ToString());
                pers.Name = row["Name"].ToString();
                pers.SecondName = row["SecondName"].ToString();
                pers.Surname = row["Surname"].ToString();
                pers.BirthDate = (DateTime?)row["BirthDate"] ?? DateTime.Now;
                pers.BirthPlace = row["BirthPlace"].ToString();
                pers.PassportTypeId = (int)(row["PassportTypeId"]);
                pers.PassportSeries = row["PassportSeries"].ToString();
                pers.PassportNumber = row["PassportNumber"].ToString();
                pers.PassportAuthor = row["PassportAuthor"].ToString();
                pers.PassportDate = QueryServ.ToNullDateTimeValue(row["PassportDate"]);
                pers.PassportCode = row["PassportCode"].ToString();
                pers.PersonalCode = row["PersonalCode"].ToString();
                pers.SNILS = row["SNILS"].ToString();
                pers.Sex = QueryServ.ToBoolValue(row["Sex"]);
                pers.CountryId = (int?)(Util.ToNullObject(row["CountryId"]));
                pers.NationalityId = (int)row["NationalityId"];
                pers.RegionId = (int?)(Util.ToNullObject(row["RegionId"]));
                pers.Phone = row["Phone"].ToString();
                pers.Mobiles = row["Mobiles"].ToString();
                pers.Email = row["Email"].ToString();
                pers.Code = row["Code"].ToString();
                pers.City = row["City"].ToString();
                pers.Street = row["Street"].ToString();
                pers.House = row["House"].ToString();
                pers.Korpus = row["Korpus"].ToString();
                pers.Flat = row["Flat"].ToString();
                pers.CodeReal = row["CodeReal"].ToString();
                pers.CityReal = row["CityReal"].ToString();
                pers.StreetReal = row["StreetReal"].ToString();
                pers.HouseReal = row["HouseReal"].ToString();
                pers.KorpusReal = row["KorpusReal"].ToString();
                pers.FlatReal = row["FlatReal"].ToString();
                pers.KladrCode = row["KladrCode"].ToString();
                pers.HostelEduc = QueryServ.ToBoolValue(row["HostelEduc"]);
                pers.HostelAbit = QueryServ.ToBoolValue(row["HostelAbit"]);
                pers.HasAssignToHostel = false;
                pers.HasExamPass = false;
                pers.IsExcellent = QueryServ.ToBoolValue(row["IsExcellent"]);
                pers.LanguageId = (int?)(Util.ToNullObject(row["LanguageId"]));
                pers.SchoolCity = row["SchoolCity"].ToString();
                pers.SchoolTypeId = (int?)(Util.ToNullObject(row["SchoolTypeId"]));
                pers.SchoolName = row["SchoolName"].ToString();
                pers.SchoolNum = row["SchoolNum"].ToString();
                int SchoolExitYear = 0;
                int.TryParse(row["SchoolExitYear"].ToString(), out SchoolExitYear);
                pers.SchoolExitYear = SchoolExitYear;
                pers.CountryEducId = (int?)(Util.ToNullObject(row["CountryEducId"]));
                pers.RegionEducId = (int?)(Util.ToNullObject(row["RegionEducId"]));
                pers.AttestatRegion = row["AttestatRegion"].ToString();
                pers.AttestatSeries = row["AttestatSeries"].ToString();
                pers.AttestatNum = row["AttestatNum"].ToString();
                pers.DiplomSeries = row["DiplomSeries"].ToString();
                pers.DiplomNum = row["DiplomNum"].ToString();
                
                double avg;                
                if(!double.TryParse(row["SchoolAVG"].ToString(), out avg))
                    pers.SchoolAVG = null;
                else
                    pers.SchoolAVG = avg;

                pers.HighEducation = row["HighEducation"].ToString();
                pers.HEProfession = row["HEProfession"].ToString();
                pers.HEQualification = row["HEQualification"].ToString();
                pers.HEEntryYear = (int?)(Util.ToNullObject(row["HEEntryYear"]));
                pers.HEExitYear = (int?)(Util.ToNullObject(row["HEExitYear"]));
                pers.HEStudyFormId = (int?)(Util.ToNullObject(row["HEStudyFormId"]));
                pers.HEWork = row["HEWork"].ToString();
                pers.PersonInfo = row["PersonInfo"].ToString();
                pers.ExtraInfo = row["ExtraInfo"].ToString();
                pers.StartEnglish = QueryServ.ToBoolValue(row["StartEnglish"]);
                int EnglishMark = 0;
                int.TryParse(row["EnglishMark"].ToString(), out EnglishMark);
                pers.EnglishMark = EnglishMark == 0 ? null : (int?)EnglishMark;

                DataSet dsWork = _bdcInet.GetDataSet(string.Format(@"
                      SELECT  PersonWork.WorkPlace + ', ' + PersonWork.WorkProfession + ', ' + PersonWork.WorkSpecifications + ' стаж: ' + PersonWork.Stage AS Work,
                      PersonWork.WorkPlace + ', ' + PersonWork.WorkProfession + ', ' + PersonWork.WorkSpecifications AS Place, PersonWork.Stage
                      FROM PersonWork WHERE PersonWork.PersonId = '{0}'", pers.Id));

                if (dsWork.Tables[0].Rows.Count == 0)
                {
                    pers.Stag = string.Empty;
                    pers.WorkPlace = string.Empty;
                }
                else if (dsWork.Tables[0].Rows.Count == 1)
                {
                    pers.Stag = dsWork.Tables[0].Rows[0]["Stage"].ToString();
                    pers.WorkPlace = dsWork.Tables[0].Rows[0]["Place"].ToString();
                }
                else
                {
                    string work = string.Empty;
                    foreach (DataRow dr in dsWork.Tables[0].Rows)
                    {
                        work += dr["Work"].ToString() + ";" + Environment.NewLine;
                    }
                    pers.WorkPlace = work;
                }

                DataSet dsScienceWork = _bdcInet.GetDataSet(string.Format(@"
                      SELECT ScienceWorkType.Name + ': ' + PersonScienceWork.WorkInfo AS ScienseWork                      
                      FROM PersonScienceWork LEFT JOIN ScienceWorkType ON PersonScienceWork.WorkTypeId = ScienceWorkType.Id WHERE PersonScienceWork.PersonId = '{0}'", pers.Id));

                if (dsScienceWork.Tables[0].Rows.Count == 0)
                    pers.ScienceWork = string.Empty;                
               
                else
                {
                    string work = string.Empty;
                    foreach (DataRow dr in dsScienceWork.Tables[0].Rows)
                    {
                        work += dr["ScienseWork"].ToString() + ";" + Environment.NewLine;
                    }
                    pers.ScienceWork = work;
                }  
               
                //pers.MSVuz = row["MSVuz"].ToString();
                //pers.MSCourse = row["MSCourse"].ToString();
                //pers.MSStudyFormId = (int?)row["MSStudyFormId"];  
              
                return pers;
            }
            catch
            {
                return null;
            }
        }

        public qAbiturient GetAbitByBarcode(int fileNum)
        {
            try
            {
                string abitQueryInet = @"SELECT qAbiturient.Barcode, EntryId, qAbiturient.HostelEduc, 
                            (Case when qAbiturient.Enabled = 1 then 0 else 1 end) AS BackDoc, 
                            qAbiturient.DateOfDisable AS BackDocDate, qAbiturient.DateOfStart AS DocDate,                           
                            qAbiturient.Priority, qAbiturient.LicenseProgramId, qAbiturient.ObrazProgramId, 
                            qAbiturient.ProfileId, qAbiturient.FacultyId, qAbiturient.StudyFormId, 
                            qAbiturient.StudyBasisId, qAbiturient.IsSecond
                            FROM qAbiturient WHERE 0=0";

                DataSet ds = _bdcInet.GetDataSet(abitQueryInet + " AND qAbiturient.Barcode = " + fileNum);
                if (ds.Tables[0].Rows.Count == 0)
                    throw new Exception("Записей не найдено");

                DataRow row = ds.Tables[0].Rows[0];
                qAbiturient abit = new qAbiturient();

                abit.IsSecond = (bool)row["IsSecond"];
                abit.EntryId = (Guid)row["EntryId"];
                abit.FacultyId = (int)row["FacultyId"];
                abit.LicenseProgramId = (int)row["LicenseProgramId"];
                abit.ObrazProgramId = (int)row["ObrazProgramId"];
                abit.ProfileId = (Guid?)(Util.ToNullObject(row["ProfileId"]));
                abit.StudyFormId = (int)row["StudyFormId"];
                abit.StudyBasisId = (int)row["StudyBasisId"];
               
                abit.BackDoc = QueryServ.ToBoolValue(row["BackDoc"]);
                abit.BackDocDate = row.Field<DateTime?>("BackDocDate");
                abit.DocDate = (DateTime)row["DocDate"];
         
                int prior;
                if (!int.TryParse(row["Priority"].ToString(), out prior))
                    abit.Priority = null;
                else
                    abit.Priority = prior;

                abit.Barcode = int.Parse(row["Barcode"].ToString());

                return abit;
            }
            catch
            {
                return null;
            }
        }    
    }
}
