using EducServLib;
using PriemLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Priem
{
    public class EssayImportClass
    {
        public static void ImportEssay(string folder)
        {
            string query = @"SELECT DISTINCT FILES.[Id]
      ,[FileName]
	  ,FILES.FileTypeId
	  ,ENT.ObrazProgramName
	  ,ENT.ObrazProgramCrypt
  FROM [OnlinePriem2015].[dbo].qAbitFiles_OnlyEssayMotivLetter FILES
  INNER JOIN Application_LOG APP ON APP.Id = FILES.ApplicationId
  INNER JOIN [Entry] ENT ON ENT.Id = APP.EntryId
  WHERE ENT.CampaignYear = 2016
	AND ENT.StudyLevelGroupId = 4
	AND FILES.FileTypeId IN (2, 3)
	AND ENT.ObrazProgramName IN
	(
		'Информационные технологии и численные методы',
		'Физика',
		'Математическая физика'
	)
UNION
SELECT DISTINCT FILES.[Id]
      ,[FileName]
	  ,FILES.FileTypeId
	  ,ENT.ObrazProgramName
	  ,ENT.ObrazProgramCrypt
  FROM [OnlinePriem2015].[dbo].qAbitFiles_OnlyEssayMotivLetter FILES
  INNER JOIN Application_LOG APP ON APP.CommitId = FILES.CommitId
  INNER JOIN [Entry] ENT ON ENT.Id = APP.EntryId
  WHERE ENT.CampaignYear = 2016
	AND ENT.StudyLevelGroupId = 4
	AND FILES.FileTypeId IN (2, 3)
	AND ENT.ObrazProgramName IN
	(
		'Информационные технологии и численные методы',
		'Физика',
		'Математическая физика'
	)
UNION
SELECT DISTINCT FILES.[Id]
      ,[FileName]
	  ,FILES.FileTypeId
	  ,ENT.ObrazProgramName
	  ,ENT.ObrazProgramCrypt
  FROM [OnlinePriem2015].[dbo].qAbitFiles_OnlyEssayMotivLetter FILES
  INNER JOIN Application_LOG APP ON APP.PersonId = FILES.PersonId
  INNER JOIN [Entry] ENT ON ENT.Id = APP.EntryId
  WHERE ENT.CampaignYear = 2016
	AND ENT.StudyLevelGroupId = 4
	AND FILES.FileTypeId IN (2, 3)
	AND ENT.ObrazProgramName IN
	(
		'Информационные технологии и численные методы',
		'Физика',
		'Математическая физика'
	)";
            ProgressForm pf = new ProgressForm();

            try
            {
                pf.Show();
                pf.SetProgressText("Получение данных...");
                DataTable tbl = MainClass.BdcOnlineReadWrite.GetDataSet(query).Tables[0];
                pf.MaxPrBarValue = tbl.Rows.Count;
                pf.SetProgressText("Запись на диск...");
                foreach (DataRow rw in tbl.Rows)
                {
                    Guid FileId = rw.Field<Guid>("Id");
                    string FileName = rw.Field<string>("FileName");
                    int FileTypeId = rw.Field<int>("FileTypeId");
                    string ObrazProgramName = rw.Field<string>("ObrazProgramName");
                    string ObrazProgramCrypt = rw.Field<string>("ObrazProgramCrypt");

                    string sPath = Path.Combine(folder, ((ObrazProgramCrypt + " ") ?? "") + ObrazProgramName);
                    if (!Directory.Exists(sPath))
                        Directory.CreateDirectory(sPath);

                    FileInfo fi = new FileInfo(FileName);
                    string sFilePrefix = FileTypeId == 3 ? "Essay" : "Motivate";

                    string sFileName = Path.Combine(sPath, sFilePrefix + "_" + FileId.ToString() + fi.Extension);
                    if (!File.Exists(sFileName))
                    {
                        string q = "SELECT FileData FROM FileStorage WHERE Id=@Id";
                        byte[] bin = (byte[])MainClass.BdcOnlineReadWrite.GetValue(q, new SortedList<string, object>() { { "@Id", FileId } });
                        if (bin != null && bin.Length > 0)
                        {
                            File.WriteAllBytes(sFileName, bin);
                        }
                    }

                    pf.PerformStep();
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error(ex);
            }
            finally
            {
                pf.Close();
                MessageBox.Show("OK");
            }
        }
    }
}
