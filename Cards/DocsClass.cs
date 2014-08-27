using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Data;

using EducServLib;
using BDClassLib;
using PriemLib;

namespace Priem
{
    public class DocsClass
    {
        private DBPriem _bdcInet;
        private string _personId;
        private string _abitId;
        private string _commitId;

        public DocsClass(int personBarcode, int? abitCommitBarcode)
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

            _personId = _bdcInet.GetStringValue("SELECT Person.Id FROM Person WHERE Person.Barcode = " + personBarcode);

            if (abitCommitBarcode == null)
                _abitId = null;
            else
            {
                _abitId = _bdcInet.GetStringValue("SELECT qAbiturient.Id FROM qAbiturient WHERE qAbiturient.CommitNumber = " + abitCommitBarcode);
                _commitId = _bdcInet.GetStringValue("SELECT qAbiturient.CommitId FROM qAbiturient WHERE qAbiturient.CommitNumber = " + abitCommitBarcode);
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

        public void OpenFile(List<KeyValuePair<string, string>> lstFiles)
        {
            try
            {
                foreach (KeyValuePair<string, string> file in lstFiles)
                {
                    byte[] bt = _bdcInet.ReadFile(string.Format("SELECT FileData FROM extAbitFiles WHERE Id = '{0}'", file.Key));

                    string filename = file.Value.Replace(@"\", "-").Replace(@":", "-");

                    StreamWriter sw = new StreamWriter(MainClass.saveTempFolder + filename);
                    BinaryWriter bw = new BinaryWriter(sw.BaseStream);
                    bw.Write(bt);
                    bw.Flush();
                    bw.Close();
                    Process.Start(MainClass.saveTempFolder + filename);
                }
            }
            catch (System.Exception exc)
            {
                WinFormsServ.Error("Ошибка открытия файла: " + exc.Message);
            }
        }

        public List<KeyValuePair<string, string>> UpdateFiles()
        {
            try
            {
                if (_personId == null)
                    return null;

                List<KeyValuePair<string, string>> lstFiles = new List<KeyValuePair<string, string>>();

                string query = string.Format("SELECT Id, FileName + ' (' + convert(nvarchar, extAbitFiles.LoadDate, 104) + ' ' + convert(nvarchar, extAbitFiles.LoadDate, 108) + ')' + FileExtention AS FileName FROM extAbitFiles WHERE extAbitFiles.PersonId = '{0}' {1} {2}", _personId,
                    !string.IsNullOrEmpty(_abitId) ? " AND (extAbitFiles.ApplicationId = '" + _abitId + "' OR extAbitFiles.ApplicationId IS NULL)" : "",
                    !string.IsNullOrEmpty(_commitId) ? " AND (extAbitFiles.CommitId = '" + _commitId + "' OR extAbitFiles.CommitId IS NULL)" : "");
                DataSet ds = _bdcInet.GetDataSet(query + " ORDER BY extAbitFiles.LoadDate DESC");
                foreach (DataRow dRow in ds.Tables[0].Rows)
                {
                    lstFiles.Add(new KeyValuePair<string, string>(dRow["Id"].ToString(), dRow["FileName"].ToString()));
                }

                return lstFiles;
            }
            catch (System.Exception exc)
            {
                WinFormsServ.Error("Ошибка обновления данных о приложениях: " + exc.Message);
                return null;
            }
        }
    }
}
