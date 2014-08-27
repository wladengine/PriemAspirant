using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;

using EducServLib;
using BDClassLib;
using BaseFormsLib;
using PriemLib;

namespace Priem
{
    public partial class PersonInetList : BookList
    {
        private DBPriem bdcInet;
        private LoadFromInet loadClass;

        //конструктор
        public PersonInetList()
        {
            InitializeComponent();

            Dgv = dgvAbiturients;
            _tableName = "ed.extPersonAspirant";
            _title = "Список абитуриентов СПбГУ";

            InitControls();
        }        
        
        //дополнительная инициализация контролов
        protected override void  ExtraInit()
        {
            base.ExtraInit();            

            if (MainClass.RightsJustView())
            {
                btnLoad.Enabled = false;
                btnAdd.Enabled = false;
            }
           
            if (MainClass.dbType == PriemType.PriemAspirant)
                tbPersonNum.Visible = lblBarcode.Visible = btnLoad.Visible = false;

            //Dgv.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(Dgv_CellDoubleClick);
        }

        //поле поиска
        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbiturients, "FIO", tbSearch.Text);
        }

        protected override void OpenCard(string id, BaseFormEx formOwner, int? index)
        {
            MainClassCards.OpenCardPerson(MainClass.mainform, id, formOwner, index);
        }

        protected override void GetSource()
        {
            _sQuery = "SELECT DISTINCT extPersonAspirant.Id, extPersonAspirant.PersonNum, extPersonAspirant.FIO, extPersonAspirant.PassportData, extPersonAspirant.EducDocument FROM ed.extPersonAspirant WHERE SchoolTypeId=4 ";
            string join = "";

            if (!chbShowAll.Checked)
            {
                join = string.Format(" AND IsAspirantOnly=1 AND extPersonAspirant.Id NOT IN (SELECT Person.Id FROM ed.Person INNER JOIN ed.Abiturient ON Abiturient.PersonId=Person.Id INNER JOIN ed.qAbiturientForeignApplicationsOnly ON qAbiturientForeignApplicationsOnly.Id = Abiturient.Id)");
            }
            
            HelpClass.FillDataGrid(Dgv, _bdc, _sQuery + join, "", " ORDER BY FIO");
            SetVisibleColumnsAndNameColumns();    
        }

        protected override void SetVisibleColumnsAndNameColumns()
        {
            Dgv.AutoGenerateColumns = false;

            foreach (DataGridViewColumn col in Dgv.Columns)
            {
                col.Visible = false;
            }
            
            this.Width = 608;
            dgvAbiturients.Columns["PersonNum"].Width = 70;
            dgvAbiturients.Columns["FIO"].Width = 246;

            SetVisibleColumnsAndNameColumnsOrdered("PersonNum", "Ид_номер", 0);
            SetVisibleColumnsAndNameColumnsOrdered("FIO", "ФИО", 1);
            SetVisibleColumnsAndNameColumnsOrdered("PassportData", "Паспортные данные", 2);
            SetVisibleColumnsAndNameColumnsOrdered("EducDocument", "Документ об образовании", 3);          
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            loadClass = new LoadFromInet();
            bdcInet = loadClass.BDCInet;

            int fileNum = 0;

            string barcText = tbPersonNum.Text.Trim();

            if (barcText == string.Empty)
            {
                WinFormsServ.Error("Не введен номер");
                return;
            }

            if (barcText.Length == 7)
            {
                if (barcText.StartsWith("2") && MainClass.dbType == PriemType.Priem)
                {
                    WinFormsServ.Error("Выбран человек, подавший заявления в магистратуру");
                    return;
                }

                barcText = barcText.Substring(1);
            }

            if (!int.TryParse(barcText, out fileNum))
            {
                WinFormsServ.Error("Неправильно введен номер");
                return;
            }

            if (MainClass.CheckAbitBarcode(fileNum))
            {
                try
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        int cnt = (from ab in context.Abiturient
                                   where ab.Barcode == fileNum
                                   select ab).Count();

                        if (cnt > 0)
                        {
                            WinFormsServ.Error("Запись уже добавлена!");
                            return;
                        }

                        string query = "SELECT Id, Enabled FROM [Application] WHERE Barcode=@Barcode";
                        DataTable tbl = bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@Barcode", fileNum } }).Tables[0];

                        if (tbl.Rows.Count == 0)
                        {
                            WinFormsServ.Error("Запись не найдена!");
                            return;
                        }
                        if (!(tbl.Rows[0].Field<bool?>("Enabled") ?? true))
                        {
                            if (MessageBox.Show("Данное заявление уже отозвано в интернете. Вы всё равно хотите загрузить его?", "Внимание!", MessageBoxButtons.YesNo) 
                                == System.Windows.Forms.DialogResult.No)
                                return;
                        }

                        CardFromInet crd = new CardFromInet(null, fileNum, false);
                        crd.ToUpdateList += new UpdateListHandler(UpdateDataGrid);
                        crd.Show();
                    }

                    /*
                    //extPersonAspirant person = loadClass.GetPersonByBarcode(fileNum);                      
                    DataTable dtEge = new DataTable();
                                        
                    //if(person != null)
                    //{
                    //    string queryEge = "SELECT EgeMark.Id, EgeMark.EgeExamNameId AS ExamId, EgeMark.Value, EgeCertificate.PrintNumber, EgeCertificate.Number, EgeMark.EgeCertificateId FROM EgeMark LEFT JOIN EgeCertificate ON EgeMark.EgeCertificateId = EgeCertificate.Id LEFT JOIN Person ON EgeCertificate.PersonId = Person.Id";
                    //    DataSet dsEge = bdcInet.GetDataSet(queryEge + " WHERE Person.Barcode = " + fileNum + " ORDER BY EgeMark.EgeCertificateId ");
                    //    dtEge = dsEge.Tables[0];
                    //}

                    CardFromInet crd = new CardFromInet(fileNum, null, true);
                    crd.ToUpdateList += new UpdateListHandler(UpdateDataGrid);
                    crd.Show();
                }
                catch (Exception exc)
                {
                    WinFormsServ.Error(exc.Message);
                    tbPersonNum.Text = "";
                    tbPersonNum.Focus();
                }
            }
                    */
                }
                catch (Exception exc)
                {
                    WinFormsServ.Error(exc.Message);
                    tbPersonNum.Text = "";
                    tbPersonNum.Focus();
                }
            }
            else
            {
                UpdateDataGrid();
                using (PriemEntities context = new PriemEntities())
                {
                    extAbitAspirant abit = (from ab in context.extAbitAspirant
                                    where ab.Barcode == fileNum
                                    select ab).FirstOrDefault();

                    string fio = abit.FIO;
                    string num = abit.PersonNum;
                    string persId = abit.PersonId.ToString();

                    WinFormsServ.Search(this.dgvAbiturients, "PersonNum", num);
                    DialogResult dr = MessageBox.Show(string.Format("Абитуриент {0} с данным номером баркода уже импортирован в базу.\nОткрыть карточку абитуриента?", fio), "Внимание", MessageBoxButtons.YesNo);
                    if (dr == System.Windows.Forms.DialogResult.Yes)
                        MainClassCards.OpenCardPerson(MainClass.mainform, persId, this, null);
                }
            }



            tbPersonNum.Text = "";
            tbPersonNum.Focus();
            loadClass.CloseDB();  
        }

        private void PersonList_Load(object sender, EventArgs e)
        {
            tbPersonNum.Focus();
        }

        private void PersonList_Activated(object sender, EventArgs e)
        {
            tbPersonNum.Focus();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            UpdateDataGrid();            
        }

        private void tbNumber_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbiturients, "PersonNum", tbNumber.Text);
        }

        private void chbShowAll_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        protected override void btnRemove_Click(object sender, EventArgs e)
        {
            if (MainClass.IsPasha())
            {
                if (MessageBox.Show("Удалить записи?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    foreach (DataGridViewRow dgvr in Dgv.SelectedRows)
                    {
                        string itemId = dgvr.Cells["Id"].Value.ToString();
                        Guid g = Guid.Empty;
                        Guid.TryParse(itemId, out g);
                        try
                        {
                            using (PriemEntities context = new PriemEntities())
                            {
                                context.Person_deleteAllInfo(g);
                            }
                        }
                        catch (Exception ex)
                        {
                            WinFormsServ.Error("Ошибка удаления данных" + ex.Message);
                            //goto Next;
                        }
                    //Next: ;
                    }
                    MainClass.DataRefresh();
                }
            }
        }
    }
}