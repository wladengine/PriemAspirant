using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.Objects;
using System.Linq;

using EducServLib;
using WordOut;
using Excel = Microsoft.Office.Interop.Excel;
using BaseFormsLib;
using PriemLib;
namespace Priem
{
    public partial class AllAbitList : BookList
    { 
        //конструктор
        public AllAbitList()
        {            
            InitializeComponent();
            
            Dgv = dgvAbitList;
            _tableName = "ed.qAbiturient";
            _title = "Список абитуриентов с заявлениями на другие факультеты";          

            InitControls();
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            
            Dgv.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(Dgv_CellDoubleClick);
            UpdateDataGrid();           
        }  

        //обновление грида
        protected override void GetSource()
        {
            _sQuery = @"SELECT ed.qAbitAll.Id, PersonNum as Ид_номер, 
                     FIO as ФИО, 
                     RegNum as Рег_номер, FacultyName as Факультет, ObrazProgramCrypt as Код, 
                     LicenseProgramName as Направление, ProfileName as Профиль, 
                     StudyFormName as Форма, StudyBasisName as Основа, Priority AS Приоритет 
                     FROM ed.qAbitAll INNER JOIN ed.extPersonAspirant ON ed.qAbitAll.PersonId =  ed.extPersonAspirant.Id              
                     WHERE personId in (SELECT distinct personId FROM ed.qAbiturient) ";

            string filter = MainClass.GetStLevelFilter("ed.qAbitAll");

            HelpClass.FillDataGrid(Dgv, _bdc, _sQuery, filter, " ORDER BY ФИО, Рег_номер");
        }

        //поиск по номеру
        private void tbNumber_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbitList, "Ид_номер", tbNumber.Text);
        }

        //поиск по фио
        private void tbFIO_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbitList, "ФИО", tbFIO.Text);
        }

        protected override void OpenCard(string itemId, BaseFormEx formOwner, int? index)
        {
            return;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Файлы Excel (.xls)|*.xls";
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                try
                {
                    List<string> columnList = new List<string>();
                    foreach (DataGridViewColumn column in dgvAbitList.Columns)
                        if (column.Visible)
                            columnList.Add(column.Name);
                    Excel.Application exc = new Excel.Application();
                    Excel.Workbook wb = exc.Workbooks.Add(System.Reflection.Missing.Value);
                    Excel.Worksheet ws = (Excel.Worksheet)exc.ActiveSheet;

                    //почему имя рабочего листа <<в ВТБ>>???
                    ws.Name = "заявления на другие факультеты";

                    ws.Cells[1, 1] = "№ п/п";
                    for (int j = 0; j < columnList.Count; j++)
                        ws.Cells[1, j + 2] = dgvAbitList.Columns[columnList[j]].HeaderText;

                    //Excel.Range rg = ws.get_Range(ws.Cells[1, 1], ws.Cells[dgvAbitList.RowCount, dgvAbitList.ColumnCount]);


                    //cellXY.NumberFormat = 0.00;

                    //ws.Cells[2, 5] = "12";

                    int i = 0;
                    // печать из грида
                    foreach (DataGridViewRow dgvr in dgvAbitList.Rows)
                    {
                        ws.Cells[i + 2, 1] = (i + 1).ToString();
                        for (int j = 0; j < columnList.Count; j++)
                            ws.Cells[i + 2, j + 2] = dgvAbitList.Rows[i].Cells[columnList[j]].Value.ToString();

                        i++;
                    }

                    wb.SaveAs(sfd.FileName, Excel.XlFileFormat.xlExcel7,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        Excel.XlSaveAsAccessMode.xlExclusive,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value);
                    exc.Visible = true;
                    //по идее, из приложения Excel выходить не надо, пользователь в силах это сделать и самостоятельно, когда насмотрится на свой отсчёт
                    //exc.Quit();
                    //exc = null;

                }
                catch (System.Runtime.InteropServices.COMException exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
            //На всякий случай
            sfd.Dispose();
        }
    }
}