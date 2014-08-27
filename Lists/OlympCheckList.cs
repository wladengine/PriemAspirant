using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using PriemLib;

namespace Priem
{
    public partial class OlympCheckList : Form
    {
        public OlympCheckList()
        {
            InitializeComponent();
            UpdateDataGrid();
        }

        private void UpdateDataGrid()
        {
            //заполняем грид
            using (PriemEntities context = new PriemEntities())
            {
                var src = (from x in context.Olympiads
                           where (x.OlympTypeId == 2 || x.OlympTypeId == 3)
                           && (chbShowChecked.Checked ? true : (x.OlympiadCheckedByRectorat == null))
                           select new
                           {

                               x.Id,
                               x.AbiturientId,
                               IsNotChecked = x.OlympiadCheckedByRectorat == null ? 1 : 0,
                               x.Abiturient.Person.Surname,
                               x.Abiturient.Person.Name,
                               x.Abiturient.Person.SecondName,
                               //x.Abiturient.Entry.LicenseProgramCode,
                               //x.Abiturient.Entry.LicenseProgramName,
                               //x.Abiturient.Entry.ObrazProgramCrypt,
                               //x.Abiturient.Entry.ObrazProgramName,
                               //x.Abiturient.Entry.ProfileName,
                               OlympType = x.OlympType.Name,
                               OlympValue = x.OlympValue.Name,
                               OlympSubject = x.OlympSubject.Name,
                               x.DocumentSeries,
                               x.DocumentNumber,
                               x.DocumentDate
                           }).Distinct().ToList().Select(x => new
                           {
                               x.Id, x.AbiturientId, x.IsNotChecked,
                               FIO = (x.Surname ?? "") + " " + (x.Name ?? "") + " " + (x.SecondName ?? ""),
                               x.OlympType,
                               x.OlympSubject,
                               x.OlympValue,
                               Document = (x.DocumentSeries ?? "") + " " + (x.DocumentNumber ?? ""),
                               x.DocumentDate
                           }).OrderBy(x => x.FIO).ThenBy(x => x.OlympType).ThenBy(x => x.OlympSubject).ThenBy(x => x.OlympValue);

                dgv.DataSource = Util.ConvertToDataTable(src.ToArray());

                if (!dgv.Columns.Contains("Check"))
                {
                    DataGridViewButtonColumn dgvbc = new DataGridViewButtonColumn();
                    dgvbc.Text = "Проверено";
                    dgvbc.Name = "Check";
                    dgvbc.HeaderText = "Проверка дипломов";
                    dgvbc.UseColumnTextForButtonValue = true;
                    dgvbc.Visible = true;
                    dgv.Columns.Add(dgvbc);
                }

                dgv.Columns["Id"].Visible = false;
                dgv.Columns["AbiturientId"].Visible = false;
                dgv.Columns["IsNotChecked"].Visible = false;

                dgv.Columns["FIO"].HeaderText = "ФИО";
                dgv.Columns["FIO"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                dgv.Columns["OlympType"].HeaderText = "Тип олимпиады";
                dgv.Columns["OlympSubject"].HeaderText = "Предмет";
                dgv.Columns["OlympValue"].HeaderText = "Диплом";
                dgv.Columns["Document"].HeaderText = "Серия и номер";
                dgv.Columns["DocumentDate"].HeaderText = "Дата выдачи";

                dgv.Update();

                lblCount.Text = "Всего: " + dgv.Rows.Count;
            }
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (e.ColumnIndex == dgv.Columns["Check"].Index)
            {
                //check by rectorat
                if (MainClass.IsPrintOrder() || MainClass.IsPasha() || MainClass.IsOwner())
                {
                    if (MessageBox.Show("Проставить отметку о проверке?", "Внимание", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                    {
                        using (PriemEntities context = new PriemEntities())
                        {
                            context.OlympiadCheckedByRectorat_Insert((Guid)dgv["Id", e.RowIndex].Value);
                            UpdateDataGrid();
                        }
                    }
                }
                else
                {
                    WinFormsServ.Error("Недостаточно прав (пользователь не является участником роли PrintOrder)");
                }
            }
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (e.ColumnIndex != dgv.Columns["Check"].Index)
            {
                //open olympiad card
                new OlympCard(dgv["Id", e.RowIndex].Value.ToString(), (Guid)dgv["AbiturientId", e.RowIndex].Value, false).Show();
            }
        }

        private void dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            if (e.ColumnIndex != dgv.Columns["FIO"].Index)
                return;

            if (dgv["IsNotChecked", e.RowIndex].Value.ToString() == "0")
                e.CellStyle.BackColor = Color.GreenYellow;
        }

        private void chbShowChecked_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }
    }
}
