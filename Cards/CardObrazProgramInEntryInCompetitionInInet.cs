using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PriemLib;

namespace Priem
{
    public partial class CardObrazProgramInEntryInCompetitionInInet : Form
    {
        private ShortCompetition comp;
        public CardObrazProgramInEntryInCompetitionInInet(ShortCompetition _comp)
        {
            this.MdiParent = MainClass.mainform;
            InitializeComponent();
            comp = _comp;
            tbLicenseProgramName.Text = _comp.LicenseProgramName;
            InitGrid();
        }

        private void InitGrid()
        {
            var src =
                comp.lstObrazProgramsInEntry.Select(x => new
                {
                    x.Id,
                    x.ObrazProgramInEntryPriority,
                    x.ObrazProgramName,
                    HasProfiles = x.ListProfiles.Count > 0 ? "по профилям" : ""
                }).ToList();

            dgvObrazProgramInEntryList.DataSource = src;
            dgvObrazProgramInEntryList.Columns["Id"].Visible = false;

            dgvObrazProgramInEntryList.Columns["ObrazProgramInEntryPriority"].HeaderText = "Приоритет";
            dgvObrazProgramInEntryList.Columns["ObrazProgramName"].HeaderText = "Обр. программа";
            dgvObrazProgramInEntryList.Columns["ObrazProgramName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvObrazProgramInEntryList.Columns["HasProfiles"].HeaderText = "";
            dgvObrazProgramInEntryList.Columns["HasProfiles"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
        }

        private void dgvObrazProgramInEntryList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dgvObrazProgramInEntryList["HasProfiles", e.RowIndex].Value.ToString() == "по профилям")
            {
                Guid gId = (Guid)dgvObrazProgramInEntryList["Id", e.RowIndex].Value;

                var OPIE = comp.lstObrazProgramsInEntry.Where(x => x.Id == gId).FirstOrDefault();

                if (OPIE != null && OPIE.ListProfiles.Count > 0)
                {
                    var crd = new CardProfileInObrazProgramInEntryInCompetition(OPIE, comp.LicenseProgramName);
                    crd.Show();
                }
            }
        }
    }
}
