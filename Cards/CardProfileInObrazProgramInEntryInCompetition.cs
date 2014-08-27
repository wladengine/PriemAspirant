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
    public partial class CardProfileInObrazProgramInEntryInCompetition : Form
    {
        private ShortObrazProgramInEntry OPIE;
        public CardProfileInObrazProgramInEntryInCompetition(ShortObrazProgramInEntry _OPIE, string LicenseProgramName)
        {
            this.MdiParent = MainClass.mainform;
            InitializeComponent();
            OPIE = _OPIE;
            tbLicenseProgramName.Text = LicenseProgramName;
            tbObrazProgramName.Text = OPIE.ObrazProgramName;

            FillGrid();
        }

        private void FillGrid()
        {
            var src = OPIE.ListProfiles.OrderBy(x => x.ProfileInObrazProgramInEntryPriority).Select(x => new { x.ProfileInObrazProgramInEntryPriority, x.ProfileName });
            dgvObrazProgramInEntryList.DataSource = src.ToList();

            dgvObrazProgramInEntryList.Columns["ProfileInObrazProgramInEntryPriority"].HeaderText = "Приоритет";
            dgvObrazProgramInEntryList.Columns["ProfileName"].HeaderText = "Название";
            dgvObrazProgramInEntryList.Columns["ProfileName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
    }
}
