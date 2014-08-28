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
    public partial class PersonStatRating : Form
    {
        private Guid _PersonId;

        public PersonStatRating(Guid PersonId)
        {
            InitializeComponent();
            _PersonId = PersonId;
            FillGrid();
        }

        private void FillGrid()
        {
            string query = @"SELECT 
SP_Faculty.Acronym AS 'Факультет',
LicenseProgramCode + ' ' + LicenseProgramName AS  'Направление',
ObrazProgramCrypt + ' ' + ObrazProgramName AS 'Обр. программа',
Entry.ProfileName AS 'Профиль',
StudyForm.Name AS 'Форма обучения',
StudyBasis.Acronym AS 'Основа обучения',
hlpStatRatingList.SUM AS 'Сумма баллов',
hlpStatRatingList.Rank AS 'Рейтинг',
(CASE WHEN Abiturient.CompetitionId IN (1,2,5,7,8) THEN 1 ELSE 0 END) AS VKs,
(CASE WHEN Abiturient.CompetitionId <> 6 AND Rank <= Entry.KCP THEN 1 ELSE (CASE WHEN Abiturient.CompetitionId = 6 AND Rank <= Entry.KCPCel THEN 1 ELSE 0 END) END) AS Green,
(CASE WHEN Abiturient.CompetitionId <> 6 THEN Entry.KCP ELSE Entry.KCPCel END) AS 'КЦ',
(CASE WHEN Abiturient.BackDoc = 'True' THEN 'Забрал' + (CASE WHEN Person.Sex = 'False' THEN 'а' ELSE '' END) + ' документы' + convert(nvarchar, BackDocDate, 104) ELSE 'Поданы ' + convert(nvarchar, DocDate, 104) END) AS 'Документы',
(CASE WHEN Abiturient.NotEnabled = 'True' THEN 'Не допущен' + (CASE WHEN Person.Sex = 'False' THEN 'а' ELSE '' END) ELSE '' END) AS 'Примечание',
(CASE WHEN Abiturient.HasOriginals = 'True' THEN 'Да' ELSE 'Нет' END) AS 'Оригиналы'
FROM ed.hlpStatRatingList
INNER JOIN ed.Abiturient ON Abiturient.Id = hlpStatRatingList.AbiturientId
INNER JOIN ed.Entry ON Entry.Id = Abiturient.EntryId
INNER JOIN ed.Person ON Person.Id = hlpStatRatingList.PersonId
INNER JOIN ed.StudyForm ON StudyForm.Id = Entry.StudyFormId
INNER JOIN ed.StudyBasis ON StudyBasis.Id = Entry.StudyBasisId
INNER JOIN ed.SP_Faculty ON SP_Faculty.Id = Entry.FacultyId
WHERE hlpStatRatingList.PersonId = @PersonId";
            DataTable tbl = MainClass.Bdc.GetDataSet(query, new SortedList<string, object>() { { "@PersonId", _PersonId } }).Tables[0];

            dgv.DataSource = tbl;
            dgv.Columns["VKs"].Visible = false;
            dgv.Columns["Green"].Visible = false;

            dgv.Columns["Факультет"].Width = 50;

            dgv.Columns["Сумма баллов"].Width = 50;
            dgv.Columns["Рейтинг"].Width = 50;
            dgv.Columns["КЦ"].Width = 30;

            dgv.Columns["Форма обучения"].Width = 50;
            dgv.Columns["Основа обучения"].Width = 50;

            dgv.Columns["Оригиналы"].Width = 65;

            dgv.Columns["Примечание"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            dgv.Columns["Документы"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;

            query = "SELECT FIO FROM ed.extPersonAspirant WHERE Id=@Id";
            string fio = MainClass.Bdc.GetStringValue(query, new SortedList<string, object>() { { "@Id", _PersonId } });
            lblFIO.Text = fio;
        }

        private void dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            if (e.ColumnIndex != dgv.Columns["Факультет"].Index)
                return;

            if ((int)dgv.Rows[e.RowIndex].Cells["Green"].Value == 1)
            {
                dgv.Rows[e.RowIndex].Cells["Рейтинг"].Style.BackColor = Color.Green;
            }
            if ((int)dgv.Rows[e.RowIndex].Cells["VKs"].Value == 1)
            {
                dgv.Rows[e.RowIndex].Cells["Факультет"].Style.BackColor = Color.LightSkyBlue;
            }
            if (dgv.Rows[e.RowIndex].Cells["Оригиналы"].Value.ToString() == "Да")
            {
                dgv.Rows[e.RowIndex].Cells["Оригиналы"].Style.BackColor = Color.Gold;
            }
        }
        private void btnOpenPersonCard_Click(object sender, EventArgs e)
        {
            MainClassCards.OpenCardPerson(_PersonId.ToString(), null, null);
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
            //just 4 test
            //Print.PrintRatingProtocol(1, 1, 1, 456, 29, null, false, 80, MainClass.dirTemplates + @"\Data.pdf", false, false, false);
        }
    }
}
