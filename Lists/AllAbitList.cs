using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.Entity.Core.Objects;
using System.Linq;

using EducServLib;
using WordOut;
using BaseFormsLib;
using PriemLib;
namespace Priem
{
    public partial class AllAbitList : BookList
    { 
        //�����������
        public AllAbitList()
        {            
            InitializeComponent();
            
            Dgv = dgvAbitList;
            _tableName = "ed.qAbiturient";
            _title = "������ ������������ � ����������� �� ������ ����������";          

            InitControls();
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            
            Dgv.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(Dgv_CellDoubleClick);
            UpdateDataGrid();           
        }  

        //���������� �����
        protected override void GetSource()
        {
            _sQuery = @"SELECT ed.qAbitAll.Id, PersonNum as ��_�����, 
                     FIO as ���, 
                     RegNum as ���_�����, FacultyName as ���������, ObrazProgramCrypt as ���, 
                     LicenseProgramName as �����������, ProfileName as �������, 
                     StudyFormName as �����, StudyBasisName as ������, Priority AS ��������� 
                     FROM ed.qAbitAll INNER JOIN ed.extPersonAspirant ON ed.qAbitAll.PersonId =  ed.extPersonAspirant.Id              
                     WHERE personId in (SELECT distinct personId FROM ed.qAbiturient) ";

            string filter = MainClass.GetStLevelFilter("ed.qAbitAll");

            HelpClass.FillDataGrid(Dgv, _bdc, _sQuery, filter, " ORDER BY ���, ���_�����");
        }

        //����� �� ������
        private void tbNumber_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbitList, "��_�����", tbNumber.Text);
        }

        //����� �� ���
        private void tbFIO_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbitList, "���", tbFIO.Text);
        }

        protected override void OpenCard(string itemId, BaseFormEx formOwner, int? index)
        {
            return;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintClass.PrintAllToExcel(this);
        }
    }
}