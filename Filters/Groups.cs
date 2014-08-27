using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using PriemLib;
namespace Priem
{
    public partial class Groups : Form
    {
        private FormFilter _owner;
        //private List<ListItem> _columnList;

        private bool flag;

        public Groups(FormFilter la)
        {
            InitializeComponent();

            _owner = la;

            foreach (ListItem li in FillGroups())
                lbNo.Items.Add(li);
            chbPrintGroup.Checked = true;
        }

        private List<ListItem> FillGroups()
        {
            List<ListItem> list = new List<ListItem>();
            list.Add(new ListItem("ФИО","Фамилия Имя Отчество"));
            list.Add(new ListItem("Направление", "Направление"));
            list.Add(new ListItem("Образ_программа", "Образовательная программа"));
            list.Add(new ListItem("Тип_паспорта","Тип документа"));            
            list.Add(new ListItem("Забрал_док", "Забрал документы"));
            list.Add(new ListItem("Пол_мужской", "Пол"));
            list.Add(new ListItem("Медалист", "Медалист"));
            list.Add(new ListItem("Год_выпуска", "Год выпуска"));
            list.Add(new ListItem("Ин_язык", "Иностранный язык"));
            list.Add(new ListItem("Профиль", "Профиль"));
            list.Add(new ListItem("Форма_обучения", "Форма обучения"));
            list.Add(new ListItem("Основа_обучения", "Основа обучения"));
            list.Add(new ListItem("Тип_конкурса", "Тип конкурса"));
            list.Add(new ListItem("второе_высшее", "Второе высшее"));
            list.Add(new ListItem("Страна", "Страна"));
            list.Add(new ListItem("Гражданство", "Гражданство"));
            list.Add(new ListItem("Оплатил", "Оплатил"));
            list.Add(new ListItem("Слушатель", "Слушатель"));           
            list.Add(new ListItem("Предоставлять_общежитие_поступление", "Предоставлять общежитие на время поступления"));
            list.Add(new ListItem("Регион", "Регион"));
            list.Add(new ListItem("Тип_уч_заведения", "Тип учебного заведения"));
            //list.Add(new ListItem("", ""));
                                      
            return list;
        }           

        //отмена
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //подтверждение с передачей результата
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (lbYes.Items.Count > 0)
            {
                List<ListItem> list = new List<ListItem>();
                foreach (ListItem li in lbYes.Items)
                {
                    list.Add(li);
                    if(_owner is ListAbit)
                        MainClass._config.AddColumnNameAbit(li.Id);
                    else if(_owner is ListPersonFilter)
                        MainClass._config.AddColumnNamePerson(li.Id);
                }
                _owner.GroupList = list;
                _owner.GroupPrint = chbPrintGroup.Checked;

                _owner.UpdateDataGrid();
            }            
            this.Close();
        }

        //функции переноса строк
        private void btnLeft_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbNo, lbYes, false);
        }

        //
        private void btnRight_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbYes, lbNo, false);
        }

        //
        private void btnLeftAll_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbNo, lbYes, true);
        }

        //
        private void btnRightAll_Click(object sender, EventArgs e)
        {
            WinFormsServ.MoveRows(lbYes, lbNo, true);
        }

        //стрелка вверх
        private void btnUp_Click(object sender, EventArgs e)
        {
            flag = true;

            if (lbYes.Items.Count == 0)
                return;

            int i = lbYes.SelectedIndex;
            if (i == 0)
                return;

            object obj = lbYes.Items[i];
            lbYes.Items.RemoveAt(i);
            lbYes.Items.Insert(i - 1, obj);
            lbYes.SetSelected(i - 1, true);
            flag = false;
        }

        //стрелка вниз
        private void btnDown_Click(object sender, EventArgs e)
        {
            flag = true;

            if (lbYes.Items.Count == 0)
                return;

            int i = lbYes.SelectedIndex;
            if (i == lbYes.Items.Count - 1)
                return;

            object obj = lbYes.Items[i];
            lbYes.Items.RemoveAt(i);
            lbYes.Items.Insert(i + 1, obj);
            lbYes.SetSelected(i + 1, true);
            flag = false;                
        }

        private void lbYes_SelectedIndexChanged(object sender, EventArgs e)
        {
            //если жмут по стрелкам, ничего не происходит
            if (flag)
            {
                return;
            }

            if (lbYes.SelectedIndex < 0)
            {
                btnUp.Enabled = false;
                btnDown.Enabled = false;
            }
            else
            {
                btnUp.Enabled = true;
                btnDown.Enabled = true;
            }
        }
    }
}