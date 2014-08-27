using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using BDClassLib;
using WordOut;
using BaseFormsLib;

namespace Priem
{
    public class FormFilter : BaseFormEx
    {
        protected List<iFilter> _list;
        protected List<ListItem> _groupList;//группировки
        protected bool _groupPrint;

        public List<iFilter> FilterList
        {
            get
            {
                return _list;
            }
            set
            {
                _list = value;
            }
        }

        public List<ListItem> GroupList
        {
            get
            {
                return _groupList;
            }
            set
            {
                _groupList = value;
            }
        }

        public bool GroupPrint
        {
            get
            {
                return _groupPrint;
            }
            set
            {
                _groupPrint = value;
            }
        }

        public virtual void UpdateDataGrid() { }
    }
}
