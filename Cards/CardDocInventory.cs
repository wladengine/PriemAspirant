using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Objects;

using EducServLib;
using BaseFormsLib;
using PriemLib;
namespace Priem
{
    public partial class CardDocInventory : BaseForm
    {
        private Guid? _abitId;
        private List<int> _oldAbitDocsIds;

        protected IList<int> AbitDocsIds
        {
            get { return Util.GetListInt(mslDocs.SelectedList); }
            set { mslDocs.SelectedList = Util.GetListString(value); }
        }

        public CardDocInventory(Guid? abitId, bool readOnly)
        {
            InitializeComponent();
            _abitId = abitId;
            InitControls();

            if (readOnly)
            {
                mslDocs.Enabled = false;
                btnSave.Enabled = false;
            }
        }

        protected void InitControls()
        {
            InitFocusHandlers();

            using (PriemEntities context = new PriemEntities())
            {
                ObjectSet<AbitDoc> queryDocs = context.CreateObjectSet<AbitDoc>();
                Dictionary<string, string> dict = (from f in queryDocs.OrderBy("it.Name")
                                                   select new { key = f.Id, value = f.Name }).ToDictionary(item => item.key.ToString(), item => item.value);
                mslDocs.InitSource(dict);              

                List<int> lst = (from inD in context.DocInventory
                                 where inD.AbiturientId == _abitId
                                 select inD.AbitDocId).ToList<int>();

                _oldAbitDocsIds = lst;
                AbitDocsIds = lst;
            }            
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                // сохраняем
                if (_oldAbitDocsIds != null && _oldAbitDocsIds.Count > 0)
                {
                    foreach (int? abDocId in _oldAbitDocsIds.Except(AbitDocsIds))
                    {
                        // удаляем
                        context.DocInventory_Delete(_abitId, abDocId);
                    }
                }

                if (AbitDocsIds != null && AbitDocsIds.Count > 0)
                {
                    foreach (int abDocId in AbitDocsIds.Except(_oldAbitDocsIds))
                    {
                        // добавляем                                    
                        context.DocInventory_Insert(_abitId, abDocId, "");
                    }
                }
            }

            this.Close();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            Print.PrintDocInventory(AbitDocsIds, _abitId);
        }
    }
}
