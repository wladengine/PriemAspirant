using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using BaseFormsLib;
using PriemLib;

namespace Priem
{
    //public delegate void DataRefreshHandler();
    //public delegate void ProtocolRefreshHandler();

    public static partial class MainClassCards
    {
        public static void OpenCardPerson(string personId, BaseFormEx formOwner, int? rowIndex)
        {
            foreach (Form frmChild in MainClass.mainform.MdiChildren)
            {
                if (frmChild is CardPerson)
                {
                    if (((CardPerson)frmChild).Id != null)
                    {
                        if (((CardPerson)frmChild).Id.CompareTo(personId) == 0)
                        {
                            frmChild.Focus();
                            return;
                        }
                    }
                }
            }

            CardPerson crd;
            if (formOwner == null && rowIndex == -1)
                crd = new CardPerson(personId);
            else
                crd = new CardPerson(personId, rowIndex, formOwner);
            crd.Show();
        }

        public static void OpenCardAbit(string abId, BaseFormEx formOwner, int? rowIndex)
        {
            foreach (Form frmChild in MainClass.mainform.MdiChildren)
            {
                if (frmChild is CardAbit)
                {
                    if (((CardAbit)frmChild).Id != null)
                    {
                        if (((CardAbit)frmChild).Id.CompareTo(abId) == 0)
                        {
                            frmChild.Focus();
                            return;
                        }
                    }
                }
            }

            new CardAbit(abId, rowIndex, formOwner).Show();
        }

        public static void OpenNewProtocol(ProtocolList formOwner, int iFacultyId, int iStudyFormId, int iStudyBasisId, ProtocolTypes _pType)
        {
            foreach (Form frmChild in MainClass.mainform.MdiChildren)
            {
                if (frmChild is ProtocolCard)
                {
                    //фокус на не закрытый протокол
                    frmChild.Focus();
                    return;
                }
            }

            ProtocolCard p;
            switch (_pType)
            {
                case ProtocolTypes.EnableProtocol: { p = new EnableProtocol(formOwner, iFacultyId, iStudyBasisId, iStudyFormId); break; }
                case ProtocolTypes.DisEnableProtocol: { p = new DisEnableProtocol(formOwner, iFacultyId, iStudyBasisId, iStudyFormId); break; }
                case ProtocolTypes.ChangeCompCelProtocol: { p = new ChangeCompCelProtocol(formOwner, iFacultyId, iStudyBasisId, iStudyFormId); break; }
                case ProtocolTypes.ChangeCompBEProtocol: { p = new ChangeCompBEProtocol(formOwner, iFacultyId, iStudyBasisId, iStudyFormId); break; }
                default: { p = new EnableProtocol(formOwner, iFacultyId, iStudyBasisId, iStudyFormId); break; }
            }

            p.Show();
        }

    }
}
