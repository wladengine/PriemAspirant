using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using EducServLib;
using BaseFormsLib;

namespace Priem
{
    public partial class ProtocolList : BaseFormEx
    {
        protected DateTime ProtocolDate
        {
            get { return _protocolDate; }
            set { _protocolDate = value; }
        }

        protected string ProtocolName
        {
            get { return _protocolName; }
            set { _protocolName = value; }
        }

        protected string ProtocolReason
        {
            get { return _protocolReason; }
            set { _protocolReason = value; }
        }

        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }

        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }

        public int? StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }

        public Guid? ProtocolNumId
        {
            get
            {
                string prId = ComboServ.GetComboId(cbProtocolNum);
                if (string.IsNullOrEmpty(prId))
                    return null;
                else
                    return new Guid(prId);
            }
        }
    }
}
