using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using BaseFormsLib;
using EducServLib;

namespace Priem
{
    public partial class EgeStatistics : BaseFormEx
    {
        public int FacultyId
        {
            get
            {
                int? iLP = ComboServ.GetComboIdInt(cbFaculty);
                return iLP.Value;
            }
        }

        public int? LicenseProgramId
        {
            get
            {
                if (cbLicenseProgram.Text == ComboServ.ALL_VALUE)
                    return null;
                int? iLP = ComboServ.GetComboIdInt(cbLicenseProgram);
                return iLP;
            }
        }

        public int? ObrazProgramId
        {
            get
            {
                if (cbObrazProgram.Text == ComboServ.ALL_VALUE)
                    return null;
                int? iOP = ComboServ.GetComboIdInt(cbObrazProgram);
                return iOP;
            }
        }

        public Guid? ProfileId
        {
            get
            {
                if (cbProfile.Text == ComboServ.ALL_VALUE || cbProfile.Text == ComboServ.NO_VALUE)
                    return null;
                Guid g = Guid.Empty;
                if (!Guid.TryParse(ComboServ.GetComboId(cbProfile), out g))
                    return null;
                else
                    return g;
            }
        }

        public int? StudyFormId
        {
            get 
            {
                if (cbStudyForm.Text == ComboServ.ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbStudyForm);
            }
        }

        public int? StudyBasisId
        {
            get
            {
                if (cbStudyBasis.Text == ComboServ.ALL_VALUE)
                    return null;
                return ComboServ.GetComboIdInt(cbStudyBasis);
            }
        }
    }
}
