using EducServLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PriemLib;
namespace Priem
{
    partial class CardCompetitionInInet
    {
        public Guid? EntryId
        {
            get
            {
                try
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        Guid? entId = (from ent in context.qEntry
                                       where ent.IsSecond == IsReduced
                                        && ent.LicenseProgramId == LicenseProgramId
                                        && ent.ObrazProgramId == ObrazProgramId
                                        && (ProfileId == null ? ent.ProfileId == null : ent.ProfileId == ProfileId)
                                        && ent.StudyFormId == StudyFormId
                                        && ent.StudyBasisId == StudyBasisId
                                       select ent.Id).FirstOrDefault();
                        return entId;
                    }
                }
                catch
                {
                    return null;
                }
            }
        }

        public int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
            set { ComboServ.SetComboId(cbStudyLevel, value); }
        }
        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
        public int? LicenseProgramId
        {
            get { return ComboServ.GetComboIdInt(cbLicenseProgram); }
            set { ComboServ.SetComboId(cbLicenseProgram, value); }
        }
        public int? ObrazProgramId
        {
            get { return ComboServ.GetComboIdInt(cbObrazProgram); }
            set { ComboServ.SetComboId(cbObrazProgram, value); }
        }
        public Guid? ProfileId
        {
            get
            {
                string prId = ComboServ.GetComboId(cbProfile);
                if (string.IsNullOrEmpty(prId))
                    return null;
                else
                    return new Guid(prId);
            }
            set
            {
                if (value == null)
                    ComboServ.SetComboId(cbProfile, (string)null);
                else
                    ComboServ.SetComboId(cbProfile, value.ToString());
            }
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

        public int? CompetitionId
        {
            get { return ComboServ.GetComboIdInt(cbCompetition); }
            set { ComboServ.SetComboId(cbCompetition, value); }
        }

        public bool HasOriginals
        {
            get { return chbHasOriginals.Checked; }
            set { chbHasOriginals.Checked = value; }
        }

        public bool IsCommonRussianCompetition
        {
            get { return chbIsCommonRussianCompetition.Checked; }
            set { chbIsCommonRussianCompetition.Checked = value; }
        }

        public bool IsReduced
        {
            get { return chbIsReduced.Checked; }
            set { chbIsReduced.Checked = value; }
        }
        public bool IsSecond
        {
            get { return chbIsSecond.Checked; }
            set { chbIsSecond.Checked = value; }
        }
        public bool IsListener
        {
            get { return chbIsListener.Checked; }
            set { chbIsListener.Checked = value; }
        }
        public bool IsGosLine
        {
            get { return chbIsGosLine.Checked; }
            set { chbIsGosLine.Checked = value; }
        }

        public DateTime? DocDate
        {
            get { return dtDocDate.Value.Date; }
            set
            {
                if (value.HasValue)
                    dtDocDate.Value = value.Value;
            }
        }
        public DateTime? DocInsertDate
        {
            get { return dtDocInsertDate.Value.Date; }
            set
            {
                if (value.HasValue)
                    dtDocInsertDate.Value = value.Value;
            }
        }

        public int? OtherCompetitionId
        {
            get
            {
                if (CompetitionId == 6 && cbOtherCompetition.SelectedIndex != 0)
                    return ComboServ.GetComboIdInt(cbOtherCompetition);
                else
                    return null;
            }
            set
            {
                if (CompetitionId == 6)
                    if (value != null)
                        ComboServ.SetComboId(cbOtherCompetition, value);
            }
        }
        public int? CelCompetitionId
        {
            get
            {
                if (CompetitionId == 6 && cbCelCompetition.SelectedIndex != 0)
                    return ComboServ.GetComboIdInt(cbCelCompetition);
                else
                    return null;
            }
            set
            {
                if (CompetitionId == 6)
                    if (value != null)
                        ComboServ.SetComboId(cbCelCompetition, value);
            }
        }
        public string CelCompetitionText
        {
            get
            {
                if (CompetitionId == 6)
                    return tbCelCompetitionText.Text;
                else
                    return string.Empty;
            }
            set
            {
                if (CompetitionId == 6)
                    tbCelCompetitionText.Text = value;
            }
        }

        public double? Priority
        {
            get
            {
                double j;
                if (double.TryParse(tbPriority.Text.Trim(), out j))
                    return j;
                else
                    return null;
            }
            set { tbPriority.Text = Util.ToStr(value); }
        }
    }
}
