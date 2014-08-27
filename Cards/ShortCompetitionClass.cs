using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PriemLib;
namespace Priem
{
    public class ShortCompetition
    {
        public Guid Id { get; private set; }
        public Guid PersonId { get; private set; }
        public Guid EntryId { get; private set; }
        public Guid CommitId { get; private set; }

        public int? VersionNum { get; private set; }
        public DateTime? VersionDate { get; private set; }

        public int StudyLevelId { get; set; }
        public string StudyLevelName { get; set; }
        public int LicenseProgramId { get; set; }
        public string LicenseProgramName { get; set; }
        public int ObrazProgramId { get; set; }
        public string ObrazProgramName { get; set; }
        public Guid? ProfileId { get; set; }
        public string ProfileName { get; set; }
        public int FacultyId { get; set; }
        public string FacultyName { get; set; }

        public int StudyFormId { get; set; }
        public string StudyFormName { get; set; }
        public int StudyBasisId { get; set; }
        public string StudyBasisName { get; set; }

        public bool HasCompetition { get; set; }
        public int CompetitionId { get; set; }
        public string CompetitionName { get; set; }

        public int Priority { get; set; }
        public int? OtherCompetitionId { get; set; }
        public int? CelCompetitionId { get; set; }
        public string CelCompetitionText { get; set; }

        public DateTime DocDate { get; set; }
        public DateTime DocInsertDate { get; set; }

        public bool HasOriginals { get; set; }

        public bool IsListener { get; set; }
        public bool IsReduced { get; set; }
        public bool IsSecond { get; set; }
        public bool IsGosLine { get; set; }
        public bool IsCommonRussianCompetition { get; set; }

        public int Barcode { get; set; }

        public bool HasInnerPriorities { get; set; }
        public bool IsApprovedByComission { get; set; }
        public string ApproverName { get; set; }
        public List<ShortObrazProgramInEntry> lstObrazProgramsInEntry { get; set; }

        public ShortCompetition(Guid _Id, Guid _CommitId, Guid _EntryId, Guid _PersonId, int? _VersionNum, DateTime? _VersionDate)
        {
            Id = _Id;
            CommitId = _CommitId;
            EntryId = _EntryId;
            PersonId = _PersonId;
            VersionNum = _VersionNum;
            VersionDate = _VersionDate;
        }

        public void ChangeEntry()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var Entry = context.Entry
                    .Where(x => x.LicenseProgramId == LicenseProgramId
                        && x.ObrazProgramId == ObrazProgramId
                        && x.StudyLevelId == StudyLevelId
                        && x.StudyFormId == StudyFormId
                        && x.StudyBasisId == StudyBasisId
                        && (ProfileId.HasValue ? x.ProfileId == ProfileId : x.ProfileId == null)).FirstOrDefault();

                if (Entry != null)
                    EntryId = Entry.Id;
            }
        }
    }
    public class ShortObrazProgramInEntry
    {
        public Guid Id { get; private set; }
        public string ObrazProgramName { get; private set; }
        public int ObrazProgramInEntryPriority { get; set; }
        public int CurrVersion { get; set; }
        public DateTime CurrDate { get; set; }
        public List<ShortProfileInObrazProgramInEntry> ListProfiles { get; set; }

        public ShortObrazProgramInEntry(Guid _id, string _obrazProgramName)
        {
            Id = _id;
            ObrazProgramName = _obrazProgramName;
        }
    }
    public class ShortProfileInObrazProgramInEntry
    {
        public Guid Id { get; private set; }
        public string ProfileName { get; private set; }
        public int ProfileInObrazProgramInEntryPriority { get; set; }

        public ShortProfileInObrazProgramInEntry(Guid _Id, string _ProfileName)
        {
            Id = _Id;
            ProfileName = _ProfileName;
        }
    }

    public delegate void UpdateCompetitionListHandler(ShortCompetition c);
}
