using EducServLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Objects;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PriemLib;

namespace Priem
{
    public partial class CardCompetitionInInet : Form
    {
        public event UpdateCompetitionListHandler OnUpdate;
        private ShortCompetition _competition;
        public CardCompetitionInInet(ShortCompetition c)
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            _competition = c;
            InitForm();
        }

        private void InitForm()
        {
            FillStudyLevel();
            StudyLevelId = _competition.StudyLevelId;

            IsReduced = _competition.IsReduced;
            IsSecond = _competition.IsSecond;

            FillLicenseProgram();
            LicenseProgramId = _competition.LicenseProgramId;
            FillObrazProgram();
            ObrazProgramId = _competition.ObrazProgramId;
            FillProfile();
            ProfileId = _competition.ProfileId;
            FillFaculty();
            FacultyId = _competition.FacultyId;
            FillStudyForm();
            StudyFormId = _competition.StudyFormId;
            FillStudyBasis();
            StudyBasisId = _competition.StudyBasisId;
            FillCompetition();
            ComboServ.FillCombo(cbCelCompetition, HelpClass.GetComboListByTable("ed.CelCompetition"), true, false);
            CompetitionId = _competition.CompetitionId;
            UpdateAfterCompetition();

            Priority = _competition.Priority;
            DocDate = _competition.DocDate;
            IsGosLine = _competition.IsGosLine;
            IsCommonRussianCompetition = _competition.IsCommonRussianCompetition;
            if (IsCommonRussianCompetition)
                chbIsCommonRussianCompetition.Visible = true;

            InitHandlers();

            if (_competition.HasInnerPriorities)
                btnHasInnerObrazProgram.Visible = true;
        }

        protected void InitHandlers()
        {
            chbIsReduced.CheckedChanged -= new EventHandler(chbIsSecond_CheckedChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);
            cbObrazProgram.SelectedIndexChanged += new EventHandler(cbObrazProgram_SelectedIndexChanged);
            cbProfile.SelectedIndexChanged += new EventHandler(cbProfile_SelectedIndexChanged);
            cbStudyForm.SelectedIndexChanged += new EventHandler(cbStudyForm_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
            cbCompetition.SelectedIndexChanged += new EventHandler(cbCompetition_SelectedIndexChanged);
        }
        protected void NullHandlers()
        {
            chbIsReduced.CheckedChanged -= new EventHandler(chbIsSecond_CheckedChanged);
            cbLicenseProgram.SelectedIndexChanged -= new EventHandler(cbLicenseProgram_SelectedIndexChanged);
            cbObrazProgram.SelectedIndexChanged -= new EventHandler(cbObrazProgram_SelectedIndexChanged);
            cbProfile.SelectedIndexChanged -= new EventHandler(cbProfile_SelectedIndexChanged);
            cbStudyForm.SelectedIndexChanged -= new EventHandler(cbStudyForm_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged -= new EventHandler(cbStudyBasis_SelectedIndexChanged);
            cbCompetition.SelectedIndexChanged -= new EventHandler(cbCompetition_SelectedIndexChanged);
        }

        void chbIsSecond_CheckedChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
        }
        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillObrazProgram();
        }
        void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillProfile();
            FillFaculty();
            FillStudyForm();
            FillStudyBasis();
        }
        void cbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillFaculty();
            FillStudyForm();
            FillStudyBasis();
        }
        void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillStudyBasis();
        }
        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillCompetition();
        }
        void cbCompetition_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateAfterCompetition();
        }

        private IEnumerable<Entry> GetEntry(PriemEntities context)
        {
            IEnumerable<Entry> entry = context.Entry;
            entry = entry.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced);

            return entry;
        }

        private void FillStudyLevel()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from ent in context.qStudyLevel
                                                               select new
                                                               {
                                                                   Id = ent.Id,
                                                                   Name = ent.Name
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    ComboServ.FillCombo(cbStudyLevel, lst, false, false);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillStudyLevel" + exc.Message);
            }
        }
        private void FillLicenseProgram()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from ent in GetEntry(context)
                                                               where ent.StudyLevelId == StudyLevelId
                                                               select new
                                                               {
                                                                   Id = ent.LicenseProgramId,
                                                                   Name = ent.SP_LicenseProgram.Name
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    ComboServ.FillCombo(cbLicenseProgram, lst, false, false);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillLicenseProgram" + exc.Message);
            }
        }
        private void FillObrazProgram()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from ent in GetEntry(context)
                                                               where ent.LicenseProgramId == LicenseProgramId
                                                               select new
                                                               {
                                                                   Id = ent.ObrazProgramId,
                                                                   Name = ent.SP_ObrazProgram.Name,
                                                                   Crypt = ent.StudyLevel.Acronym + "." + ent.SP_ObrazProgram.Number + "." + MainClass.sPriemYear
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name + ' ' + u.Crypt)).ToList();

                    ComboServ.FillCombo(cbObrazProgram, lst, false, false);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillObrazProgram" + exc.Message);
            }
        }
        private void FillProfile()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from ent in GetEntry(context)
                                                               where ent.LicenseProgramId == LicenseProgramId && ent.ObrazProgramId == ObrazProgramId && ent.ProfileId != null
                                                               select new
                                                               {
                                                                   Id = ent.ProfileId,
                                                                   Name = ent.ProfileName
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    if (lst.Count() > 0)
                    {
                        ComboServ.FillCombo(cbProfile, lst, false, false);
                        cbProfile.Enabled = true;
                    }
                    else
                    {
                        ComboServ.FillCombo(cbProfile, new List<KeyValuePair<string, string>>(), true, false);
                        cbProfile.Enabled = false;
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillProfile" + exc.Message);
            }
        }
        private void FillFaculty()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from ent in GetEntry(context)
                                                               where ent.LicenseProgramId == LicenseProgramId
                                                               && ent.ObrazProgramId == ObrazProgramId
                                                               && (ProfileId == null ? ent.ProfileId == null : ent.ProfileId == ProfileId)
                                                               select new
                                                               {
                                                                   Id = ent.FacultyId,
                                                                   Name = ent.SP_Faculty.Name
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    ComboServ.FillCombo(cbFaculty, lst, false, false);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillFaculty" + exc.Message);
            }
        }
        private void FillStudyForm()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from ent in GetEntry(context)
                                                               where ent.LicenseProgramId == LicenseProgramId
                                                               && ent.ObrazProgramId == ObrazProgramId
                                                               && (ProfileId == null ? ent.ProfileId == null : ent.ProfileId == ProfileId)
                                                               && ent.FacultyId == FacultyId
                                                               select new
                                                               {
                                                                   Id = ent.StudyFormId,
                                                                   Name = ent.StudyForm.Name
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    ComboServ.FillCombo(cbStudyForm, lst, false, false);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillStudyForm" + exc.Message);
            }
        }
        private void FillStudyBasis()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from ent in GetEntry(context)
                                                               where ent.LicenseProgramId == LicenseProgramId
                                                               && ent.ObrazProgramId == ObrazProgramId
                                                               && (ProfileId == null ? ent.ProfileId == null : ent.ProfileId == ProfileId)
                                                               && ent.FacultyId == FacultyId
                                                               && ent.StudyFormId == StudyFormId
                                                               select new
                                                               {
                                                                   Id = ent.StudyBasisId,
                                                                   Name = ent.StudyBasis.Name
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    ComboServ.FillCombo(cbStudyBasis, lst, false, false);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillStudyBasis" + exc.Message);
            }
        }
        private void FillCompetition()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst = ((from cp in context.Competition
                                                               where cp.StudyBasisId == StudyBasisId
                                                               select new
                                                               {
                                                                   cp.Id,
                                                                   cp.Name
                                                               }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    ComboServ.FillCombo(cbCompetition, lst, false, false);

                    lst = ((from cp in context.Competition
                            where cp.Id == StudyBasisId && cp.Id < 6
                            select new
                            {
                                cp.Id,
                                cp.Name
                            }).Distinct()).ToList().Select(u => new KeyValuePair<string, string>(u.Id.ToString(), u.Name)).ToList();

                    ComboServ.FillCombo(cbOtherCompetition, lst, true, false);

                    if (StudyBasisId == 1)
                    {
                        chbIsListener.Checked = false;
                        chbIsListener.Enabled = false;
                        ComboServ.SetComboId(cbCompetition, 4);
                    }
                    else
                    {
                        chbIsListener.Enabled = true;
                        ComboServ.SetComboId(cbCompetition, 3);
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы FillCompetition" + exc.Message);
            }
        }

        private void UpdateAfterCompetition()
        {
            if (CompetitionId == 6)
            {
                cbOtherCompetition.SelectedIndex = 0;
                cbCelCompetition.SelectedIndex = 0;
                lblOtherCompetition.Visible = true;
                cbOtherCompetition.Visible = true;

                lblCelCompetition.Visible = true;
                cbCelCompetition.Visible = true;
                tbCelCompetitionText.Visible = true;
            }
            else
            {
                cbOtherCompetition.SelectedIndex = 0;
                lblOtherCompetition.Visible = false;
                cbOtherCompetition.Visible = false;
                tbCelCompetitionText.Text = "";
                tbCelCompetitionText.Visible = false;

                lblCelCompetition.Visible = false;
                cbCelCompetition.Visible = false;
                cbCelCompetition.SelectedIndex = 0;
            }
        }

        private bool CheckIsClosed(PriemEntities context)
        {
            bool isClosed = (from ent in context.qEntry
                             where ent.Id == EntryId
                             select ent.IsClosed).FirstOrDefault();
            return !isClosed;
        }

        private bool Check()
        {
            using (PriemEntities context = new PriemEntities())
            {
                if (LicenseProgramId == null || ObrazProgramId == null || FacultyId == null || StudyFormId == null || StudyBasisId == null)
                {
                    epError.SetError(cbLicenseProgram, "Прием документов на данную программу не осуществляется!");
                    return false;
                }
                else
                    epError.Clear();

                if (EntryId == null)
                {
                    epError.SetError(cbLicenseProgram, "Прием документов на данную программу не осуществляется!");
                    return false;
                }
                else
                    epError.Clear();

                if (!CheckIsClosed(context))
                {
                    epError.SetError(cbLicenseProgram, "Прием документов на данную программу закрыт!");
                    return false;
                }
                else
                    epError.Clear();

                if (DocDate > DateTime.Now)
                {
                    epError.SetError(dtDocDate, "Неправильная дата");
                    return false;
                }
                else
                    epError.Clear();

                if (!CheckIdent(context))
                {
                    WinFormsServ.Error("У абитуриента уже существует заявление на данный факультет, направление, профиль, форму и основу обучения!");
                    return false;
                }

                if (!CheckThreeAbits(context))
                {
                    WinFormsServ.Error("У абитуриента уже существует 3 заявления на различные образовательные программы!");
                    return false;
                }

                //List<int> lstCompetitionForOriginals = context.Competition.Where(x => x.NeedOriginals).Select(x => x.Id).ToList();
                //if ((CompetitionId.HasValue && lstCompetitionForOriginals.Contains(CompetitionId.Value)) && !HasOriginals)
                //{
                //    WinFormsServ.Error("Для данного типа конкурса требуется обязательная подача оригиналов документов об образовании");
                //    return false;
                //}
            }

            return true;
        }

        // проверка на уникальность заявления
        private bool CheckIdent(PriemEntities context)
        {
            ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

            if (_competition.PersonId != null)
                context.CheckAbitIdent(_competition.PersonId, EntryId, boolPar);

            return Convert.ToBoolean(boolPar.Value);
        }

        private bool CheckThreeAbits(PriemEntities context)
        {
            return SomeMethodsClass.CheckThreeAbits(context, _competition.PersonId, LicenseProgramId);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (Check())
            {
                if (LicenseProgramId.HasValue)
                    _competition.LicenseProgramId = LicenseProgramId.Value;
                if (ObrazProgramId.HasValue)
                    _competition.ObrazProgramId = ObrazProgramId.Value;
                _competition.ProfileId = ProfileId;
                if (CompetitionId.HasValue)
                    _competition.CompetitionId = CompetitionId.Value;
                _competition.CompetitionName = cbCompetition.Text;


                _competition.DocInsertDate = DocInsertDate ?? DateTime.Now;
                _competition.IsGosLine = IsGosLine;
                _competition.IsListener = IsListener;
                _competition.IsReduced = IsReduced;

                _competition.ChangeEntry();
                if (OnUpdate != null)
                    OnUpdate(_competition);

                _competition.HasOriginals = HasOriginals;
                _competition.IsCommonRussianCompetition = IsCommonRussianCompetition;

                this.Close();
            }
        }

        private void btnHasInnerObrazProgram_Click(object sender, EventArgs e)
        {
            new CardObrazProgramInEntryInCompetitionInInet(_competition).Show();
        }
    }
}
