using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.Linq;
using System.Transactions;
 
//using BDClassLib;
using EducServLib;
using WordOut;
using BaseFormsLib;
using PriemLib;

using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Priem
{
    public partial class ExamsVedList : BaseFormEx
    {
        private DBPriem bdc;
        private string sQuery;        
        private string sOrderby;
        
        //конструктор
        public ExamsVedList()
        {
            this.CenterToParent();
            this.MdiParent = MainClass.mainform;

            this.sQuery = @"SELECT DISTINCT ed.extPersonAspirant.Id, ed.extPersonAspirant.PersonNum as Ид_номер, ed.extPersonAspirant.Surname AS Фамилия, ed.extPersonAspirant.Name AS Имя, 
                            ed.extPersonAspirant.SecondName AS Отчество, ed.extPersonAspirant.BirthDate AS Дата_рождения 
                            FROM ed.extPersonAspirant INNER JOIN ed.ExamsVedHistory ON ed.ExamsVedHistory.PersonId = ed.extPersonAspirant.Id ";
                       
            this.sOrderby = " ORDER BY Фамилия ";           

            InitializeComponent();
            InitControls();
        }        

        //дополнительная инициализация контролов
        private void InitControls()
        {
            InitFocusHandlers();			
            bdc = MainClass.Bdc; 

            try
            {
                btnDeleteFromVed.Visible = btnDeleteFromVed.Enabled = false;
                btnUnload.Visible = btnUnload.Enabled = false;
                
                //наверное нужно писать права по отдельным группам. Чуть больше, зато на порядок понятнее. И чтобы "левых" вставок не было
                //Паше наклейки печатать нельзя??? Непорядок!!!
                if (MainClass.IsFacMain())
                {
                    btnCreate.Visible = true;
                    btnChange.Visible = true;                    
                    btnDelete.Visible = false;

                    tbCountCell.Visible = false;
                    lblCountCell.Visible = false;
                    btnLock.Visible = false;
                    btnCreateAdd.Visible = false;
                    btnPrintSticker.Visible = false;                    
                }
                else if (MainClass.IsCryptoMain() || MainClass.IsPasha())
                {
                    btnCreate.Visible = false;
                    btnChange.Visible = false;
                    btnDelete.Visible = false;
                    
                    tbCountCell.Visible = true;
                    lblCountCell.Visible = true;
                    btnLock.Visible = true;
                    btnCreateAdd.Visible = true;
                    btnPrintSticker.Visible = true;                    
                }                
                else
                {
                    btnCreate.Visible = false;
                    btnChange.Visible = false;
                    btnDelete.Visible = false;

                    tbCountCell.Visible = false;
                    lblCountCell.Visible = false;
                    btnLock.Visible = false;
                    btnCreateAdd.Visible = false;
                    btnPrintSticker.Visible = false;                    
                }

                if (MainClass.IsPasha())
                {
                    btnDelete.Visible = true;
                    btnChange.Visible = true;
                    btnDeleteFromVed.Visible = btnDeleteFromVed.Enabled = true;
                    btnUnload.Visible = btnUnload.Enabled = true;
                }

                if (MainClass.IsOwner())
                {
                    btnCreate.Visible = true;
                    btnChange.Visible = true;
                    btnDelete.Visible = true;
                    tbCountCell.Visible = true;
                    lblCountCell.Visible = true;
                    btnLock.Visible = true;
                    btnCreateAdd.Visible = true;
                    btnPrintSticker.Visible = true;                    
                }
                    
                tbCountCell.Text = (2).ToString();

                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.qFaculty", "ORDER BY Acronym"), false, false);
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, true);
                    
                    UpdateVedList();
                    UpdateDataGrid();

                    cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
                    cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
                    cbExamVed.SelectedIndexChanged += new EventHandler(cbExamVed_SelectedIndexChanged);
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при инициализации формы ведомостей: " + ex.Message);
            }
        }

        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateVedList();
        }

        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateVedList();
        }

        void cbExamVed_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }        

        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }

        public int? StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }

        public Guid? ExamVedId
        {
            get 
            {
                string valId = ComboServ.GetComboId(cbExamVed);
                if(string.IsNullOrEmpty(valId))
                    return null;
                else
                    return new Guid(valId);               
            }
            set 
            { 
                if(value == null)
                    ComboServ.SetComboId(cbExamVed, (string)null); 
                else
                    ComboServ.SetComboId(cbExamVed, value.ToString()); 
            }
        }

        //обновление списка 
        public void UpdateVedList()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    List<KeyValuePair<string, string>> lst =
                        ((from ent in context.extExamsVed
                          where MainClass.lstStudyLevelGroupId.Contains(ent.StudyLevelGroupId)
                          && ent.FacultyId == FacultyId
                          && (StudyBasisId != null ? ent.StudyBasisId == StudyBasisId : true == true)

                          select new
                          {
                              ent.Id,
                              ent.ExamName,
                              ent.Date,
                              StBasis = ent.StudyBasisId == null ? "" : ent.StudyBasisAcr,
                              AddVed = ent.IsAddVed ? " дополнительная" : "",
                              ent.AddCount
                          }).Distinct()).ToList()
                          .Select(u => new KeyValuePair<string, string>(u.Id.ToString(),
                                u.ExamName + ' ' + u.Date.ToShortDateString() + ' ' + u.StBasis + u.AddVed +
                                (u.AddCount > 1 ? "(" + Convert.ToString(u.AddCount) + ")" : ""))).ToList();

                    ComboServ.FillCombo(cbExamVed, lst, true, false);                    
                }            
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при обновлении списка ведомостей: " + ex.Message);
            }            
        }

        //обновление грида
        public virtual void UpdateDataGrid()
        {            
            //скрыли/показали кнопку, если надо            
            if (ExamVedId == null)
            {
                btnChange.Enabled = false;
                btnDelete.Enabled = false;
                btnCreateAdd.Enabled = false;
                btnLock.Enabled = false;

                lblLocked.Visible = false;

                tbCountCell.Enabled = false;
                btnPrintSticker.Enabled = false;

                dgvList.DataSource = null;
                dgvList.Update();
                return;  
            }
            else
            {
                if (cbExamVed.SelectedItem.ToString().Contains("дополнительная")) 
                    btnCreateAdd.Enabled = false; 
                else
                    btnCreateAdd.Enabled = true;
                
                using (PriemEntities context = new PriemEntities())
                {
                    bool isLocked = (from ev in context.extExamsVed
                                     where ev.Id == ExamVedId
                                     select ev.IsLocked).FirstOrDefault();
                    if (isLocked)
                    {
                        lblLocked.Visible = true;
                        btnChange.Enabled = false;
                        btnDelete.Enabled = false;
                        btnLock.Enabled = false;
                        btnPrintSticker.Enabled = true;
                        tbCountCell.Enabled = true;
                    }
                    else
                    {
                        lblLocked.Visible = false;
                        btnChange.Enabled = true;
                        btnDelete.Enabled = true;
                        btnLock.Enabled = true;
                        btnPrintSticker.Enabled = false;
                        tbCountCell.Enabled = false;
                    }
                }                     
            }                               
            
            //обработали номер            
            string sFilters = string.Format("WHERE ed.ExamsVedHistory.ExamsVedId = '{0}' ", ExamVedId.ToString());

            if (!dgvList.Columns.Contains("Number"))
            {
                dgvList.Columns.Add("Number", "№");
                dgvList.Update();
            }
            dgvList.Columns["Number"].DisplayIndex = 0; 

            HelpClass.FillDataGrid(dgvList, bdc, sQuery, sFilters, sOrderby);                       
        }
        
        //закрытие
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //печать
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (ExamVedId == null)
                return;

            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\Templates\CryptoExamsVed.dot", Application.StartupPath));
                TableDoc td = wd.Tables[0];
                
                using (PriemEntities context = new PriemEntities())
                {
                    extExamsVed ved = (from ev in context.extExamsVed
                                       where ev.Id == ExamVedId
                                       select ev).FirstOrDefault();

                    wd.Fields["Faculty"].Text = cbFaculty.Text.ToLower();
                    wd.Fields["Exam"].Text = ved.ExamName;
                    wd.Fields["StudyBasis"].Text = Util.ToStr(ved.StudyBasisId == null ? "все" : ved.StudyBasisName);
                    wd.Fields["Date"].Text = ved.Date.ToShortDateString();
                    wd.Fields["VedNum"].Text = ved.Number.ToString();

                    int i = 1;

                    // печать из грида
                    foreach (DataGridViewRow dgvr in dgvList.Rows)
                    {
                        td[0, i] = i.ToString();
                        td[1, i] = dgvr.Cells["Фамилия"].Value.ToString();
                        td[2, i] = dgvr.Cells["Имя"].Value.ToString();
                        td[3, i] = dgvr.Cells["Отчество"].Value.ToString();
                        td[4, i] = DateTime.Parse(dgvr.Cells["Дата_рождения"].Value.ToString()).ToShortDateString();
                        td[5, i] = dgvr.Cells["Ид_номер"].Value.ToString();
                        td[6, i] = FacultyId.ToString();
                        td[7, i] = ved.ExamName;
                        td[8, i] = ved.ExamId.ToString();
                        td[9, i] = ved.Date.ToShortDateString(); ;

                        td.AddRow(1);
                        i++;
                    }

                    td.DeleteLastRow();
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка вывода в Word: \n" + exc.Message);
            }           
        }        

        //изменение - только для супер
        private void btnChange_Click(object sender, EventArgs e)
        {
            if (MainClass.RightsFacMain())
            {  
                ExamsVedCard p = new ExamsVedCard(this, ExamVedId);
                p.Show();
            }
        }
        
        //создать новый протокол
        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (MainClass.RightsFacMain())
            {
                SelectExamCrypto frm = new SelectExamCrypto(this, FacultyId, StudyBasisId);
                frm.Show();
            }                 
        }     

        //выбрать ведомость в списке
        public void SelectVed(Guid? vedId)
        {
            if (vedId != null)
                ExamVedId = vedId;                
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            if (ExamVedId == null)
                return;

            if (MainClass.IsOwner() || MainClass.IsCrypto() || MainClass.IsPasha())
            {
                if (MessageBox.Show("Ведомость будет закрыта для редактирования, продолжить? ", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        using (PriemEntities context = new PriemEntities())
                        {
                            using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                            {
                                List<int> lstNumbers = new List<int>();
                                int curNum;
                                foreach (DataGridViewRow dgvr in dgvList.Rows)
                                {
                                    Guid persId = new Guid(dgvr.Cells["Id"].Value.ToString());

                                    curNum = GetRandomNumber(ref lstNumbers);
                                    context.ExamsVedHistory_UpdateNumber(ExamVedId, persId, curNum);
                                }

                                context.ExamsVed_UpdateLock(true, ExamVedId);
                                
                                btnChange.Enabled = false;
                                btnDelete.Enabled = false;
                                btnLock.Enabled = false;
                                lblLocked.Visible = true;
                                MessageBox.Show("Выполнено");

                                btnPrintSticker.Enabled = true;
                                tbCountCell.Enabled = true;

                                transaction.Complete();
                            }
                        }
                    }

                    catch (Exception ex)
                    {
                        WinFormsServ.Error("Ошибка обновления данных: " + ex.Message);
                    }
                }
            }
            else
                WinFormsServ.Error("Невозможно закрытие ведомостей, недостаточно прав"); 
        }

        private int GetRandomNumber(ref List<int> lstNums)
        {
            int g;
            Random r = new Random();
            g = r.Next(10000);

            if (!lstNums.Contains(g))
            {
                lstNums.Add(g);
                return g;
            }
            else            
                return GetRandomNumber(ref lstNums);
        }

        
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (ExamVedId == null)
                return;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    int cnt = (from ev in context.extExamsVed
                               where ev.Id == ExamVedId && (ev.IsLocked || ev.IsLoad)
                               select ev).Count();
                    
                    if (cnt > 0)
                    {
                        WinFormsServ.Error("Данная ведомость уже закрыта. Удаление невозможно!");
                        return;
                    }

                    if (MainClass.IsPasha())
                    {
                        if (MessageBox.Show("Удалить выбранную ведомость? ", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                            {
                                context.ExamsVedHistory_DeleteByVedId(ExamVedId);
                                context.ExamsVed_Delete(ExamVedId);

                                transaction.Complete();

                                UpdateVedList();
                            }
                        }
                    }
                }
            }
            catch (Exception de)
            {
                WinFormsServ.Error("Ошибка удаления данных " + de.Message);
            }
             
        }

        private void btnCreateAdd_Click(object sender, EventArgs e)
        {
            if (MainClass.IsCrypto() || MainClass.IsOwner() || MainClass.IsPasha())
            {
                using (PriemEntities context = new PriemEntities())
                {
                    int? stBas = null;
                    if (cbExamVed.SelectedItem.ToString().Contains("г/б"))
                        stBas = 1;
                    else if (cbExamVed.SelectedItem.ToString().Contains("дог"))
                        stBas = 2;

                    extExamsVed ved = (from ev in context.extExamsVed
                                        where ev.Id == ExamVedId
                                        select ev).FirstOrDefault();

                    DateTime passDate = ved.Date;
                    int examId = ved.ExamId;

                    SelectExamCrypto frm = new SelectExamCrypto(this, FacultyId, stBas, passDate, examId);
                    frm.Show();
                }
            } 
        }

        private void dgvList_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == dgvList.Columns["Number"].Index)
            {
                e.Value = string.Format("{0}", e.RowIndex+1);
            }  
        }

        private void btnPrintSticker_Click(object sender, EventArgs e)
        {
            if (ExamVedId == null)
                return;

            if (MainClass.IsOwner() || MainClass.IsCrypto() || MainClass.IsPasha() || MainClass.IsCryptoMain())
            {
                FileStream fileS = null;                
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Sticker.pdf";
                                
                float fontsize = 8;
                
                try
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        Document document = new Document(PageSize.A4, 50, 50, 50, 50);
                        document.SetMargins(18, 18, 36, 5);

                        using (fileS = new FileStream(savePath, FileMode.Create))
                        {
                            BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                            iTextSharp.text.Font font = new iTextSharp.text.Font(bfTimes, 10);

                            PdfWriter pw = PdfWriter.GetInstance(document, fileS);
                            document.Open();

                            int cntCells = 2;
                            int.TryParse(tbCountCell.Text.Trim(), out cntCells);

                            string examId;
                            string examName;
                            string examDate;
                            string facId;
                            string vedNumber;

                            extExamsVed ved = (from ev in context.extExamsVed
                                               where ev.Id == ExamVedId
                                               select ev).FirstOrDefault();                           

                            examId = ved.ExamId.ToString();
                            examName = ved.ExamName.ToString();
                            if (examName.Length > 80)
                                examName = examName.Substring(0, 36) + "..." + examName.Substring(examName.Length - 36);
                            examDate = ved.Date.ToShortDateString();                           
                            vedNumber = ved.Number.ToString();
                            facId = FacultyId.ToString();
                            
                            DataSet dsPersons = bdc.GetDataSet(string.Format("SELECT DISTINCT ed.extPersonAspirant.Id, ed.extPersonAspirant.PersonNum as RegNum, ed.extPersonAspirant.FIO, ed.ExamsVedHistory.PersonVedNumber " +
                                          "FROM ed.extPersonAspirant LEFT JOIN ed.ExamsVedHistory ON ed.ExamsVedHistory.PersonId = ed.extPersonAspirant.Id WHERE ed.ExamsVedHistory.ExamsVedId = '{0}' ORDER BY FIO ", ExamVedId.ToString()));
                            
                            PdfPTable t = new PdfPTable(3);
                            float pgW = (PageSize.A4.Width - 36) / 3;
                            float[] headerwidths = { pgW, pgW, pgW };
                            t.SetWidths(headerwidths);
                            t.WidthPercentage = 100f;
                            t.SpacingBefore = 10f;
                            t.SpacingAfter = 10f;
                            t.DefaultCell.MinimumHeight = 120;

                            int cellsNum = (dsPersons.Tables[0].Rows.Count) * (cntCells + 1);
                            int ost = cellsNum % 3;

                            foreach (DataRow drr in dsPersons.Tables[0].Rows)
                            {
                                string text = drr["FIO"].ToString() + "\n" + drr["RegNum"].ToString() + " " + facId + "\n";
                                text += examDate + " " + examName;

                                Barcode128 barcode1 = new Barcode128();
                                barcode1.Code = vedNumber + "==" + drr["PersonVedNumber"].ToString() + "-";

                                Barcode128 barcode2 = new Barcode128();
                                barcode2.Code = vedNumber + "==" + drr["PersonVedNumber"].ToString() + "-";

                                PdfContentByte cb = pw.DirectContent;
                                iTextSharp.text.Image img1 = barcode1.CreateImageWithBarcode(cb, null, null);
                                img1.ScaleAbsolute(80f, 60f);

                                iTextSharp.text.Image img2 = barcode2.CreateImageWithBarcode(cb, iTextSharp.text.Color.BLACK, iTextSharp.text.Color.WHITE);
                                img2.ScaleAbsolute(80f, 60f);


                                PdfPTable ptPl = new PdfPTable(1);
                                float[] hwh = { pgW };
                                ptPl.SetWidthPercentage(hwh, PageSize.A4);

                                PdfPCell clPlText = new PdfPCell(new Phrase(text, new iTextSharp.text.Font(bfTimes, fontsize)));
                                clPlText.HorizontalAlignment = iTextSharp.text.Rectangle.ALIGN_CENTER;
                                clPlText.PaddingBottom = 2;
                                clPlText.PaddingTop = 2;
                                clPlText.Border = iTextSharp.text.Rectangle.NO_BORDER;

                                PdfPCell clPlBarc = new PdfPCell();
                                clPlBarc.AddElement(img1);
                                clPlBarc.HorizontalAlignment = iTextSharp.text.Rectangle.ALIGN_CENTER;
                                clPlBarc.PaddingTop = 1;
                                clPlBarc.PaddingLeft = 40;
                                clPlBarc.Border = iTextSharp.text.Rectangle.NO_BORDER;

                                ptPl.AddCell(clPlText);
                                ptPl.AddCell(clPlBarc);

                                PdfPCell pcell = new PdfPCell(ptPl);
                                pcell.PaddingTop = 6;
                                pcell.PaddingBottom = 6;
                                pcell.PaddingLeft = 6;
                                pcell.PaddingRight = 6;
                                pcell.FixedHeight = 100;
                                pcell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                                t.AddCell(pcell);

                                PdfPCell pcell1;
                                for (int i = 0; i < cntCells; i++)
                                {
                                    ptPl = new PdfPTable(1);
                                    ptPl.SetWidthPercentage(hwh, PageSize.A4);

                                    clPlText = new PdfPCell();
                                    clPlText.AddElement(img2);
                                    clPlText.PaddingLeft = 40;
                                    clPlText.PaddingRight = 40;
                                    clPlText.PaddingTop = 20;
                                    clPlText.Border = iTextSharp.text.Rectangle.NO_BORDER;

                                    clPlBarc = new PdfPCell(new Phrase((i + 1).ToString(), new iTextSharp.text.Font(bfTimes, fontsize)));
                                    clPlBarc.HorizontalAlignment = iTextSharp.text.Rectangle.ALIGN_CENTER;
                                    clPlBarc.PaddingTop = 1;
                                    clPlBarc.Border = iTextSharp.text.Rectangle.NO_BORDER;

                                    ptPl.AddCell(clPlText);
                                    ptPl.AddCell(clPlBarc);

                                    pcell1 = new PdfPCell(ptPl);
                                    pcell1.FixedHeight = 100;
                                    pcell1.Border = iTextSharp.text.Rectangle.NO_BORDER;
                                    t.AddCell(pcell1);
                                }
                            }

                            for (int i = 0; i < 3 - ost; i++)
                            {
                                PdfPCell pc = new PdfPCell();
                                pc.Border = iTextSharp.text.Rectangle.NO_BORDER;
                                t.AddCell(pc);
                            }

                            if (t != null)
                                document.Add(t);

                            document.Close();

                            Process pr = new Process();

                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();

                            pr.Close();
                        }
                    }
                }

                catch (Exception exc)
                {
                    WinFormsServ.Error(exc.Message);
                }
                finally
                {
                    if (fileS != null)
                        fileS.Dispose();

                }
            }

            else
                WinFormsServ.Error("Невозможно создание наклеек, недостаточно прав"); 
        }

        private void btnDeleteFromVed_Click(object sender, EventArgs e)
        {
            if (ExamVedId == null)
                return;

            if (MainClass.IsPasha())
            {
                if (MessageBox.Show("Удалить person из ведомости?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        foreach (DataGridViewRow dgvr in dgvList.SelectedRows)
                        {
                            Guid persId = new Guid(dgvr.Cells["Id"].Value.ToString());
                            try
                            {
                                context.ExamsVedHistory_DeleteByPersonAndVedId(ExamVedId, persId);                              
                            }
                            catch (Exception ex)
                            {
                                WinFormsServ.Error("Ошибка удаления данных" + ex.Message);
                                goto Next;
                            }
                        Next: ;
                        }
                        UpdateDataGrid();
                    }
                }
            }
        }

        private void btnUnload_Click(object sender, EventArgs e)
        {
            if (ExamVedId == null)
                return;
            using (PriemEntities context = new PriemEntities())
            {
                bool isLocked = (from ev in context.extExamsVed
                                 where ev.Id == ExamVedId
                                 select ev.IsLocked).FirstOrDefault();

                if (!isLocked)
                    return;

                if (MainClass.IsPasha())
                {
                    if (MessageBox.Show("Разлочить ведомость и удалить загруженные оценки?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                            {
                                context.Mark_DeleteByExamVedId(ExamVedId);
                                context.ExamsVed_UpdateLoad(false, ExamVedId);

                                transaction.Complete();
                            }
                           
                            MessageBox.Show("Выполнено");
                        }
                        catch (Exception ex)
                        {
                            WinFormsServ.Error("Ошибка удаления данных" + ex.Message);
                        }
                    }
                }
            }
        }         
    }
}