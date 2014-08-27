namespace Priem
{
    partial class ExamsVedList
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExamsVedList));
            this.btnPrint = new System.Windows.Forms.Button();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnChange = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.dgvList = new System.Windows.Forms.DataGridView();
            this.lblProtocolNum = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnLock = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.btnDelete = new System.Windows.Forms.Button();
            this.lblLocked = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnCreateAdd = new System.Windows.Forms.Button();
            this.btnPrintSticker = new System.Windows.Forms.Button();
            this.tbCountCell = new System.Windows.Forms.TextBox();
            this.lblCountCell = new System.Windows.Forms.Label();
            this.btnDeleteFromVed = new System.Windows.Forms.Button();
            this.btnUnload = new System.Windows.Forms.Button();
            this.cbFaculty = new System.Windows.Forms.ComboBox();
            this.cbExamVed = new System.Windows.Forms.ComboBox();
            this.cbStudyBasis = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvList)).BeginInit();
            this.SuspendLayout();
            // 
            // btnPrint
            // 
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.Location = new System.Drawing.Point(904, 444);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(96, 42);
            this.btnPrint.TabIndex = 1;
            this.btnPrint.Text = "Печать ведомости";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCreate
            // 
            this.btnCreate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreate.Location = new System.Drawing.Point(929, 63);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(165, 23);
            this.btnCreate.TabIndex = 44;
            this.btnCreate.Text = "Новая ведомость";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Visible = false;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnChange
            // 
            this.btnChange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnChange.Enabled = false;
            this.btnChange.Location = new System.Drawing.Point(12, 440);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(87, 23);
            this.btnChange.TabIndex = 42;
            this.btnChange.Text = "Изменить";
            this.btnChange.UseVisualStyleBackColor = true;
            this.btnChange.Visible = false;
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(1019, 463);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 41;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // dgvList
            // 
            this.dgvList.AllowUserToAddRows = false;
            this.dgvList.AllowUserToDeleteRows = false;
            this.dgvList.AllowUserToOrderColumns = true;
            this.dgvList.AllowUserToResizeRows = false;
            this.dgvList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvList.Location = new System.Drawing.Point(12, 92);
            this.dgvList.MultiSelect = false;
            this.dgvList.Name = "dgvList";
            this.dgvList.ReadOnly = true;
            this.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvList.Size = new System.Drawing.Size(1082, 342);
            this.dgvList.TabIndex = 40;
            this.dgvList.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvList_CellFormatting);
            // 
            // lblProtocolNum
            // 
            this.lblProtocolNum.AutoSize = true;
            this.lblProtocolNum.Location = new System.Drawing.Point(12, 49);
            this.lblProtocolNum.Name = "lblProtocolNum";
            this.lblProtocolNum.Size = new System.Drawing.Size(63, 13);
            this.lblProtocolNum.TabIndex = 37;
            this.lblProtocolNum.Text = "Ведомость";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 36;
            this.label1.Text = "Факультет";
            // 
            // btnLock
            // 
            this.btnLock.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnLock.Location = new System.Drawing.Point(219, 444);
            this.btnLock.Name = "btnLock";
            this.btnLock.Size = new System.Drawing.Size(87, 42);
            this.btnLock.TabIndex = 45;
            this.btnLock.Text = "Закрыть ведомость";
            this.btnLock.UseVisualStyleBackColor = true;
            this.btnLock.Click += new System.EventHandler(this.btnLock_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(12, 469);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(87, 23);
            this.btnDelete.TabIndex = 46;
            this.btnDelete.Text = "Удалить";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Visible = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // lblLocked
            // 
            this.lblLocked.AutoSize = true;
            this.lblLocked.ForeColor = System.Drawing.Color.Red;
            this.lblLocked.Location = new System.Drawing.Point(439, 68);
            this.lblLocked.Name = "lblLocked";
            this.lblLocked.Size = new System.Drawing.Size(51, 13);
            this.lblLocked.TabIndex = 47;
            this.lblLocked.Text = "Закрыта";
            this.lblLocked.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(241, 5);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(94, 13);
            this.label3.TabIndex = 49;
            this.label3.Text = "Основа обучения";
            // 
            // btnCreateAdd
            // 
            this.btnCreateAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCreateAdd.Location = new System.Drawing.Point(122, 444);
            this.btnCreateAdd.Name = "btnCreateAdd";
            this.btnCreateAdd.Size = new System.Drawing.Size(91, 42);
            this.btnCreateAdd.TabIndex = 50;
            this.btnCreateAdd.Text = "Создать доп.ведомость";
            this.btnCreateAdd.UseVisualStyleBackColor = true;
            this.btnCreateAdd.Visible = false;
            this.btnCreateAdd.Click += new System.EventHandler(this.btnCreateAdd_Click);
            // 
            // btnPrintSticker
            // 
            this.btnPrintSticker.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnPrintSticker.Location = new System.Drawing.Point(728, 444);
            this.btnPrintSticker.Name = "btnPrintSticker";
            this.btnPrintSticker.Size = new System.Drawing.Size(81, 42);
            this.btnPrintSticker.TabIndex = 51;
            this.btnPrintSticker.Text = "Печать наклеек";
            this.btnPrintSticker.UseVisualStyleBackColor = true;
            this.btnPrintSticker.Click += new System.EventHandler(this.btnPrintSticker_Click);
            // 
            // tbCountCell
            // 
            this.tbCountCell.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tbCountCell.Location = new System.Drawing.Point(815, 465);
            this.tbCountCell.Name = "tbCountCell";
            this.tbCountCell.Size = new System.Drawing.Size(59, 20);
            this.tbCountCell.TabIndex = 52;
            // 
            // lblCountCell
            // 
            this.lblCountCell.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblCountCell.AutoSize = true;
            this.lblCountCell.Location = new System.Drawing.Point(812, 437);
            this.lblCountCell.Name = "lblCountCell";
            this.lblCountCell.Size = new System.Drawing.Size(67, 26);
            this.lblCountCell.TabIndex = 53;
            this.lblCountCell.Text = "Кол-во\r\nштрихкодов";
            // 
            // btnDeleteFromVed
            // 
            this.btnDeleteFromVed.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDeleteFromVed.Location = new System.Drawing.Point(595, 444);
            this.btnDeleteFromVed.Name = "btnDeleteFromVed";
            this.btnDeleteFromVed.Size = new System.Drawing.Size(91, 42);
            this.btnDeleteFromVed.TabIndex = 54;
            this.btnDeleteFromVed.Text = "Удалить из ведомости";
            this.btnDeleteFromVed.UseVisualStyleBackColor = true;
            this.btnDeleteFromVed.Visible = false;
            this.btnDeleteFromVed.Click += new System.EventHandler(this.btnDeleteFromVed_Click);
            // 
            // btnUnload
            // 
            this.btnUnload.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnUnload.Location = new System.Drawing.Point(394, 444);
            this.btnUnload.Name = "btnUnload";
            this.btnUnload.Size = new System.Drawing.Size(83, 42);
            this.btnUnload.TabIndex = 55;
            this.btnUnload.Text = "Разлочить ведомость";
            this.btnUnload.UseVisualStyleBackColor = true;
            this.btnUnload.Click += new System.EventHandler(this.btnUnload_Click);
            // 
            // cbFaculty
            // 
            this.cbFaculty.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbFaculty.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFaculty.FormattingEnabled = true;
            this.cbFaculty.Location = new System.Drawing.Point(12, 20);
            this.cbFaculty.Name = "cbFaculty";
            this.cbFaculty.Size = new System.Drawing.Size(226, 21);
            this.cbFaculty.TabIndex = 142;
            // 
            // cbExamVed
            // 
            this.cbExamVed.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbExamVed.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbExamVed.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbExamVed.FormattingEnabled = true;
            this.cbExamVed.Location = new System.Drawing.Point(12, 65);
            this.cbExamVed.Name = "cbExamVed";
            this.cbExamVed.Size = new System.Drawing.Size(421, 21);
            this.cbExamVed.TabIndex = 141;
            // 
            // cbStudyBasis
            // 
            this.cbStudyBasis.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbStudyBasis.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbStudyBasis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyBasis.FormattingEnabled = true;
            this.cbStudyBasis.Location = new System.Drawing.Point(244, 21);
            this.cbStudyBasis.Name = "cbStudyBasis";
            this.cbStudyBasis.Size = new System.Drawing.Size(136, 21);
            this.cbStudyBasis.TabIndex = 145;
            // 
            // ExamsVedList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1106, 498);
            this.Controls.Add(this.cbStudyBasis);
            this.Controls.Add(this.cbFaculty);
            this.Controls.Add(this.cbExamVed);
            this.Controls.Add(this.btnUnload);
            this.Controls.Add(this.btnDeleteFromVed);
            this.Controls.Add(this.lblCountCell);
            this.Controls.Add(this.tbCountCell);
            this.Controls.Add(this.btnPrintSticker);
            this.Controls.Add(this.btnCreateAdd);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblLocked);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnLock);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.btnChange);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.dgvList);
            this.Controls.Add(this.lblProtocolNum);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ExamsVedList";
            this.Text = "Список ведомостей для экзаменов";
            ((System.ComponentModel.ISupportInitialize)(this.dgvList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        protected System.Windows.Forms.Button btnPrint;
        protected System.Windows.Forms.Button btnCreate;
        protected System.Windows.Forms.Button btnChange;
        protected System.Windows.Forms.Button btnClose;
        protected System.Windows.Forms.DataGridView dgvList;       
        protected System.Windows.Forms.Label lblProtocolNum;       
        protected System.Windows.Forms.Label label1;
        protected System.Windows.Forms.Button btnLock;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        protected System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Label lblLocked;       
        private System.Windows.Forms.Label label3;
        protected System.Windows.Forms.Button btnCreateAdd;
        protected System.Windows.Forms.Button btnPrintSticker;
        private System.Windows.Forms.TextBox tbCountCell;
        private System.Windows.Forms.Label lblCountCell;
        protected System.Windows.Forms.Button btnDeleteFromVed;
        protected System.Windows.Forms.Button btnUnload;
        private System.Windows.Forms.ComboBox cbFaculty;
        private System.Windows.Forms.ComboBox cbExamVed;
        private System.Windows.Forms.ComboBox cbStudyBasis;
    }
}