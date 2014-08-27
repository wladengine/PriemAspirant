namespace Priem
{
    partial class Olymp2Competition
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Olymp2Competition));
            this.btnOk = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.lblSearch = new System.Windows.Forms.Label();
            this.tbFIO = new System.Windows.Forms.TextBox();
            this.chbAll = new System.Windows.Forms.CheckBox();
            this.btnPrint = new System.Windows.Forms.Button();
            this.cbOlympName = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cbOlympSubject = new System.Windows.Forms.ComboBox();
            this.cbOlympValue = new System.Windows.Forms.ComboBox();
            this.cbOlympLevel = new System.Windows.Forms.ComboBox();
            this.cbOlympType = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cbStudyBasis = new System.Windows.Forms.ComboBox();
            this.cbStudyForm = new System.Windows.Forms.ComboBox();
            this.cbLicenseProgram = new System.Windows.Forms.ComboBox();
            this.cbObrazProgram = new System.Windows.Forms.ComboBox();
            this.cbFaculty = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.lblSpeciality = new System.Windows.Forms.Label();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dgvAbitList = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAbitList)).BeginInit();
            this.SuspendLayout();
            // 
            // lblCount
            // 
            this.lblCount.Location = new System.Drawing.Point(361, 501);
            // 
            // btnCard
            // 
            this.btnCard.Location = new System.Drawing.Point(846, 530);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(550, 530);
            this.btnRemove.Visible = false;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(550, 501);
            this.btnAdd.Visible = false;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(928, 530);
            this.btnClose.Size = new System.Drawing.Size(87, 23);
            // 
            // btnOk
            // 
            this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOk.Location = new System.Drawing.Point(928, 501);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(87, 23);
            this.btnOk.TabIndex = 63;
            this.btnOk.Text = "Подтвердить";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(610, 156);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(34, 13);
            this.label4.TabIndex = 56;
            this.label4.Text = "ФИО";
            // 
            // lblSearch
            // 
            this.lblSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSearch.AutoSize = true;
            this.lblSearch.Location = new System.Drawing.Point(567, 175);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Size = new System.Drawing.Size(39, 13);
            this.lblSearch.TabIndex = 52;
            this.lblSearch.Text = "Поиск";
            // 
            // tbFIO
            // 
            this.tbFIO.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.tbFIO.Location = new System.Drawing.Point(612, 172);
            this.tbFIO.Name = "tbFIO";
            this.tbFIO.Size = new System.Drawing.Size(224, 20);
            this.tbFIO.TabIndex = 47;
            this.tbFIO.TextChanged += new System.EventHandler(this.tbFIO_TextChanged);
            // 
            // chbAll
            // 
            this.chbAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.chbAll.AutoSize = true;
            this.chbAll.Location = new System.Drawing.Point(914, 174);
            this.chbAll.Name = "chbAll";
            this.chbAll.Size = new System.Drawing.Size(101, 17);
            this.chbAll.TabIndex = 74;
            this.chbAll.Text = "Отметить всех";
            this.chbAll.UseVisualStyleBackColor = true;
            this.chbAll.CheckedChanged += new System.EventHandler(this.chbAll_CheckedChanged);
            // 
            // btnPrint
            // 
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnPrint.Location = new System.Drawing.Point(13, 509);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(281, 44);
            this.btnPrint.TabIndex = 75;
            this.btnPrint.Text = "Ведомость абитуриентов с измененным типом конкурса";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // cbOlympName
            // 
            this.cbOlympName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbOlympName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbOlympName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOlympName.FormattingEnabled = true;
            this.cbOlympName.Location = new System.Drawing.Point(195, 112);
            this.cbOlympName.Name = "cbOlympName";
            this.cbOlympName.Size = new System.Drawing.Size(320, 21);
            this.cbOlympName.TabIndex = 178;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(192, 96);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(57, 13);
            this.label10.TabIndex = 177;
            this.label10.Text = "Название";
            // 
            // cbOlympSubject
            // 
            this.cbOlympSubject.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbOlympSubject.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbOlympSubject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOlympSubject.FormattingEnabled = true;
            this.cbOlympSubject.Location = new System.Drawing.Point(521, 112);
            this.cbOlympSubject.Name = "cbOlympSubject";
            this.cbOlympSubject.Size = new System.Drawing.Size(271, 21);
            this.cbOlympSubject.TabIndex = 176;
            // 
            // cbOlympValue
            // 
            this.cbOlympValue.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbOlympValue.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbOlympValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOlympValue.FormattingEnabled = true;
            this.cbOlympValue.Location = new System.Drawing.Point(13, 155);
            this.cbOlympValue.Name = "cbOlympValue";
            this.cbOlympValue.Size = new System.Drawing.Size(117, 21);
            this.cbOlympValue.TabIndex = 175;
            // 
            // cbOlympLevel
            // 
            this.cbOlympLevel.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbOlympLevel.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbOlympLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOlympLevel.FormattingEnabled = true;
            this.cbOlympLevel.Location = new System.Drawing.Point(798, 113);
            this.cbOlympLevel.Name = "cbOlympLevel";
            this.cbOlympLevel.Size = new System.Drawing.Size(137, 21);
            this.cbOlympLevel.TabIndex = 174;
            // 
            // cbOlympType
            // 
            this.cbOlympType.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbOlympType.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbOlympType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOlympType.FormattingEnabled = true;
            this.cbOlympType.Location = new System.Drawing.Point(13, 113);
            this.cbOlympType.Name = "cbOlympType";
            this.cbOlympType.Size = new System.Drawing.Size(176, 21);
            this.cbOlympType.TabIndex = 173;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(796, 97);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 172;
            this.label2.Text = "Уровень";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(13, 139);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(96, 13);
            this.label7.TabIndex = 170;
            this.label7.Text = "Степень диплома";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 95);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(26, 13);
            this.label6.TabIndex = 171;
            this.label6.Text = "Вид";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(518, 96);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(52, 13);
            this.label5.TabIndex = 169;
            this.label5.Text = "Предмет";
            // 
            // cbStudyBasis
            // 
            this.cbStudyBasis.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbStudyBasis.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbStudyBasis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyBasis.FormattingEnabled = true;
            this.cbStudyBasis.Location = new System.Drawing.Point(424, 23);
            this.cbStudyBasis.Name = "cbStudyBasis";
            this.cbStudyBasis.Size = new System.Drawing.Size(136, 21);
            this.cbStudyBasis.TabIndex = 168;
            // 
            // cbStudyForm
            // 
            this.cbStudyForm.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbStudyForm.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbStudyForm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyForm.FormattingEnabled = true;
            this.cbStudyForm.Location = new System.Drawing.Point(260, 23);
            this.cbStudyForm.Name = "cbStudyForm";
            this.cbStudyForm.Size = new System.Drawing.Size(136, 21);
            this.cbStudyForm.TabIndex = 167;
            // 
            // cbLicenseProgram
            // 
            this.cbLicenseProgram.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbLicenseProgram.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbLicenseProgram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLicenseProgram.FormattingEnabled = true;
            this.cbLicenseProgram.Location = new System.Drawing.Point(13, 69);
            this.cbLicenseProgram.Name = "cbLicenseProgram";
            this.cbLicenseProgram.Size = new System.Drawing.Size(260, 21);
            this.cbLicenseProgram.TabIndex = 166;
            // 
            // cbObrazProgram
            // 
            this.cbObrazProgram.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbObrazProgram.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbObrazProgram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbObrazProgram.FormattingEnabled = true;
            this.cbObrazProgram.Location = new System.Drawing.Point(279, 69);
            this.cbObrazProgram.Name = "cbObrazProgram";
            this.cbObrazProgram.Size = new System.Drawing.Size(366, 21);
            this.cbObrazProgram.TabIndex = 165;
            // 
            // cbFaculty
            // 
            this.cbFaculty.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbFaculty.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFaculty.FormattingEnabled = true;
            this.cbFaculty.Location = new System.Drawing.Point(13, 25);
            this.cbFaculty.Name = "cbFaculty";
            this.cbFaculty.Size = new System.Drawing.Size(232, 21);
            this.cbFaculty.TabIndex = 164;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(276, 53);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(158, 13);
            this.label9.TabIndex = 163;
            this.label9.Text = "Образовательная программа";
            // 
            // lblSpeciality
            // 
            this.lblSpeciality.AutoSize = true;
            this.lblSpeciality.Location = new System.Drawing.Point(11, 51);
            this.lblSpeciality.Name = "lblSpeciality";
            this.lblSpeciality.Size = new System.Drawing.Size(75, 13);
            this.lblSpeciality.TabIndex = 162;
            this.lblSpeciality.Text = "Направление";
            // 
            // lblLanguage
            // 
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Location = new System.Drawing.Point(421, 8);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(94, 13);
            this.lblLanguage.TabIndex = 161;
            this.lblLanguage.Text = "Основа обучения";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(257, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 160;
            this.label3.Text = "Форма обучения";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 159;
            this.label1.Text = "Факультет";
            // 
            // dgvAbitList
            // 
            this.dgvAbitList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAbitList.Location = new System.Drawing.Point(13, 198);
            this.dgvAbitList.Name = "dgvAbitList";
            this.dgvAbitList.Size = new System.Drawing.Size(1002, 297);
            this.dgvAbitList.TabIndex = 179;
            // 
            // Olymp2Competition
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1026, 563);
            this.Controls.Add(this.dgvAbitList);
            this.Controls.Add(this.cbOlympName);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.cbOlympSubject);
            this.Controls.Add(this.cbOlympValue);
            this.Controls.Add(this.cbOlympLevel);
            this.Controls.Add(this.cbOlympType);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cbStudyBasis);
            this.Controls.Add(this.cbStudyForm);
            this.Controls.Add(this.cbLicenseProgram);
            this.Controls.Add(this.cbObrazProgram);
            this.Controls.Add(this.cbFaculty);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.lblSpeciality);
            this.Controls.Add(this.lblLanguage);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.chbAll);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lblSearch);
            this.Controls.Add(this.tbFIO);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Olymp2Competition";
            this.Text = "Зачет результатов олимпиад как изменение типа конкурса на б/э";
            this.Controls.SetChildIndex(this.tbFIO, 0);
            this.Controls.SetChildIndex(this.lblSearch, 0);
            this.Controls.SetChildIndex(this.btnOk, 0);
            this.Controls.SetChildIndex(this.label4, 0);
            this.Controls.SetChildIndex(this.btnRemove, 0);
            this.Controls.SetChildIndex(this.btnClose, 0);
            this.Controls.SetChildIndex(this.chbAll, 0);
            this.Controls.SetChildIndex(this.btnAdd, 0);
            this.Controls.SetChildIndex(this.btnCard, 0);
            this.Controls.SetChildIndex(this.btnPrint, 0);
            this.Controls.SetChildIndex(this.lblCount, 0);
            this.Controls.SetChildIndex(this.label1, 0);
            this.Controls.SetChildIndex(this.label3, 0);
            this.Controls.SetChildIndex(this.lblLanguage, 0);
            this.Controls.SetChildIndex(this.lblSpeciality, 0);
            this.Controls.SetChildIndex(this.label9, 0);
            this.Controls.SetChildIndex(this.cbFaculty, 0);
            this.Controls.SetChildIndex(this.cbObrazProgram, 0);
            this.Controls.SetChildIndex(this.cbLicenseProgram, 0);
            this.Controls.SetChildIndex(this.cbStudyForm, 0);
            this.Controls.SetChildIndex(this.cbStudyBasis, 0);
            this.Controls.SetChildIndex(this.label5, 0);
            this.Controls.SetChildIndex(this.label6, 0);
            this.Controls.SetChildIndex(this.label7, 0);
            this.Controls.SetChildIndex(this.label2, 0);
            this.Controls.SetChildIndex(this.cbOlympType, 0);
            this.Controls.SetChildIndex(this.cbOlympLevel, 0);
            this.Controls.SetChildIndex(this.cbOlympValue, 0);
            this.Controls.SetChildIndex(this.cbOlympSubject, 0);
            this.Controls.SetChildIndex(this.label10, 0);
            this.Controls.SetChildIndex(this.cbOlympName, 0);
            this.Controls.SetChildIndex(this.dgvAbitList, 0);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAbitList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblSearch;
        private System.Windows.Forms.TextBox tbFIO;
        private System.Windows.Forms.CheckBox chbAll;
        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.ComboBox cbOlympName;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cbOlympSubject;
        private System.Windows.Forms.ComboBox cbOlympValue;
        private System.Windows.Forms.ComboBox cbOlympLevel;
        private System.Windows.Forms.ComboBox cbOlympType;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbStudyBasis;
        private System.Windows.Forms.ComboBox cbStudyForm;
        private System.Windows.Forms.ComboBox cbLicenseProgram;
        private System.Windows.Forms.ComboBox cbObrazProgram;
        private System.Windows.Forms.ComboBox cbFaculty;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblSpeciality;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgvAbitList;
       
    }
}