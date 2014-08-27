namespace Priem
{
    partial class ProtocolList
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
            this.cbFaculty = new System.Windows.Forms.ComboBox();
            this.cbStudyForm = new System.Windows.Forms.ComboBox();
            this.cbStudyBasis = new System.Windows.Forms.ComboBox();
            this.cbProtocolNum = new System.Windows.Forms.ComboBox();
            this.comboBox5 = new System.Windows.Forms.ComboBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.btnPrint = new System.Windows.Forms.Button();
            this.cbPrint = new System.Windows.Forms.ComboBox();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnMake = new System.Windows.Forms.Button();
            this.btnCard = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.dgvProtocols = new System.Windows.Forms.DataGridView();
            this.lblProtocolNum = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblBase = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProtocols)).BeginInit();
            this.SuspendLayout();
            // 
            // cbFaculty
            // 
            this.cbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFaculty.FormattingEnabled = true;
            this.cbFaculty.Location = new System.Drawing.Point(12, 24);
            this.cbFaculty.Name = "cbFaculty";
            this.cbFaculty.Size = new System.Drawing.Size(177, 21);
            this.cbFaculty.TabIndex = 0;
            // 
            // cbStudyForm
            // 
            this.cbStudyForm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyForm.FormattingEnabled = true;
            this.cbStudyForm.Location = new System.Drawing.Point(195, 24);
            this.cbStudyForm.Name = "cbStudyForm";
            this.cbStudyForm.Size = new System.Drawing.Size(121, 21);
            this.cbStudyForm.TabIndex = 1;
            this.cbStudyForm.SelectedIndexChanged += new System.EventHandler(this.UpdateProtocolList);
            // 
            // cbStudyBasis
            // 
            this.cbStudyBasis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyBasis.FormattingEnabled = true;
            this.cbStudyBasis.Location = new System.Drawing.Point(339, 24);
            this.cbStudyBasis.Name = "cbStudyBasis";
            this.cbStudyBasis.Size = new System.Drawing.Size(121, 21);
            this.cbStudyBasis.TabIndex = 2;
            this.cbStudyBasis.SelectedIndexChanged += new System.EventHandler(this.UpdateProtocolList);
            // 
            // cbProtocolNum
            // 
            this.cbProtocolNum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbProtocolNum.FormattingEnabled = true;
            this.cbProtocolNum.Location = new System.Drawing.Point(15, 66);
            this.cbProtocolNum.Name = "cbProtocolNum";
            this.cbProtocolNum.Size = new System.Drawing.Size(121, 21);
            this.cbProtocolNum.TabIndex = 3;
            // 
            // comboBox5
            // 
            this.comboBox5.FormattingEnabled = true;
            this.comboBox5.Location = new System.Drawing.Point(254, 228);
            this.comboBox5.Name = "comboBox5";
            this.comboBox5.Size = new System.Drawing.Size(121, 21);
            this.comboBox5.TabIndex = 4;
            // 
            // groupBox9
            // 
            this.groupBox9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox9.Controls.Add(this.btnPrint);
            this.groupBox9.Controls.Add(this.cbPrint);
            this.groupBox9.Location = new System.Drawing.Point(371, 466);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(242, 68);
            this.groupBox9.TabIndex = 45;
            this.groupBox9.TabStop = false;
            // 
            // btnPrint
            // 
            this.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnPrint.Location = new System.Drawing.Point(6, 12);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(117, 23);
            this.btnPrint.TabIndex = 1;
            this.btnPrint.Text = "Печать протокола";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // cbPrint
            // 
            this.cbPrint.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.cbPrint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbPrint.FormattingEnabled = true;
            this.cbPrint.Location = new System.Drawing.Point(6, 41);
            this.cbPrint.Name = "cbPrint";
            this.cbPrint.Size = new System.Drawing.Size(230, 21);
            this.cbPrint.TabIndex = 0;
            // 
            // btnCreate
            // 
            this.btnCreate.Location = new System.Drawing.Point(153, 66);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(99, 23);
            this.btnCreate.TabIndex = 44;
            this.btnCreate.Text = "Новый протокол";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Visible = false;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnMake
            // 
            this.btnMake.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnMake.Enabled = false;
            this.btnMake.Location = new System.Drawing.Point(94, 511);
            this.btnMake.Name = "btnMake";
            this.btnMake.Size = new System.Drawing.Size(95, 23);
            this.btnMake.TabIndex = 42;
            this.btnMake.Text = "Изменить";
            this.btnMake.UseVisualStyleBackColor = true;
            this.btnMake.Visible = false;
            this.btnMake.Click += new System.EventHandler(this.btnMake_Click);
            // 
            // btnCard
            // 
            this.btnCard.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCard.Location = new System.Drawing.Point(13, 511);
            this.btnCard.Name = "btnCard";
            this.btnCard.Size = new System.Drawing.Size(75, 23);
            this.btnCard.TabIndex = 43;
            this.btnCard.Text = "Карточка";
            this.btnCard.UseVisualStyleBackColor = true;
            this.btnCard.Click += new System.EventHandler(this.btnCard_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(619, 511);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 41;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // dgvProtocols
            // 
            this.dgvProtocols.AllowUserToAddRows = false;
            this.dgvProtocols.AllowUserToDeleteRows = false;
            this.dgvProtocols.AllowUserToOrderColumns = true;
            this.dgvProtocols.AllowUserToResizeRows = false;
            this.dgvProtocols.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvProtocols.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvProtocols.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvProtocols.Location = new System.Drawing.Point(16, 95);
            this.dgvProtocols.MultiSelect = false;
            this.dgvProtocols.Name = "dgvProtocols";
            this.dgvProtocols.ReadOnly = true;
            this.dgvProtocols.Size = new System.Drawing.Size(681, 335);
            this.dgvProtocols.TabIndex = 40;
            // 
            // lblProtocolNum
            // 
            this.lblProtocolNum.AutoSize = true;
            this.lblProtocolNum.Location = new System.Drawing.Point(13, 52);
            this.lblProtocolNum.Name = "lblProtocolNum";
            this.lblProtocolNum.Size = new System.Drawing.Size(97, 13);
            this.lblProtocolNum.TabIndex = 38;
            this.lblProtocolNum.Text = "Номер протокола";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(193, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 39;
            this.label3.Text = "Форма обучения";
            // 
            // lblBase
            // 
            this.lblBase.AutoSize = true;
            this.lblBase.Location = new System.Drawing.Point(340, 10);
            this.lblBase.Name = "lblBase";
            this.lblBase.Size = new System.Drawing.Size(94, 13);
            this.lblBase.TabIndex = 36;
            this.lblBase.Text = "Основа обучения";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 37;
            this.label1.Text = "Факультет";
            // 
            // ProtocolList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(706, 544);
            this.Controls.Add(this.groupBox9);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.btnMake);
            this.Controls.Add(this.btnCard);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.dgvProtocols);
            this.Controls.Add(this.lblProtocolNum);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblBase);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox5);
            this.Controls.Add(this.cbProtocolNum);
            this.Controls.Add(this.cbStudyBasis);
            this.Controls.Add(this.cbStudyForm);
            this.Controls.Add(this.cbFaculty);
            this.Name = "ProtocolList";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "ProtocolList";
            this.groupBox9.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvProtocols)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbFaculty;
        private System.Windows.Forms.ComboBox cbStudyForm;
        private System.Windows.Forms.ComboBox cbStudyBasis;
        private System.Windows.Forms.ComboBox comboBox5;
        protected System.Windows.Forms.GroupBox groupBox9;
        protected System.Windows.Forms.Button btnPrint;
        protected System.Windows.Forms.ComboBox cbPrint;
        protected System.Windows.Forms.Button btnCreate;
        protected System.Windows.Forms.Button btnMake;
        protected System.Windows.Forms.Button btnCard;
        protected System.Windows.Forms.Button btnClose;
        protected System.Windows.Forms.DataGridView dgvProtocols;
        protected System.Windows.Forms.Label lblProtocolNum;
        protected System.Windows.Forms.Label label3;
        protected System.Windows.Forms.Label lblBase;
        protected System.Windows.Forms.Label label1;
        protected System.Windows.Forms.ComboBox cbProtocolNum;
    }
}