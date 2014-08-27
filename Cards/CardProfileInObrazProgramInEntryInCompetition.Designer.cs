namespace Priem
{
    partial class CardProfileInObrazProgramInEntryInCompetition
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
            this.dgvObrazProgramInEntryList = new System.Windows.Forms.DataGridView();
            this.tbLicenseProgramName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbObrazProgramName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvObrazProgramInEntryList)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvObrazProgramInEntryList
            // 
            this.dgvObrazProgramInEntryList.AllowUserToAddRows = false;
            this.dgvObrazProgramInEntryList.AllowUserToDeleteRows = false;
            this.dgvObrazProgramInEntryList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvObrazProgramInEntryList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvObrazProgramInEntryList.Location = new System.Drawing.Point(12, 64);
            this.dgvObrazProgramInEntryList.Name = "dgvObrazProgramInEntryList";
            this.dgvObrazProgramInEntryList.ReadOnly = true;
            this.dgvObrazProgramInEntryList.Size = new System.Drawing.Size(543, 251);
            this.dgvObrazProgramInEntryList.TabIndex = 5;
            // 
            // tbLicenseProgramName
            // 
            this.tbLicenseProgramName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbLicenseProgramName.Location = new System.Drawing.Point(176, 12);
            this.tbLicenseProgramName.Name = "tbLicenseProgramName";
            this.tbLicenseProgramName.ReadOnly = true;
            this.tbLicenseProgramName.Size = new System.Drawing.Size(379, 20);
            this.tbLicenseProgramName.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(95, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Направление";
            // 
            // tbObrazProgramName
            // 
            this.tbObrazProgramName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbObrazProgramName.Location = new System.Drawing.Point(176, 38);
            this.tbObrazProgramName.Name = "tbObrazProgramName";
            this.tbObrazProgramName.ReadOnly = true;
            this.tbObrazProgramName.Size = new System.Drawing.Size(379, 20);
            this.tbObrazProgramName.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(158, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Образовательная программа";
            // 
            // CardProfileInObrazProgramInEntryInCompetition
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(567, 327);
            this.Controls.Add(this.tbObrazProgramName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dgvObrazProgramInEntryList);
            this.Controls.Add(this.tbLicenseProgramName);
            this.Controls.Add(this.label1);
            this.Name = "CardProfileInObrazProgramInEntryInCompetition";
            this.Text = "Приоритеты профилей в образовательной программе";
            ((System.ComponentModel.ISupportInitialize)(this.dgvObrazProgramInEntryList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvObrazProgramInEntryList;
        private System.Windows.Forms.TextBox tbLicenseProgramName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbObrazProgramName;
        private System.Windows.Forms.Label label2;
    }
}