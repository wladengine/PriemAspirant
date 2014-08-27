namespace Priem
{
    partial class LoadEgeMarks
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LoadEgeMarks));
            this.btnOk = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.dtDateExam = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.label36 = new System.Windows.Forms.Label();
            this.cbFaculty = new System.Windows.Forms.ComboBox();
            this.cbExam = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(113, 142);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 52;
            this.btnOk.Text = "ОК";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(14, 112);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(33, 13);
            this.label6.TabIndex = 51;
            this.label6.Text = "Дата";
            // 
            // dtDateExam
            // 
            this.dtDateExam.Location = new System.Drawing.Point(53, 106);
            this.dtDateExam.Name = "dtDateExam";
            this.dtDateExam.Size = new System.Drawing.Size(152, 20);
            this.dtDateExam.TabIndex = 50;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(14, 49);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(52, 13);
            this.label8.TabIndex = 49;
            this.label8.Text = "Экзамен";
            // 
            // label36
            // 
            this.label36.AutoSize = true;
            this.label36.Location = new System.Drawing.Point(12, 9);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(63, 13);
            this.label36.TabIndex = 46;
            this.label36.Text = "Факультет";
            // 
            // cbFaculty
            // 
            this.cbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFaculty.FormattingEnabled = true;
            this.cbFaculty.Location = new System.Drawing.Point(17, 25);
            this.cbFaculty.Name = "cbFaculty";
            this.cbFaculty.Size = new System.Drawing.Size(250, 21);
            this.cbFaculty.TabIndex = 60;
            // 
            // cbExam
            // 
            this.cbExam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbExam.FormattingEnabled = true;
            this.cbExam.Location = new System.Drawing.Point(17, 65);
            this.cbExam.Name = "cbExam";
            this.cbExam.Size = new System.Drawing.Size(250, 21);
            this.cbExam.TabIndex = 61;
            // 
            // LoadEgeMarks
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(315, 179);
            this.Controls.Add(this.cbExam);
            this.Controls.Add(this.cbFaculty);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.dtDateExam);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label36);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "LoadEgeMarks";
            this.Text = "Зачет ЕГЭ";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker dtDateExam;
        private System.Windows.Forms.Label label8;        
        private System.Windows.Forms.Label label36;
        private System.Windows.Forms.ComboBox cbFaculty;
        private System.Windows.Forms.ComboBox cbExam;
    }
}