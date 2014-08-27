namespace Priem
{
    partial class Migrator
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Migrator));
            this.btnStart = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.tbFolder = new System.Windows.Forms.TextBox();
            this.lblPath = new System.Windows.Forms.Label();
            this.btnFolder = new System.Windows.Forms.Button();
            this.btnMetro = new System.Windows.Forms.Button();
            this.cbFaculty = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(12, 135);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(143, 23);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "база mdb для Студента";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 13);
            this.label2.TabIndex = 74;
            this.label2.Text = "Факультет";
            // 
            // tbFolder
            // 
            this.tbFolder.Location = new System.Drawing.Point(12, 95);
            this.tbFolder.Name = "tbFolder";
            this.tbFolder.ReadOnly = true;
            this.tbFolder.Size = new System.Drawing.Size(240, 20);
            this.tbFolder.TabIndex = 76;
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(9, 79);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(209, 13);
            this.lblPath.TabIndex = 77;
            this.lblPath.Text = "Куда сохранить базу в старом формате";
            // 
            // btnFolder
            // 
            this.btnFolder.Location = new System.Drawing.Point(258, 93);
            this.btnFolder.Name = "btnFolder";
            this.btnFolder.Size = new System.Drawing.Size(75, 23);
            this.btnFolder.TabIndex = 78;
            this.btnFolder.Text = "Выбрать";
            this.btnFolder.UseVisualStyleBackColor = true;
            this.btnFolder.Click += new System.EventHandler(this.btnFolder_Click);
            // 
            // btnMetro
            // 
            this.btnMetro.Location = new System.Drawing.Point(12, 164);
            this.btnMetro.Name = "btnMetro";
            this.btnMetro.Size = new System.Drawing.Size(143, 23);
            this.btnMetro.TabIndex = 80;
            this.btnMetro.Text = "база для Метро";
            this.btnMetro.UseVisualStyleBackColor = true;
            this.btnMetro.Click += new System.EventHandler(this.btnMetro_Click);
            // 
            // cbFaculty
            // 
            this.cbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFaculty.FormattingEnabled = true;
            this.cbFaculty.Location = new System.Drawing.Point(15, 28);
            this.cbFaculty.Name = "cbFaculty";
            this.cbFaculty.Size = new System.Drawing.Size(250, 21);
            this.cbFaculty.TabIndex = 81;
            // 
            // Migrator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(352, 209);
            this.Controls.Add(this.cbFaculty);
            this.Controls.Add(this.btnMetro);
            this.Controls.Add(this.btnFolder);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.tbFolder);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnStart);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Migrator";
            this.Text = "Мигратор данных";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowser;
        private System.Windows.Forms.TextBox tbFolder;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Button btnFolder;
        private System.Windows.Forms.Button btnMetro;
        private System.Windows.Forms.ComboBox cbFaculty;
    }
}

