namespace Priem
{
    partial class Filters
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
            this.components = new System.ComponentModel.Container();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.lbFilters = new System.Windows.Forms.ListBox();
            this.btnBrackets = new System.Windows.Forms.Button();
            this.btnOr = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnChange = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.gbFilter = new System.Windows.Forms.GroupBox();
            this.fFromTo = new EducServLib.FilterFromTo();
            this.fMult = new EducServLib.FilterMultySelect();
            this.fBool = new EducServLib.FilterBool();
            this.fDateFromTo = new EducServLib.FilterDateFromTo();
            this.fText = new System.Windows.Forms.TextBox();
            this.cmbFilters = new System.Windows.Forms.ComboBox();
            this.epError = new System.Windows.Forms.ErrorProvider(this.components);
            this.gbFilter.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.epError)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(7, 159);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 14;
            this.btnClear.Text = "Очистить";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnDown
            // 
            this.btnDown.Location = new System.Drawing.Point(439, 69);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(30, 31);
            this.btnDown.TabIndex = 12;
            this.btnDown.Text = "\\/";
            this.btnDown.UseVisualStyleBackColor = true;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.Location = new System.Drawing.Point(439, 32);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(30, 31);
            this.btnUp.TabIndex = 13;
            this.btnUp.Text = "/\\";
            this.btnUp.UseVisualStyleBackColor = true;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // lbFilters
            // 
            this.lbFilters.FormattingEnabled = true;
            this.lbFilters.Location = new System.Drawing.Point(7, 3);
            this.lbFilters.Name = "lbFilters";
            this.lbFilters.Size = new System.Drawing.Size(416, 147);
            this.lbFilters.TabIndex = 11;
            this.lbFilters.SelectedIndexChanged += new System.EventHandler(this.lbFilters_SelectedIndexChanged);
            // 
            // btnBrackets
            // 
            this.btnBrackets.Location = new System.Drawing.Point(322, 465);
            this.btnBrackets.Name = "btnBrackets";
            this.btnBrackets.Size = new System.Drawing.Size(52, 36);
            this.btnBrackets.TabIndex = 10;
            this.btnBrackets.Text = "Скобки";
            this.btnBrackets.UseVisualStyleBackColor = true;
            this.btnBrackets.Click += new System.EventHandler(this.btnBrackets_Click);
            // 
            // btnOr
            // 
            this.btnOr.Location = new System.Drawing.Point(265, 465);
            this.btnOr.Name = "btnOr";
            this.btnOr.Size = new System.Drawing.Size(51, 36);
            this.btnOr.TabIndex = 6;
            this.btnOr.Text = "ИЛИ";
            this.btnOr.UseVisualStyleBackColor = true;
            this.btnOr.Click += new System.EventHandler(this.btnOr_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(169, 465);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(61, 36);
            this.btnRemove.TabIndex = 5;
            this.btnRemove.Text = "Удалить";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(414, 465);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(55, 36);
            this.btnExit.TabIndex = 8;
            this.btnExit.Text = "ОК";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnChange
            // 
            this.btnChange.Enabled = false;
            this.btnChange.Location = new System.Drawing.Point(88, 465);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(75, 36);
            this.btnChange.TabIndex = 7;
            this.btnChange.Text = "Изменить текущий";
            this.btnChange.UseVisualStyleBackColor = true;
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(7, 465);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 36);
            this.btnAdd.TabIndex = 9;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // gbFilter
            // 
            this.gbFilter.Controls.Add(this.fFromTo);
            this.gbFilter.Controls.Add(this.fMult);
            this.gbFilter.Controls.Add(this.fBool);
            this.gbFilter.Controls.Add(this.fDateFromTo);
            this.gbFilter.Controls.Add(this.fText);
            this.gbFilter.Controls.Add(this.cmbFilters);
            this.gbFilter.Location = new System.Drawing.Point(7, 188);
            this.gbFilter.Name = "gbFilter";
            this.gbFilter.Size = new System.Drawing.Size(462, 269);
            this.gbFilter.TabIndex = 15;
            this.gbFilter.TabStop = false;
            // 
            // fFromTo
            // 
            this.fFromTo.FromValue = "";
            this.fFromTo.Location = new System.Drawing.Point(43, 117);
            this.fFromTo.Name = "fFromTo";
            this.fFromTo.Size = new System.Drawing.Size(384, 22);
            this.fFromTo.TabIndex = 9;
            this.fFromTo.ToValue = "";
            // 
            // fMult
            // 
            this.fMult.Location = new System.Drawing.Point(25, 36);
            this.fMult.Name = "fMult";
            this.fMult.Size = new System.Drawing.Size(412, 206);
            this.fMult.TabIndex = 6;
            // 
            // fBool
            // 
            this.fBool.Location = new System.Drawing.Point(162, 117);
            this.fBool.Name = "fBool";
            this.fBool.Size = new System.Drawing.Size(124, 23);
            this.fBool.TabIndex = 8;
            this.fBool.Value = true;
            // 
            // fDateFromTo
            // 
            this.fDateFromTo.FromValue = new System.DateTime(2012, 4, 12, 0, 0, 0, 0);
            this.fDateFromTo.Location = new System.Drawing.Point(42, 118);
            this.fDateFromTo.Name = "fDateFromTo";
            this.fDateFromTo.Size = new System.Drawing.Size(385, 22);
            this.fDateFromTo.TabIndex = 7;
            this.fDateFromTo.ToValue = new System.DateTime(2012, 4, 12, 0, 0, 0, 0);
            // 
            // fText
            // 
            this.fText.Location = new System.Drawing.Point(122, 120);
            this.fText.Name = "fText";
            this.fText.Size = new System.Drawing.Size(212, 20);
            this.fText.TabIndex = 5;
            this.fText.Visible = false;
            // 
            // cmbFilters
            // 
            this.cmbFilters.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFilters.FormattingEnabled = true;
            this.cmbFilters.Location = new System.Drawing.Point(91, 9);
            this.cmbFilters.Name = "cmbFilters";
            this.cmbFilters.Size = new System.Drawing.Size(267, 21);
            this.cmbFilters.Sorted = true;
            this.cmbFilters.TabIndex = 0;
            this.cmbFilters.SelectedIndexChanged += new System.EventHandler(this.cmbFilters_SelectedIndexChanged);
            // 
            // epError
            // 
            this.epError.ContainerControl = this;
            // 
            // Filters
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(476, 505);
            this.Controls.Add(this.gbFilter);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnDown);
            this.Controls.Add(this.btnUp);
            this.Controls.Add(this.lbFilters);
            this.Controls.Add(this.btnBrackets);
            this.Controls.Add(this.btnOr);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnChange);
            this.Controls.Add(this.btnAdd);
            this.Name = "Filters";
            this.Text = "Фильтры";
            this.gbFilter.ResumeLayout(false);
            this.gbFilter.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.epError)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnDown;
        private System.Windows.Forms.Button btnUp;
        private System.Windows.Forms.ListBox lbFilters;
        private System.Windows.Forms.Button btnBrackets;
        private System.Windows.Forms.Button btnOr;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnChange;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.GroupBox gbFilter;
        private System.Windows.Forms.TextBox fText;
        private System.Windows.Forms.ComboBox cmbFilters;
        private EducServLib.FilterFromTo fFromTo;
        private EducServLib.FilterBool fBool;
        private EducServLib.FilterDateFromTo fDateFromTo;
        private EducServLib.FilterMultySelect fMult;
        private System.Windows.Forms.ErrorProvider epError;
    }
}