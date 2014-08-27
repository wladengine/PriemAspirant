namespace Priem
{
    partial class Groups
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Groups));
            this.btnOK = new System.Windows.Forms.Button();
            this.chbPrintGroup = new System.Windows.Forms.CheckBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.btnRightAll = new System.Windows.Forms.Button();
            this.btnLeftAll = new System.Windows.Forms.Button();
            this.btnRight = new System.Windows.Forms.Button();
            this.btnLeft = new System.Windows.Forms.Button();
            this.lblExclude = new System.Windows.Forms.Label();
            this.lblInclude = new System.Windows.Forms.Label();
            this.lbNo = new System.Windows.Forms.ListBox();
            this.lbYes = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(203, 249);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(118, 23);
            this.btnOK.TabIndex = 27;
            this.btnOK.Text = "Подтвердить";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // chbPrintGroup
            // 
            this.chbPrintGroup.AutoSize = true;
            this.chbPrintGroup.Location = new System.Drawing.Point(262, 217);
            this.chbPrintGroup.Name = "chbPrintGroup";
            this.chbPrintGroup.Size = new System.Drawing.Size(192, 17);
            this.chbPrintGroup.TabIndex = 29;
            this.chbPrintGroup.Text = "Разделять при печати на группы";
            this.chbPrintGroup.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(336, 249);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(118, 23);
            this.btnCancel.TabIndex = 27;
            this.btnCancel.Text = "Отменить";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnDown
            // 
            this.btnDown.Enabled = false;
            this.btnDown.Location = new System.Drawing.Point(8, 111);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(30, 31);
            this.btnDown.TabIndex = 31;
            this.btnDown.Text = "\\/";
            this.btnDown.UseVisualStyleBackColor = true;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.Enabled = false;
            this.btnUp.Location = new System.Drawing.Point(8, 74);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(30, 31);
            this.btnUp.TabIndex = 30;
            this.btnUp.Text = "/\\";
            this.btnUp.UseVisualStyleBackColor = true;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // btnRightAll
            // 
            this.btnRightAll.Location = new System.Drawing.Point(247, 168);
            this.btnRightAll.Name = "btnRightAll";
            this.btnRightAll.Size = new System.Drawing.Size(33, 23);
            this.btnRightAll.TabIndex = 36;
            this.btnRightAll.Text = ">>";
            this.btnRightAll.UseVisualStyleBackColor = true;
            this.btnRightAll.Click += new System.EventHandler(this.btnRightAll_Click);
            // 
            // btnLeftAll
            // 
            this.btnLeftAll.Location = new System.Drawing.Point(212, 130);
            this.btnLeftAll.Name = "btnLeftAll";
            this.btnLeftAll.Size = new System.Drawing.Size(33, 23);
            this.btnLeftAll.TabIndex = 37;
            this.btnLeftAll.Text = "<<";
            this.btnLeftAll.UseVisualStyleBackColor = true;
            this.btnLeftAll.Click += new System.EventHandler(this.btnLeftAll_Click);
            // 
            // btnRight
            // 
            this.btnRight.Location = new System.Drawing.Point(247, 74);
            this.btnRight.Name = "btnRight";
            this.btnRight.Size = new System.Drawing.Size(33, 23);
            this.btnRight.TabIndex = 38;
            this.btnRight.Text = ">";
            this.btnRight.UseVisualStyleBackColor = true;
            this.btnRight.Click += new System.EventHandler(this.btnRight_Click);
            // 
            // btnLeft
            // 
            this.btnLeft.Location = new System.Drawing.Point(212, 36);
            this.btnLeft.Name = "btnLeft";
            this.btnLeft.Size = new System.Drawing.Size(33, 23);
            this.btnLeft.TabIndex = 39;
            this.btnLeft.Text = "<";
            this.btnLeft.UseVisualStyleBackColor = true;
            this.btnLeft.Click += new System.EventHandler(this.btnLeft_Click);
            // 
            // lblExclude
            // 
            this.lblExclude.AutoSize = true;
            this.lblExclude.Location = new System.Drawing.Point(333, 9);
            this.lblExclude.Name = "lblExclude";
            this.lblExclude.Size = new System.Drawing.Size(71, 13);
            this.lblExclude.TabIndex = 35;
            this.lblExclude.Text = "Не выбрано:";
            // 
            // lblInclude
            // 
            this.lblInclude.AutoSize = true;
            this.lblInclude.Location = new System.Drawing.Point(85, 9);
            this.lblInclude.Name = "lblInclude";
            this.lblInclude.Size = new System.Drawing.Size(55, 13);
            this.lblInclude.TabIndex = 34;
            this.lblInclude.Text = "Выбрано:";
            // 
            // lbNo
            // 
            this.lbNo.FormattingEnabled = true;
            this.lbNo.HorizontalScrollbar = true;
            this.lbNo.Location = new System.Drawing.Point(286, 25);
            this.lbNo.Name = "lbNo";
            this.lbNo.Size = new System.Drawing.Size(163, 186);
            this.lbNo.Sorted = true;
            this.lbNo.TabIndex = 32;
            // 
            // lbYes
            // 
            this.lbYes.FormattingEnabled = true;
            this.lbYes.HorizontalScrollbar = true;
            this.lbYes.Location = new System.Drawing.Point(43, 25);
            this.lbYes.Name = "lbYes";
            this.lbYes.Size = new System.Drawing.Size(163, 186);
            this.lbYes.TabIndex = 33;
            this.lbYes.SelectedIndexChanged += new System.EventHandler(this.lbYes_SelectedIndexChanged);
            // 
            // Groups
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(462, 286);
            this.Controls.Add(this.btnRightAll);
            this.Controls.Add(this.btnLeftAll);
            this.Controls.Add(this.btnRight);
            this.Controls.Add(this.btnLeft);
            this.Controls.Add(this.lblExclude);
            this.Controls.Add(this.lblInclude);
            this.Controls.Add(this.lbNo);
            this.Controls.Add(this.lbYes);
            this.Controls.Add(this.btnDown);
            this.Controls.Add(this.btnUp);
            this.Controls.Add(this.chbPrintGroup);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Groups";
            this.Text = "Группировки";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.CheckBox chbPrintGroup;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnDown;
        private System.Windows.Forms.Button btnUp;
        private System.Windows.Forms.Button btnRightAll;
        private System.Windows.Forms.Button btnLeftAll;
        private System.Windows.Forms.Button btnRight;
        private System.Windows.Forms.Button btnLeft;
        private System.Windows.Forms.Label lblExclude;
        private System.Windows.Forms.Label lblInclude;
        private System.Windows.Forms.ListBox lbNo;
        private System.Windows.Forms.ListBox lbYes;
    }
}