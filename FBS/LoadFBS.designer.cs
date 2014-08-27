namespace Priem
{
    partial class LoadFBS
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
            this.ofdFile = new System.Windows.Forms.OpenFileDialog();
            this.tbFile = new System.Windows.Forms.TextBox();
            this.btnLoad = new System.Windows.Forms.Button();
            this.dgvResult = new System.Windows.Forms.DataGridView();
            this.btnCreate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tbProtocolNum = new System.Windows.Forms.TextBox();
            this.rbEgeAnswerType2 = new System.Windows.Forms.RadioButton();
            this.rbEgeAnswerType1 = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ofdFile
            // 
            this.ofdFile.Filter = "Файл ФБС|*.csv|Все файлы|*.*";
            // 
            // tbFile
            // 
            this.tbFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tbFile.Location = new System.Drawing.Point(15, 68);
            this.tbFile.Name = "tbFile";
            this.tbFile.ReadOnly = true;
            this.tbFile.Size = new System.Drawing.Size(409, 20);
            this.tbFile.TabIndex = 0;
            // 
            // btnLoad
            // 
            this.btnLoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLoad.Location = new System.Drawing.Point(437, 65);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(75, 23);
            this.btnLoad.TabIndex = 1;
            this.btnLoad.Text = "Выбрать";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // dgvResult
            // 
            this.dgvResult.AllowUserToAddRows = false;
            this.dgvResult.AllowUserToDeleteRows = false;
            this.dgvResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvResult.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvResult.Location = new System.Drawing.Point(12, 94);
            this.dgvResult.Name = "dgvResult";
            this.dgvResult.Size = new System.Drawing.Size(500, 311);
            this.dgvResult.TabIndex = 2;
            // 
            // btnCreate
            // 
            this.btnCreate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreate.Location = new System.Drawing.Point(335, 434);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(177, 23);
            this.btnCreate.TabIndex = 3;
            this.btnCreate.Text = "Создать протокол  и Закрыть";
            this.btnCreate.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 439);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Номер протокола";
            // 
            // tbProtocolNum
            // 
            this.tbProtocolNum.Location = new System.Drawing.Point(115, 436);
            this.tbProtocolNum.Name = "tbProtocolNum";
            this.tbProtocolNum.ReadOnly = true;
            this.tbProtocolNum.Size = new System.Drawing.Size(100, 20);
            this.tbProtocolNum.TabIndex = 5;
            // 
            // rbEgeAnswerType2
            // 
            this.rbEgeAnswerType2.Checked = true;
            this.rbEgeAnswerType2.Location = new System.Drawing.Point(229, 13);
            this.rbEgeAnswerType2.Name = "rbEgeAnswerType2";
            this.rbEgeAnswerType2.Size = new System.Drawing.Size(174, 38);
            this.rbEgeAnswerType2.TabIndex = 6;
            this.rbEgeAnswerType2.TabStop = true;
            this.rbEgeAnswerType2.Text = "Запрос по ФИО(загрузить сертификаты целиком)";
            this.rbEgeAnswerType2.UseVisualStyleBackColor = true;
            // 
            // rbEgeAnswerType1
            // 
            this.rbEgeAnswerType1.Location = new System.Drawing.Point(9, 13);
            this.rbEgeAnswerType1.Name = "rbEgeAnswerType1";
            this.rbEgeAnswerType1.Size = new System.Drawing.Size(191, 38);
            this.rbEgeAnswerType1.TabIndex = 7;
            this.rbEgeAnswerType1.TabStop = true;
            this.rbEgeAnswerType1.Text = "Запрос по номеру сертификата (обновить статусы ФБС)";
            this.rbEgeAnswerType1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rbEgeAnswerType2);
            this.groupBox1.Controls.Add(this.rbEgeAnswerType1);
            this.groupBox1.Location = new System.Drawing.Point(15, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(409, 57);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            // 
            // LoadFBS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 469);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.tbProtocolNum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.dgvResult);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.tbFile);
            this.Name = "LoadFBS";
            this.Text = "Загрузка ответа ФБС";
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog ofdFile;
        private System.Windows.Forms.TextBox tbFile;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.DataGridView dgvResult;
        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbProtocolNum;
        private System.Windows.Forms.RadioButton rbEgeAnswerType2;
        private System.Windows.Forms.RadioButton rbEgeAnswerType1;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}