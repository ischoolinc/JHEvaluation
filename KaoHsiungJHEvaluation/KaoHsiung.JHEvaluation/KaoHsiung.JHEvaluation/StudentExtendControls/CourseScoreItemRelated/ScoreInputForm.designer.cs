﻿namespace KaoHsiung.JHEvaluation.StudentExtendControls.CourseScoreItemControls
{
    partial class ScoreInputForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblCourseName = new DevComponents.DotNetBar.LabelX();
            this.dgv = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.chExamName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chEffort = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.lblNotSave = new DevComponents.DotNetBar.LabelX();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblStudent = new DevComponents.DotNetBar.LabelX();
            this.lblWaringMsg = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // lblCourseName
            // 
            this.lblCourseName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblCourseName.BackgroundStyle.Class = "";
            this.lblCourseName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblCourseName.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblCourseName.Location = new System.Drawing.Point(6, 5);
            this.lblCourseName.Name = "lblCourseName";
            this.lblCourseName.Size = new System.Drawing.Size(316, 35);
            this.lblCourseName.TabIndex = 0;
            this.lblCourseName.Text = "301 (語文)國語文";
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.BackgroundColor = System.Drawing.Color.White;
            this.dgv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chExamName,
            this.chScore,
            this.chEffort,
            this.chText});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgv.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgv.ImeMode = System.Windows.Forms.ImeMode.Close;
            this.dgv.Location = new System.Drawing.Point(6, 46);
            this.dgv.Name = "dgv";
            this.dgv.RowHeadersVisible = false;
            this.dgv.RowTemplate.Height = 24;
            this.dgv.Size = new System.Drawing.Size(589, 200);
            this.dgv.TabIndex = 1;
            this.dgv.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            // 
            // chExamName
            // 
            this.chExamName.HeaderText = "評量名稱";
            this.chExamName.Name = "chExamName";
            this.chExamName.ReadOnly = true;
            this.chExamName.Width = 170;
            // 
            // chScore
            // 
            this.chScore.HeaderText = "分數評量";
            this.chScore.Name = "chScore";
            this.chScore.Width = 83;
            // 
            // chEffort
            // 
            this.chEffort.HeaderText = "努力程度";
            this.chEffort.Name = "chEffort";
            this.chEffort.Width = 83;
            // 
            // chText
            // 
            this.chText.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chText.HeaderText = "文字描述";
            this.chText.Name = "chText";
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(438, 252);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(520, 252);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // lblNotSave
            // 
            this.lblNotSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblNotSave.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblNotSave.BackgroundStyle.Class = "";
            this.lblNotSave.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblNotSave.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblNotSave.ForeColor = System.Drawing.Color.Red;
            this.lblNotSave.Location = new System.Drawing.Point(357, 252);
            this.lblNotSave.Name = "lblNotSave";
            this.lblNotSave.Size = new System.Drawing.Size(78, 23);
            this.lblNotSave.TabIndex = 3;
            this.lblNotSave.Text = "尚未儲存";
            this.lblNotSave.TextAlignment = System.Drawing.StringAlignment.Center;
            this.lblNotSave.Visible = false;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "評量名稱";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 170;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.HeaderText = "分數評量";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 83;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.HeaderText = "努力程度";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 83;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn4.HeaderText = "文字描述";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // lblStudent
            // 
            this.lblStudent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStudent.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblStudent.BackgroundStyle.Class = "";
            this.lblStudent.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblStudent.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblStudent.Location = new System.Drawing.Point(357, 5);
            this.lblStudent.Name = "lblStudent";
            this.lblStudent.Size = new System.Drawing.Size(238, 35);
            this.lblStudent.TabIndex = 0;
            this.lblStudent.TextAlignment = System.Drawing.StringAlignment.Far;
            // 
            // lblWaringMsg
            // 
            this.lblWaringMsg.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblWaringMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblWaringMsg.BackgroundStyle.Class = "";
            this.lblWaringMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblWaringMsg.ForeColor = System.Drawing.Color.Red;
            this.lblWaringMsg.Location = new System.Drawing.Point(6, 252);
            this.lblWaringMsg.Name = "lblWaringMsg";
            this.lblWaringMsg.Size = new System.Drawing.Size(345, 23);
            this.lblWaringMsg.TabIndex = 4;
            this.lblWaringMsg.Text = "紅色評量名稱， 因評分樣板中並沒有設定這個評量名稱。";
            // 
            // ScoreInputForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(601, 281);
            this.Controls.Add(this.lblWaringMsg);
            this.Controls.Add(this.lblNotSave);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.lblStudent);
            this.Controls.Add(this.lblCourseName);
            this.DoubleBuffered = true;
            this.Name = "ScoreInputForm";
            this.Text = "評量成績";
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblCourseName;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgv;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private System.Windows.Forms.DataGridViewTextBoxColumn chExamName;
        private System.Windows.Forms.DataGridViewTextBoxColumn chScore;
        private System.Windows.Forms.DataGridViewTextBoxColumn chEffort;
        private System.Windows.Forms.DataGridViewTextBoxColumn chText;
        private DevComponents.DotNetBar.LabelX lblNotSave;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private DevComponents.DotNetBar.LabelX lblStudent;
        private DevComponents.DotNetBar.LabelX lblWaringMsg;
    }
}