
namespace HsinChu.JHEvaluation.EduAdminExtendControls.Ribbon
{
    partial class CheckCourseScoreInput
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnExport = new DevComponents.DotNetBar.ButtonX();
            this.lblSave = new DevComponents.DotNetBar.LabelX();
            this.btnClose = new DevComponents.DotNetBar.ButtonX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.cboExamList = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lblExam = new DevComponents.DotNetBar.LabelX();
            this.dgv = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.chCourseName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chTeacherName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chClassName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chSeatNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chStudentNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chInputScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chInputAssignmentScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chInputText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExport
            // 
            this.btnExport.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnExport.BackColor = System.Drawing.Color.Transparent;
            this.btnExport.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExport.Location = new System.Drawing.Point(12, 527);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 26;
            this.btnExport.Text = "匯出";
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // lblSave
            // 
            this.lblSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSave.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblSave.BackgroundStyle.Class = "";
            this.lblSave.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblSave.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblSave.ForeColor = System.Drawing.Color.Red;
            this.lblSave.Location = new System.Drawing.Point(679, 527);
            this.lblSave.Name = "lblSave";
            this.lblSave.Size = new System.Drawing.Size(57, 23);
            this.lblSave.TabIndex = 25;
            this.lblSave.Text = "未儲存";
            this.lblSave.TextAlignment = System.Drawing.StringAlignment.Center;
            this.lblSave.Visible = false;
            // 
            // btnClose
            // 
            this.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackColor = System.Drawing.Color.Transparent;
            this.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnClose.Location = new System.Drawing.Point(825, 527);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 23;
            this.btnClose.Text = "關閉";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(745, 527);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 24;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // cboExamList
            // 
            this.cboExamList.DisplayMember = "DisplayText";
            this.cboExamList.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboExamList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboExamList.FormattingEnabled = true;
            this.cboExamList.ItemHeight = 19;
            this.cboExamList.Location = new System.Drawing.Point(56, 45);
            this.cboExamList.Name = "cboExamList";
            this.cboExamList.Size = new System.Drawing.Size(206, 25);
            this.cboExamList.TabIndex = 22;
            this.cboExamList.SelectedIndexChanged += new System.EventHandler(this.cboExamList_SelectedIndexChanged);
            // 
            // lblExam
            // 
            this.lblExam.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblExam.BackgroundStyle.Class = "";
            this.lblExam.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblExam.Location = new System.Drawing.Point(12, 46);
            this.lblExam.Name = "lblExam";
            this.lblExam.Size = new System.Drawing.Size(52, 23);
            this.lblExam.TabIndex = 21;
            this.lblExam.Text = "試別：";
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.AliceBlue;
            this.dgv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv.BackgroundColor = System.Drawing.Color.White;
            this.dgv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chCourseName,
            this.chTeacherName,
            this.chClassName,
            this.chSeatNo,
            this.chName,
            this.chStudentNumber,
            this.chInputScore,
            this.chInputAssignmentScore,
            this.chInputText});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgv.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgv.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.dgv.Location = new System.Drawing.Point(12, 75);
            this.dgv.Name = "dgv";
            this.dgv.RowHeadersWidth = 25;
            this.dgv.RowTemplate.Height = 24;
            this.dgv.Size = new System.Drawing.Size(888, 446);
            this.dgv.TabIndex = 20;
            this.dgv.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellClick);
            this.dgv.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            this.dgv.SortCompare += new System.Windows.Forms.DataGridViewSortCompareEventHandler(this.dgv_SortCompare);
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(12, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(60, 21);
            this.labelX1.TabIndex = 27;
            this.labelX1.Text = "學年度：";
            // 
            // iptSchoolYear
            // 
            this.iptSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSchoolYear.Location = new System.Drawing.Point(78, 10);
            this.iptSchoolYear.MaxValue = 300;
            this.iptSchoolYear.MinValue = 80;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.ShowUpDown = true;
            this.iptSchoolYear.Size = new System.Drawing.Size(62, 25);
            this.iptSchoolYear.TabIndex = 28;
            this.iptSchoolYear.Value = 80;
            // 
            // iptSemester
            // 
            this.iptSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSemester.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSemester.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSemester.Location = new System.Drawing.Point(216, 10);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 1;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.ShowUpDown = true;
            this.iptSemester.Size = new System.Drawing.Size(62, 25);
            this.iptSemester.TabIndex = 30;
            this.iptSemester.Value = 1;
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(163, 12);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(47, 21);
            this.labelX2.TabIndex = 29;
            this.labelX2.Text = "學期：";
            // 
            // chCourseName
            // 
            this.chCourseName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chCourseName.FillWeight = 250F;
            this.chCourseName.HeaderText = "課程名稱";
            this.chCourseName.Name = "chCourseName";
            this.chCourseName.ReadOnly = true;
            // 
            // chTeacherName
            // 
            this.chTeacherName.HeaderText = "授課教師";
            this.chTeacherName.Name = "chTeacherName";
            this.chTeacherName.ReadOnly = true;
            // 
            // chClassName
            // 
            this.chClassName.HeaderText = "班級";
            this.chClassName.Name = "chClassName";
            this.chClassName.ReadOnly = true;
            // 
            // chSeatNo
            // 
            this.chSeatNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.chSeatNo.HeaderText = "座號";
            this.chSeatNo.Name = "chSeatNo";
            this.chSeatNo.ReadOnly = true;
            this.chSeatNo.Width = 59;
            // 
            // chName
            // 
            this.chName.HeaderText = "姓名";
            this.chName.Name = "chName";
            this.chName.ReadOnly = true;
            // 
            // chStudentNumber
            // 
            this.chStudentNumber.HeaderText = "學號";
            this.chStudentNumber.Name = "chStudentNumber";
            this.chStudentNumber.ReadOnly = true;
            // 
            // chInputScore
            // 
            this.chInputScore.HeaderText = "定期分數";
            this.chInputScore.Name = "chInputScore";
            // 
            // chInputAssignmentScore
            // 
            this.chInputAssignmentScore.HeaderText = "平時分數";
            this.chInputAssignmentScore.Name = "chInputAssignmentScore";
            // 
            // chInputText
            // 
            this.chInputText.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chInputText.HeaderText = "文字描述";
            this.chInputText.Name = "chInputText";
            this.chInputText.Visible = false;
            // 
            // CheckCourseScoreInput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(915, 561);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.lblSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.iptSchoolYear);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.cboExamList);
            this.Controls.Add(this.lblExam);
            this.Controls.Add(this.dgv);
            this.DoubleBuffered = true;
            this.Name = "CheckCourseScoreInput";
            this.Text = "評量缺免成績查詢調整";
            this.Load += new System.EventHandler(this.CheckCourseScoreInput_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnExport;
        private DevComponents.DotNetBar.LabelX lblSave;
        private DevComponents.DotNetBar.ButtonX btnClose;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboExamList;
        private DevComponents.DotNetBar.LabelX lblExam;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgv;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.Editors.IntegerInput iptSchoolYear;
        private DevComponents.Editors.IntegerInput iptSemester;
        private DevComponents.DotNetBar.LabelX labelX2;
        private System.Windows.Forms.DataGridViewTextBoxColumn chCourseName;
        private System.Windows.Forms.DataGridViewTextBoxColumn chTeacherName;
        private System.Windows.Forms.DataGridViewTextBoxColumn chClassName;
        private System.Windows.Forms.DataGridViewTextBoxColumn chSeatNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn chName;
        private System.Windows.Forms.DataGridViewTextBoxColumn chStudentNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn chInputScore;
        private System.Windows.Forms.DataGridViewTextBoxColumn chInputAssignmentScore;
        private System.Windows.Forms.DataGridViewTextBoxColumn chInputText;
    }
}