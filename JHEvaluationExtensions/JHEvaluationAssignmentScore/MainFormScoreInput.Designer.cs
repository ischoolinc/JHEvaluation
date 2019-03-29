namespace JHEvaluation.AssignmentScore
{
    partial class MainFormScoreInput
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblCourseName = new DevComponents.DotNetBar.LabelX();
            this.btn_Save = new DevComponents.DotNetBar.ButtonX();
            this.btn_Exit = new DevComponents.DotNetBar.ButtonX();
            this.lblSave = new DevComponents.DotNetBar.LabelX();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dgv = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colInputEffort = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colInputScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStudentNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSeatNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colClassName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
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
            this.lblCourseName.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblCourseName.Location = new System.Drawing.Point(7, 13);
            this.lblCourseName.Name = "lblCourseName";
            this.lblCourseName.Size = new System.Drawing.Size(448, 23);
            this.lblCourseName.TabIndex = 0;
            this.lblCourseName.Text = "labelX1";
            // 
            // btn_Save
            // 
            this.btn_Save.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btn_Save.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_Save.BackColor = System.Drawing.Color.Transparent;
            this.btn_Save.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btn_Save.Location = new System.Drawing.Point(549, 13);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(75, 23);
            this.btn_Save.TabIndex = 1;
            this.btn_Save.Text = "儲存";
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // btn_Exit
            // 
            this.btn_Exit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btn_Exit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_Exit.BackColor = System.Drawing.Color.Transparent;
            this.btn_Exit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btn_Exit.Location = new System.Drawing.Point(553, 454);
            this.btn_Exit.Name = "btn_Exit";
            this.btn_Exit.Size = new System.Drawing.Size(75, 23);
            this.btn_Exit.TabIndex = 2;
            this.btn_Exit.Text = "離開";
            this.btn_Exit.Click += new System.EventHandler(this.btn_Exit_Click);
            // 
            // lblSave
            // 
            this.lblSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSave.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblSave.BackgroundStyle.Class = "";
            this.lblSave.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblSave.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblSave.ForeColor = System.Drawing.Color.Red;
            this.lblSave.Location = new System.Drawing.Point(461, 13);
            this.lblSave.Name = "lblSave";
            this.lblSave.Size = new System.Drawing.Size(82, 23);
            this.lblSave.TabIndex = 6;
            this.lblSave.Text = "未儲存";
            this.lblSave.TextAlignment = System.Drawing.StringAlignment.Center;
            this.lblSave.Visible = false;
            // 
            // splitContainer1
            // 
            this.splitContainer1.BackColor = System.Drawing.Color.Transparent;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(644, 486);
            this.splitContainer1.SplitterDistance = 442;
            this.splitContainer1.TabIndex = 8;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            this.dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colClassName,
            this.colSeatNo,
            this.colName,
            this.colStudentNumber,
            this.colInputScore,
            this.colInputEffort});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgv.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgv.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.dgv.Location = new System.Drawing.Point(0, 0);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 24;
            this.dgv.Size = new System.Drawing.Size(644, 394);
            this.dgv.TabIndex = 3;
            this.dgv.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellClick);
            this.dgv.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            // 
            // colInputEffort
            // 
            this.colInputEffort.HeaderText = "努力程度";
            this.colInputEffort.Name = "colInputEffort";
            // 
            // colInputScore
            // 
            this.colInputScore.HeaderText = "平時評量";
            this.colInputScore.Name = "colInputScore";
            // 
            // colStudentNumber
            // 
            this.colStudentNumber.HeaderText = "學號";
            this.colStudentNumber.Name = "colStudentNumber";
            // 
            // colName
            // 
            this.colName.HeaderText = "姓名";
            this.colName.Name = "colName";
            // 
            // colSeatNo
            // 
            this.colSeatNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colSeatNo.FillWeight = 65F;
            this.colSeatNo.HeaderText = "座號";
            this.colSeatNo.Name = "colSeatNo";
            // 
            // colClassName
            // 
            this.colClassName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.colClassName.HeaderText = "班級";
            this.colClassName.Name = "colClassName";
            this.colClassName.Width = 73;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.dgv);
            this.splitContainer2.Size = new System.Drawing.Size(644, 442);
            this.splitContainer2.SplitterDistance = 44;
            this.splitContainer2.TabIndex = 4;
            // 
            // MainFormScoreInput
            // 
            this.ClientSize = new System.Drawing.Size(644, 486);
            this.Controls.Add(this.btn_Exit);
            this.Controls.Add(this.lblSave);
            this.Controls.Add(this.btn_Save);
            this.Controls.Add(this.lblCourseName);
            this.Controls.Add(this.splitContainer1);
            this.DoubleBuffered = true;
            this.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.Name = "MainFormScoreInput";
            this.Text = "平時評量成績輸入";
            this.Load += new System.EventHandler(this.MainFormScoreInput_Load);
            this.Shown += new System.EventHandler(this.MainFormScoreInput_Shown);
            this.splitContainer1.Panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblCourseName;
        private DevComponents.DotNetBar.ButtonX btn_Save;
        private DevComponents.DotNetBar.ButtonX btn_Exit;
        private DevComponents.DotNetBar.LabelX lblSave;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn colClassName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSeatNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStudentNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInputScore;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInputEffort;
        private System.Windows.Forms.SplitContainer splitContainer2;
    }
}