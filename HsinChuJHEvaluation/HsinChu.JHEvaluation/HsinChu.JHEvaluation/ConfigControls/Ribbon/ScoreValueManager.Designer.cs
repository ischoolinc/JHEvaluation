
namespace HsinChu.JHEvaluation.ConfigControls.Ribbon
{
    partial class ScoreValueManager
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgvScoreConfig = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.lblNotSave = new DevComponents.DotNetBar.LabelX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.lblStudent = new DevComponents.DotNetBar.LabelX();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colUseText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colAllowCalculation = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colActive = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colUseValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSystemValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colReportValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvScoreConfig)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvScoreConfig
            // 
            this.dgvScoreConfig.AllowUserToAddRows = false;
            this.dgvScoreConfig.AllowUserToDeleteRows = false;
            this.dgvScoreConfig.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvScoreConfig.BackgroundColor = System.Drawing.Color.White;
            this.dgvScoreConfig.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvScoreConfig.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvScoreConfig.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvScoreConfig.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colUseText,
            this.colAllowCalculation,
            this.colScore,
            this.colActive,
            this.colUseValue,
            this.colSystemValue,
            this.colReportValue});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvScoreConfig.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvScoreConfig.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvScoreConfig.Location = new System.Drawing.Point(6, 6);
            this.dgvScoreConfig.Name = "dgvScoreConfig";
            this.dgvScoreConfig.RowHeadersVisible = false;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.dgvScoreConfig.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgvScoreConfig.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.dgvScoreConfig.RowTemplate.Height = 24;
            this.dgvScoreConfig.Size = new System.Drawing.Size(343, 117);
            this.dgvScoreConfig.TabIndex = 6;
            // 
            // lblNotSave
            // 
            this.lblNotSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblNotSave.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblNotSave.BackgroundStyle.Class = "";
            this.lblNotSave.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblNotSave.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblNotSave.ForeColor = System.Drawing.Color.Red;
            this.lblNotSave.Location = new System.Drawing.Point(6, 134);
            this.lblNotSave.Name = "lblNotSave";
            this.lblNotSave.Size = new System.Drawing.Size(90, 23);
            this.lblNotSave.TabIndex = 9;
            this.lblNotSave.Text = "尚未儲存";
            this.lblNotSave.TextAlignment = System.Drawing.StringAlignment.Center;
            this.lblNotSave.Visible = false;
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(274, 134);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 7;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(192, 134);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 8;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblStudent
            // 
            this.lblStudent.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblStudent.BackgroundStyle.Class = "";
            this.lblStudent.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblStudent.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblStudent.Location = new System.Drawing.Point(328, 5);
            this.lblStudent.Name = "lblStudent";
            this.lblStudent.Size = new System.Drawing.Size(238, 35);
            this.lblStudent.TabIndex = 5;
            this.lblStudent.TextAlignment = System.Drawing.StringAlignment.Far;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridViewTextBoxColumn1.HeaderText = "評量名稱";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 85;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridViewTextBoxColumn2.HeaderText = "分數評量";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 85;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.HeaderText = "努力程度";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            this.dataGridViewTextBoxColumn3.Visible = false;
            this.dataGridViewTextBoxColumn3.Width = 83;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn4.HeaderText = "文字描述";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.Visible = false;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.HeaderText = "報表分數";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.Visible = false;
            // 
            // colUseText
            // 
            this.colUseText.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colUseText.HeaderText = "使用文字";
            this.colUseText.Name = "colUseText";
            this.colUseText.Width = 85;
            // 
            // colAllowCalculation
            // 
            this.colAllowCalculation.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colAllowCalculation.HeaderText = "是否計算成績";
            this.colAllowCalculation.Name = "colAllowCalculation";
            this.colAllowCalculation.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colAllowCalculation.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colAllowCalculation.Width = 111;
            // 
            // colScore
            // 
            this.colScore.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colScore.HeaderText = "計算分數";
            this.colScore.Name = "colScore";
            this.colScore.Width = 85;
            // 
            // colActive
            // 
            this.colActive.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colActive.HeaderText = "啟用";
            this.colActive.Name = "colActive";
            this.colActive.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colActive.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colActive.Width = 59;
            // 
            // colUseValue
            // 
            this.colUseValue.HeaderText = "替代分數";
            this.colUseValue.Name = "colUseValue";
            this.colUseValue.ReadOnly = true;
            this.colUseValue.Visible = false;
            // 
            // colSystemValue
            // 
            this.colSystemValue.HeaderText = "鍵值";
            this.colSystemValue.Name = "colSystemValue";
            this.colSystemValue.Visible = false;
            // 
            // colReportValue
            // 
            this.colReportValue.HeaderText = "報表分數";
            this.colReportValue.Name = "colReportValue";
            this.colReportValue.Visible = false;
            // 
            // ScoreValueManager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(355, 163);
            this.Controls.Add(this.dgvScoreConfig);
            this.Controls.Add(this.lblNotSave);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblStudent);
            this.DoubleBuffered = true;
            this.Name = "ScoreValueManager";
            this.Text = "評量成績缺考/免試設定";
            this.Load += new System.EventHandler(this.ScoreValueManager_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvScoreConfig)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgvScoreConfig;
        private DevComponents.DotNetBar.LabelX lblNotSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private DevComponents.DotNetBar.LabelX lblStudent;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUseText;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colAllowCalculation;
        private System.Windows.Forms.DataGridViewTextBoxColumn colScore;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colActive;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUseValue;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSystemValue;
        private System.Windows.Forms.DataGridViewTextBoxColumn colReportValue;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
    }
}