namespace JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReportControls
{
    partial class GraduationPredictReportForm
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
            this.gpScore = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.chConditionGr1 = new System.Windows.Forms.CheckBox();
            this.gpDaily = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.chCondition4c = new System.Windows.Forms.CheckBox();
            this.chCondition5c = new System.Windows.Forms.CheckBox();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.ExportDoctSetup = new System.Windows.Forms.LinkLabel();
            this.checkExportDoc = new System.Windows.Forms.CheckBox();
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            this.gpScore.SuspendLayout();
            this.gpDaily.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            this.SuspendLayout();
            // 
            // gpScore
            // 
            this.gpScore.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gpScore.BackColor = System.Drawing.Color.Transparent;
            this.gpScore.CanvasColor = System.Drawing.SystemColors.Control;
            this.gpScore.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.gpScore.Controls.Add(this.chConditionGr1);
            this.gpScore.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.gpScore.Location = new System.Drawing.Point(9, 48);
            this.gpScore.Name = "gpScore";
            this.gpScore.Size = new System.Drawing.Size(378, 78);
            // 
            // 
            // 
            this.gpScore.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.gpScore.Style.BackColorGradientAngle = 90;
            this.gpScore.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.gpScore.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpScore.Style.BorderBottomWidth = 1;
            this.gpScore.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.gpScore.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpScore.Style.BorderLeftWidth = 1;
            this.gpScore.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpScore.Style.BorderRightWidth = 1;
            this.gpScore.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpScore.Style.BorderTopWidth = 1;
            this.gpScore.Style.Class = "";
            this.gpScore.Style.CornerDiameter = 4;
            this.gpScore.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.gpScore.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.gpScore.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.gpScore.StyleMouseDown.Class = "";
            this.gpScore.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.gpScore.StyleMouseOver.Class = "";
            this.gpScore.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.gpScore.TabIndex = 0;
            this.gpScore.Text = "學業成績";
            // 
            // chConditionGr1
            // 
            this.chConditionGr1.AutoSize = true;
            this.chConditionGr1.BackColor = System.Drawing.Color.Transparent;
            this.chConditionGr1.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.chConditionGr1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chConditionGr1.Location = new System.Drawing.Point(17, 15);
            this.chConditionGr1.Name = "chConditionGr1";
            this.chConditionGr1.Size = new System.Drawing.Size(304, 26);
            this.chConditionGr1.TabIndex = 3;
            this.chConditionGr1.Tag = "GraduateDomain";
            this.chConditionGr1.Text = "學習領域畢業總平均成績符合規範。";
            this.chConditionGr1.UseVisualStyleBackColor = false;
            this.chConditionGr1.CheckedChanged += new System.EventHandler(this.chConditionGr1_CheckedChanged);
            // 
            // gpDaily
            // 
            this.gpDaily.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gpDaily.BackColor = System.Drawing.Color.Transparent;
            this.gpDaily.CanvasColor = System.Drawing.SystemColors.Control;
            this.gpDaily.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.gpDaily.Controls.Add(this.checkBox4);
            this.gpDaily.Controls.Add(this.chCondition4c);
            this.gpDaily.Controls.Add(this.chCondition5c);
            this.gpDaily.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.gpDaily.Location = new System.Drawing.Point(9, 132);
            this.gpDaily.Name = "gpDaily";
            this.gpDaily.Size = new System.Drawing.Size(379, 170);
            // 
            // 
            // 
            this.gpDaily.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.gpDaily.Style.BackColorGradientAngle = 90;
            this.gpDaily.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.gpDaily.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDaily.Style.BorderBottomWidth = 1;
            this.gpDaily.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.gpDaily.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDaily.Style.BorderLeftWidth = 1;
            this.gpDaily.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDaily.Style.BorderRightWidth = 1;
            this.gpDaily.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDaily.Style.BorderTopWidth = 1;
            this.gpDaily.Style.Class = "";
            this.gpDaily.Style.CornerDiameter = 4;
            this.gpDaily.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.gpDaily.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.gpDaily.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.gpDaily.StyleMouseDown.Class = "";
            this.gpDaily.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.gpDaily.StyleMouseOver.Class = "";
            this.gpDaily.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.gpDaily.TabIndex = 3;
            this.gpDaily.Text = "日常生活表現";
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.BackColor = System.Drawing.Color.Transparent;
            this.checkBox4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.checkBox4.Location = new System.Drawing.Point(17, 17);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(134, 26);
            this.checkBox4.TabIndex = 20;
            this.checkBox4.Text = "勾選以下條件";
            this.checkBox4.UseVisualStyleBackColor = true;
            this.checkBox4.CheckedChanged += new System.EventHandler(this.checkBox4_CheckedChanged);
            // 
            // chCondition4c
            // 
            this.chCondition4c.AutoSize = true;
            this.chCondition4c.BackColor = System.Drawing.Color.Transparent;
            this.chCondition4c.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.chCondition4c.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition4c.Location = new System.Drawing.Point(17, 46);
            this.chCondition4c.Name = "chCondition4c";
            this.chCondition4c.Size = new System.Drawing.Size(338, 26);
            this.chCondition4c.TabIndex = 17;
            this.chCondition4c.Tag = "AbsenceAmountAllFraction";
            this.chCondition4c.Text = "所有學期缺課節數超過總節數指定比例。";
            this.chCondition4c.UseVisualStyleBackColor = false;
            // 
            // chCondition5c
            // 
            this.chCondition5c.AutoSize = true;
            this.chCondition5c.BackColor = System.Drawing.Color.Transparent;
            this.chCondition5c.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.chCondition5c.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition5c.Location = new System.Drawing.Point(17, 73);
            this.chCondition5c.Name = "chCondition5c";
            this.chCondition5c.Size = new System.Drawing.Size(253, 26);
            this.chCondition5c.TabIndex = 18;
            this.chCondition5c.Tag = "DemeritAmountAll";
            this.chCondition5c.Text = "所有學期懲戒次數合計超次。";
            this.chCondition5c.UseVisualStyleBackColor = false;
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnPrint.Location = new System.Drawing.Point(231, 322);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 23);
            this.btnPrint.TabIndex = 8;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnExit.Location = new System.Drawing.Point(312, 322);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 9;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(9, 14);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(75, 23);
            this.labelX1.TabIndex = 10;
            this.labelX1.Text = "目前學年度";
            // 
            // labelX2
            // 
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(179, 14);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(42, 23);
            this.labelX2.TabIndex = 11;
            this.labelX2.Text = "學期";
            // 
            // ExportDoctSetup
            // 
            this.ExportDoctSetup.AutoSize = true;
            this.ExportDoctSetup.BackColor = System.Drawing.Color.Transparent;
            this.ExportDoctSetup.Location = new System.Drawing.Point(113, 325);
            this.ExportDoctSetup.Name = "ExportDoctSetup";
            this.ExportDoctSetup.Size = new System.Drawing.Size(129, 22);
            this.ExportDoctSetup.TabIndex = 14;
            this.ExportDoctSetup.TabStop = true;
            this.ExportDoctSetup.Text = "樣板與列印設定";
            this.ExportDoctSetup.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ExportDoctSetup_LinkClicked);
            // 
            // checkExportDoc
            // 
            this.checkExportDoc.AutoSize = true;
            this.checkExportDoc.BackColor = System.Drawing.Color.Transparent;
            this.checkExportDoc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.checkExportDoc.Location = new System.Drawing.Point(15, 324);
            this.checkExportDoc.Name = "checkExportDoc";
            this.checkExportDoc.Size = new System.Drawing.Size(117, 26);
            this.checkExportDoc.TabIndex = 16;
            this.checkExportDoc.Text = "產生通知單";
            this.checkExportDoc.UseVisualStyleBackColor = false;
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
            this.iptSchoolYear.Location = new System.Drawing.Point(91, 13);
            this.iptSchoolYear.MaxValue = 999;
            this.iptSchoolYear.MinValue = 1;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.ShowUpDown = true;
            this.iptSchoolYear.Size = new System.Drawing.Size(80, 29);
            this.iptSchoolYear.TabIndex = 17;
            this.iptSchoolYear.Value = 1;
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
            this.iptSemester.Location = new System.Drawing.Point(219, 13);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 1;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.ShowUpDown = true;
            this.iptSemester.Size = new System.Drawing.Size(80, 29);
            this.iptSemester.TabIndex = 18;
            this.iptSemester.Value = 1;
            // 
            // GraduationPredictReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(398, 354);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.iptSchoolYear);
            this.Controls.Add(this.ExportDoctSetup);
            this.Controls.Add(this.checkExportDoc);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.gpDaily);
            this.Controls.Add(this.gpScore);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "GraduationPredictReportForm";
            this.Text = "畢業預警報表";
            this.Load += new System.EventHandler(this.GraduationPredictReportForm_Load);
            this.gpScore.ResumeLayout(false);
            this.gpScore.PerformLayout();
            this.gpDaily.ResumeLayout(false);
            this.gpDaily.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private DevComponents.DotNetBar.Controls.GroupPanel gpScore;
        private DevComponents.DotNetBar.Controls.GroupPanel gpDaily;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private System.Windows.Forms.LinkLabel ExportDoctSetup;
        private System.Windows.Forms.CheckBox checkExportDoc;
        private System.Windows.Forms.CheckBox chConditionGr1;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.CheckBox chCondition4c;
        private System.Windows.Forms.CheckBox chCondition5c;
        private DevComponents.Editors.IntegerInput iptSchoolYear;
        private DevComponents.Editors.IntegerInput iptSemester;
    }
}