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
            this.chCondition1 = new System.Windows.Forms.CheckBox();
            this.chCondition2 = new System.Windows.Forms.CheckBox();
            this.chCondition3 = new System.Windows.Forms.CheckBox();
            this.chCondition4 = new System.Windows.Forms.CheckBox();
            this.chCondition5 = new System.Windows.Forms.CheckBox();
            this.chCondition6 = new System.Windows.Forms.CheckBox();
            this.gpScore = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.chConditionGr1 = new System.Windows.Forms.CheckBox();
            this.gpDaily = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.chCondition3b = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.chCondition4b = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.chCondition5b = new System.Windows.Forms.CheckBox();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.cboSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cboSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.ExportDoctSetup = new System.Windows.Forms.LinkLabel();
            this.gpDoc = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.checkExportDoc = new System.Windows.Forms.CheckBox();
            this.gpScore.SuspendLayout();
            this.gpDaily.SuspendLayout();
            this.gpDoc.SuspendLayout();
            this.SuspendLayout();
            // 
            // chCondition1
            // 
            this.chCondition1.AutoSize = true;
            this.chCondition1.BackColor = System.Drawing.Color.Transparent;
            this.chCondition1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition1.Location = new System.Drawing.Point(22, 11);
            this.chCondition1.Name = "chCondition1";
            this.chCondition1.Size = new System.Drawing.Size(196, 21);
            this.chCondition1.TabIndex = 1;
            this.chCondition1.Tag = "LearnDomainEach";
            this.chCondition1.Text = "各學期領域成績均符合規範。";
            this.chCondition1.UseVisualStyleBackColor = false;
            // 
            // chCondition2
            // 
            this.chCondition2.AutoSize = true;
            this.chCondition2.BackColor = System.Drawing.Color.Transparent;
            this.chCondition2.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition2.Location = new System.Drawing.Point(22, 40);
            this.chCondition2.Name = "chCondition2";
            this.chCondition2.Size = new System.Drawing.Size(209, 21);
            this.chCondition2.TabIndex = 2;
            this.chCondition2.Tag = "LearnDomainLast";
            this.chCondition2.Text = "第六學期各領域成績符合規範。";
            this.chCondition2.UseVisualStyleBackColor = false;
            // 
            // chCondition3
            // 
            this.chCondition3.AutoSize = true;
            this.chCondition3.BackColor = System.Drawing.Color.Transparent;
            this.chCondition3.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition3.Location = new System.Drawing.Point(51, 33);
            this.chCondition3.Name = "chCondition3";
            this.chCondition3.Size = new System.Drawing.Size(209, 21);
            this.chCondition3.TabIndex = 4;
            this.chCondition3.Tag = "AbsenceAmountEach";
            this.chCondition3.Text = "各學期缺課節數超過指定節數。";
            this.chCondition3.UseVisualStyleBackColor = false;
            // 
            // chCondition4
            // 
            this.chCondition4.AutoSize = true;
            this.chCondition4.BackColor = System.Drawing.Color.Transparent;
            this.chCondition4.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition4.Location = new System.Drawing.Point(51, 62);
            this.chCondition4.Name = "chCondition4";
            this.chCondition4.Size = new System.Drawing.Size(248, 21);
            this.chCondition4.TabIndex = 5;
            this.chCondition4.Tag = "AbsenceAmountEachFraction";
            this.chCondition4.Text = "各學期缺課節數超過總節數指定比例。";
            this.chCondition4.UseVisualStyleBackColor = false;
            // 
            // chCondition5
            // 
            this.chCondition5.AutoSize = true;
            this.chCondition5.BackColor = System.Drawing.Color.Transparent;
            this.chCondition5.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition5.Location = new System.Drawing.Point(51, 91);
            this.chCondition5.Name = "chCondition5";
            this.chCondition5.Size = new System.Drawing.Size(183, 21);
            this.chCondition5.TabIndex = 6;
            this.chCondition5.Tag = "DemeritAmountEach";
            this.chCondition5.Text = "各學期懲戒次數合計超次。";
            this.chCondition5.UseVisualStyleBackColor = false;
            // 
            // chCondition6
            // 
            this.chCondition6.AutoSize = true;
            this.chCondition6.BackColor = System.Drawing.Color.Transparent;
            this.chCondition6.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition6.Location = new System.Drawing.Point(51, 120);
            this.chCondition6.Name = "chCondition6";
            this.chCondition6.Size = new System.Drawing.Size(248, 21);
            this.chCondition6.TabIndex = 7;
            this.chCondition6.Tag = "DailyBehavior";
            this.chCondition6.Text = "各學期日常行為表現指標未符合標準。";
            this.chCondition6.UseVisualStyleBackColor = false;
            // 
            // gpScore
            // 
            this.gpScore.BackColor = System.Drawing.Color.Transparent;
            this.gpScore.CanvasColor = System.Drawing.SystemColors.Control;
            this.gpScore.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.gpScore.Controls.Add(this.chConditionGr1);
            this.gpScore.Controls.Add(this.chCondition1);
            this.gpScore.Controls.Add(this.chCondition2);
            this.gpScore.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.gpScore.Location = new System.Drawing.Point(9, 48);
            this.gpScore.Name = "gpScore";
            this.gpScore.Size = new System.Drawing.Size(346, 125);
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
            this.chConditionGr1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chConditionGr1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chConditionGr1.Location = new System.Drawing.Point(22, 69);
            this.chConditionGr1.Name = "chConditionGr1";
            this.chConditionGr1.Size = new System.Drawing.Size(235, 21);
            this.chConditionGr1.TabIndex = 3;
            this.chConditionGr1.Tag = "GraduateDomain";
            this.chConditionGr1.Text = "學習領域畢業總平均成績符合規範。";
            this.chConditionGr1.UseVisualStyleBackColor = false;
            this.chConditionGr1.CheckedChanged += new System.EventHandler(this.chConditionGr1_CheckedChanged);
            // 
            // gpDaily
            // 
            this.gpDaily.BackColor = System.Drawing.Color.Transparent;
            this.gpDaily.CanvasColor = System.Drawing.SystemColors.Control;
            this.gpDaily.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.gpDaily.Controls.Add(this.checkBox3);
            this.gpDaily.Controls.Add(this.chCondition3b);
            this.gpDaily.Controls.Add(this.checkBox2);
            this.gpDaily.Controls.Add(this.chCondition3);
            this.gpDaily.Controls.Add(this.chCondition4b);
            this.gpDaily.Controls.Add(this.chCondition4);
            this.gpDaily.Controls.Add(this.checkBox1);
            this.gpDaily.Controls.Add(this.chCondition6);
            this.gpDaily.Controls.Add(this.chCondition5b);
            this.gpDaily.Controls.Add(this.chCondition5);
            this.gpDaily.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.gpDaily.Location = new System.Drawing.Point(9, 179);
            this.gpDaily.Name = "gpDaily";
            this.gpDaily.Size = new System.Drawing.Size(346, 314);
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
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.BackColor = System.Drawing.Color.Transparent;
            this.checkBox3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.checkBox3.Location = new System.Drawing.Point(22, 149);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(105, 21);
            this.checkBox3.TabIndex = 15;
            this.checkBox3.Text = "勾選以下條件";
            this.checkBox3.UseVisualStyleBackColor = true;
            this.checkBox3.CheckedChanged += new System.EventHandler(this.checkBox3_CheckedChanged);
            // 
            // chCondition3b
            // 
            this.chCondition3b.AutoSize = true;
            this.chCondition3b.BackColor = System.Drawing.Color.Transparent;
            this.chCondition3b.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition3b.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition3b.Location = new System.Drawing.Point(51, 178);
            this.chCondition3b.Name = "chCondition3b";
            this.chCondition3b.Size = new System.Drawing.Size(222, 21);
            this.chCondition3b.TabIndex = 4;
            this.chCondition3b.Tag = "AbsenceAmountLast";
            this.chCondition3b.Text = "第六學期缺課節數超過指定節數。";
            this.chCondition3b.UseVisualStyleBackColor = false;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.BackColor = System.Drawing.Color.Transparent;
            this.checkBox2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.checkBox2.Location = new System.Drawing.Point(22, 4);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(105, 21);
            this.checkBox2.TabIndex = 14;
            this.checkBox2.Text = "勾選以下條件";
            this.checkBox2.UseVisualStyleBackColor = false;
            this.checkBox2.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // chCondition4b
            // 
            this.chCondition4b.AutoSize = true;
            this.chCondition4b.BackColor = System.Drawing.Color.Transparent;
            this.chCondition4b.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition4b.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition4b.Location = new System.Drawing.Point(51, 207);
            this.chCondition4b.Name = "chCondition4b";
            this.chCondition4b.Size = new System.Drawing.Size(261, 21);
            this.chCondition4b.TabIndex = 5;
            this.chCondition4b.Tag = "AbsenceAmountLastFraction";
            this.chCondition4b.Text = "第六學期缺課節數超過總節數指定比例。";
            this.chCondition4b.UseVisualStyleBackColor = false;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.BackColor = System.Drawing.Color.Transparent;
            this.checkBox1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.checkBox1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.checkBox1.Location = new System.Drawing.Point(51, 265);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(261, 21);
            this.checkBox1.TabIndex = 7;
            this.checkBox1.Tag = "DailyBehaviorLast";
            this.checkBox1.Text = "第六學期日常行為表現指標未符合標準。";
            this.checkBox1.UseVisualStyleBackColor = false;
            // 
            // chCondition5b
            // 
            this.chCondition5b.AutoSize = true;
            this.chCondition5b.BackColor = System.Drawing.Color.Transparent;
            this.chCondition5b.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.chCondition5b.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.chCondition5b.Location = new System.Drawing.Point(51, 236);
            this.chCondition5b.Name = "chCondition5b";
            this.chCondition5b.Size = new System.Drawing.Size(196, 21);
            this.chCondition5b.TabIndex = 6;
            this.chCondition5b.Tag = "DemeritAmountLast";
            this.chCondition5b.Text = "第六學期懲戒次數合計超次。";
            this.chCondition5b.UseVisualStyleBackColor = false;
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.btnPrint.Location = new System.Drawing.Point(197, 575);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 23);
            this.btnPrint.TabIndex = 8;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F);
            this.btnExit.Location = new System.Drawing.Point(278, 575);
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
            // cboSchoolYear
            // 
            this.cboSchoolYear.DisplayMember = "Text";
            this.cboSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSchoolYear.FormattingEnabled = true;
            this.cboSchoolYear.ItemHeight = 19;
            this.cboSchoolYear.Location = new System.Drawing.Point(88, 12);
            this.cboSchoolYear.Name = "cboSchoolYear";
            this.cboSchoolYear.Size = new System.Drawing.Size(65, 25);
            this.cboSchoolYear.TabIndex = 12;
            // 
            // cboSemester
            // 
            this.cboSemester.DisplayMember = "Text";
            this.cboSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSemester.FormattingEnabled = true;
            this.cboSemester.ItemHeight = 19;
            this.cboSemester.Location = new System.Drawing.Point(216, 12);
            this.cboSemester.Name = "cboSemester";
            this.cboSemester.Size = new System.Drawing.Size(65, 25);
            this.cboSemester.TabIndex = 13;
            // 
            // ExportDoctSetup
            // 
            this.ExportDoctSetup.AutoSize = true;
            this.ExportDoctSetup.BackColor = System.Drawing.Color.Transparent;
            this.ExportDoctSetup.Location = new System.Drawing.Point(164, 12);
            this.ExportDoctSetup.Name = "ExportDoctSetup";
            this.ExportDoctSetup.Size = new System.Drawing.Size(99, 17);
            this.ExportDoctSetup.TabIndex = 14;
            this.ExportDoctSetup.TabStop = true;
            this.ExportDoctSetup.Text = "樣板與列印設定";
            this.ExportDoctSetup.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ExportDoctSetup_LinkClicked);
            // 
            // gpDoc
            // 
            this.gpDoc.BackColor = System.Drawing.Color.Transparent;
            this.gpDoc.CanvasColor = System.Drawing.SystemColors.Control;
            this.gpDoc.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.gpDoc.Controls.Add(this.checkExportDoc);
            this.gpDoc.Controls.Add(this.ExportDoctSetup);
            this.gpDoc.Location = new System.Drawing.Point(9, 499);
            this.gpDoc.Name = "gpDoc";
            this.gpDoc.Size = new System.Drawing.Size(346, 67);
            // 
            // 
            // 
            this.gpDoc.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.gpDoc.Style.BackColorGradientAngle = 90;
            this.gpDoc.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.gpDoc.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDoc.Style.BorderBottomWidth = 1;
            this.gpDoc.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.gpDoc.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDoc.Style.BorderLeftWidth = 1;
            this.gpDoc.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDoc.Style.BorderRightWidth = 1;
            this.gpDoc.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpDoc.Style.BorderTopWidth = 1;
            this.gpDoc.Style.Class = "";
            this.gpDoc.Style.CornerDiameter = 4;
            this.gpDoc.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.gpDoc.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.gpDoc.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.gpDoc.StyleMouseDown.Class = "";
            this.gpDoc.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.gpDoc.StyleMouseOver.Class = "";
            this.gpDoc.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.gpDoc.TabIndex = 16;
            this.gpDoc.Text = "未達畢業標準通知單";
            // 
            // checkExportDoc
            // 
            this.checkExportDoc.AutoSize = true;
            this.checkExportDoc.BackColor = System.Drawing.Color.Transparent;
            this.checkExportDoc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.checkExportDoc.Location = new System.Drawing.Point(51, 11);
            this.checkExportDoc.Name = "checkExportDoc";
            this.checkExportDoc.Size = new System.Drawing.Size(92, 21);
            this.checkExportDoc.TabIndex = 16;
            this.checkExportDoc.Text = "產生通知單";
            this.checkExportDoc.UseVisualStyleBackColor = true;
            // 
            // GraduationPredictReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(364, 607);
            this.Controls.Add(this.gpDoc);
            this.Controls.Add(this.cboSemester);
            this.Controls.Add(this.cboSchoolYear);
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
            this.gpDoc.ResumeLayout(false);
            this.gpDoc.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckBox chCondition1;
        private System.Windows.Forms.CheckBox chCondition2;
        private System.Windows.Forms.CheckBox chCondition3;
        private System.Windows.Forms.CheckBox chCondition4;
        private System.Windows.Forms.CheckBox chCondition5;
        private System.Windows.Forms.CheckBox chCondition6;
        private DevComponents.DotNetBar.Controls.GroupPanel gpScore;
        private DevComponents.DotNetBar.Controls.GroupPanel gpDaily;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSchoolYear;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSemester;
        private System.Windows.Forms.CheckBox chCondition3b;
        private System.Windows.Forms.CheckBox chCondition4b;
        private System.Windows.Forms.CheckBox chCondition5b;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.LinkLabel ExportDoctSetup;
        private DevComponents.DotNetBar.Controls.GroupPanel gpDoc;
        private System.Windows.Forms.CheckBox checkExportDoc;
        private System.Windows.Forms.CheckBox chConditionGr1;
    }
}