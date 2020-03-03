namespace HsinChuSemesterClassFixedRank
{
    partial class PrintForm
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
            this.cboConfigure = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lnkDelConfig = new System.Windows.Forms.LinkLabel();
            this.lnkCopyConfig = new System.Windows.Forms.LinkLabel();
            this.labelX11 = new DevComponents.DotNetBar.LabelX();
            this.cboSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cboSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.lnkViewMapColumns = new System.Windows.Forms.LinkLabel();
            this.btnSaveConfig = new DevComponents.DotNetBar.ButtonX();
            this.lnkViewTemplate = new System.Windows.Forms.LinkLabel();
            this.lnkChangeTemplate = new System.Windows.Forms.LinkLabel();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.lvSubject = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.groupPanel2 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.lvDomain = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.groupPanel3 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.txtFailScoreMark = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtReScoreMark = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtNeeedReScoreMark = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.labelX5 = new DevComponents.DotNetBar.LabelX();
            this.groupPanel1.SuspendLayout();
            this.groupPanel2.SuspendLayout();
            this.groupPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // cboConfigure
            // 
            this.cboConfigure.DisplayMember = "Name";
            this.cboConfigure.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboConfigure.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboConfigure.FormattingEnabled = true;
            this.cboConfigure.ItemHeight = 19;
            this.cboConfigure.Location = new System.Drawing.Point(111, 21);
            this.cboConfigure.Name = "cboConfigure";
            this.cboConfigure.Size = new System.Drawing.Size(273, 25);
            this.cboConfigure.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboConfigure.TabIndex = 0;
            this.cboConfigure.SelectedIndexChanged += new System.EventHandler(this.cboConfigure_SelectedIndexChanged);
            // 
            // lnkDelConfig
            // 
            this.lnkDelConfig.AutoSize = true;
            this.lnkDelConfig.BackColor = System.Drawing.Color.Transparent;
            this.lnkDelConfig.Location = new System.Drawing.Point(481, 29);
            this.lnkDelConfig.Name = "lnkDelConfig";
            this.lnkDelConfig.Size = new System.Drawing.Size(73, 17);
            this.lnkDelConfig.TabIndex = 2;
            this.lnkDelConfig.TabStop = true;
            this.lnkDelConfig.Text = "刪除設定檔";
            this.lnkDelConfig.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDelConfig_LinkClicked);
            // 
            // lnkCopyConfig
            // 
            this.lnkCopyConfig.AutoSize = true;
            this.lnkCopyConfig.BackColor = System.Drawing.Color.Transparent;
            this.lnkCopyConfig.Location = new System.Drawing.Point(402, 29);
            this.lnkCopyConfig.Name = "lnkCopyConfig";
            this.lnkCopyConfig.Size = new System.Drawing.Size(73, 17);
            this.lnkCopyConfig.TabIndex = 1;
            this.lnkCopyConfig.TabStop = true;
            this.lnkCopyConfig.Text = "複製設定檔";
            // 
            // labelX11
            // 
            this.labelX11.AutoSize = true;
            this.labelX11.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX11.BackgroundStyle.Class = "";
            this.labelX11.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX11.Location = new System.Drawing.Point(20, 23);
            this.labelX11.Name = "labelX11";
            this.labelX11.Size = new System.Drawing.Size(87, 21);
            this.labelX11.TabIndex = 13;
            this.labelX11.Text = "樣板設定檔：";
            // 
            // cboSemester
            // 
            this.cboSemester.DisplayMember = "Text";
            this.cboSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSemester.FormattingEnabled = true;
            this.cboSemester.ItemHeight = 19;
            this.cboSemester.Location = new System.Drawing.Point(243, 62);
            this.cboSemester.Name = "cboSemester";
            this.cboSemester.Size = new System.Drawing.Size(48, 25);
            this.cboSemester.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboSemester.TabIndex = 4;
            // 
            // cboSchoolYear
            // 
            this.cboSchoolYear.DisplayMember = "Text";
            this.cboSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSchoolYear.FormattingEnabled = true;
            this.cboSchoolYear.ItemHeight = 19;
            this.cboSchoolYear.Location = new System.Drawing.Point(111, 62);
            this.cboSchoolYear.Name = "cboSchoolYear";
            this.cboSchoolYear.Size = new System.Drawing.Size(69, 25);
            this.cboSchoolYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboSchoolYear.TabIndex = 3;
            this.cboSchoolYear.SelectedIndexChanged += new System.EventHandler(this.cboSchoolYear_SelectedIndexChanged);
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
            this.labelX2.Location = new System.Drawing.Point(203, 64);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(34, 21);
            this.labelX2.TabIndex = 29;
            this.labelX2.Text = "學期";
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
            this.labelX1.Location = new System.Drawing.Point(60, 64);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(47, 21);
            this.labelX1.TabIndex = 31;
            this.labelX1.Text = "學年度";
            // 
            // lnkViewMapColumns
            // 
            this.lnkViewMapColumns.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkViewMapColumns.AutoSize = true;
            this.lnkViewMapColumns.BackColor = System.Drawing.Color.Transparent;
            this.lnkViewMapColumns.Location = new System.Drawing.Point(200, 467);
            this.lnkViewMapColumns.Name = "lnkViewMapColumns";
            this.lnkViewMapColumns.Size = new System.Drawing.Size(112, 17);
            this.lnkViewMapColumns.TabIndex = 9;
            this.lnkViewMapColumns.TabStop = true;
            this.lnkViewMapColumns.Text = "檢視合併欄位總表";
            this.lnkViewMapColumns.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkViewMapColumns_LinkClicked);
            // 
            // btnSaveConfig
            // 
            this.btnSaveConfig.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSaveConfig.BackColor = System.Drawing.Color.Transparent;
            this.btnSaveConfig.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSaveConfig.Enabled = false;
            this.btnSaveConfig.Location = new System.Drawing.Point(334, 464);
            this.btnSaveConfig.Name = "btnSaveConfig";
            this.btnSaveConfig.Size = new System.Drawing.Size(75, 23);
            this.btnSaveConfig.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSaveConfig.TabIndex = 10;
            this.btnSaveConfig.Text = "儲存設定";
            this.btnSaveConfig.Tooltip = "儲存當前的樣板設定。";
            this.btnSaveConfig.Click += new System.EventHandler(this.btnSaveConfig_Click);
            // 
            // lnkViewTemplate
            // 
            this.lnkViewTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkViewTemplate.AutoSize = true;
            this.lnkViewTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkViewTemplate.Location = new System.Drawing.Point(21, 467);
            this.lnkViewTemplate.Name = "lnkViewTemplate";
            this.lnkViewTemplate.Size = new System.Drawing.Size(86, 17);
            this.lnkViewTemplate.TabIndex = 7;
            this.lnkViewTemplate.TabStop = true;
            this.lnkViewTemplate.Text = "檢視套印樣板";
            this.lnkViewTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkViewTemplate_LinkClicked);
            // 
            // lnkChangeTemplate
            // 
            this.lnkChangeTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkChangeTemplate.AutoSize = true;
            this.lnkChangeTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkChangeTemplate.Location = new System.Drawing.Point(110, 467);
            this.lnkChangeTemplate.Name = "lnkChangeTemplate";
            this.lnkChangeTemplate.Size = new System.Drawing.Size(86, 17);
            this.lnkChangeTemplate.TabIndex = 8;
            this.lnkChangeTemplate.TabStop = true;
            this.lnkChangeTemplate.Text = "變更套印樣板";
            this.lnkChangeTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkChangeTemplate_LinkClicked);
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnPrint.Enabled = false;
            this.btnPrint.Location = new System.Drawing.Point(419, 464);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(67, 23);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 11;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(492, 464);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(67, 23);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "離開";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.lvSubject);
            this.groupPanel1.Location = new System.Drawing.Point(24, 184);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(265, 256);
            // 
            // 
            // 
            this.groupPanel1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel1.Style.BackColorGradientAngle = 90;
            this.groupPanel1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderBottomWidth = 1;
            this.groupPanel1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderLeftWidth = 1;
            this.groupPanel1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderRightWidth = 1;
            this.groupPanel1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderTopWidth = 1;
            this.groupPanel1.Style.Class = "";
            this.groupPanel1.Style.CornerDiameter = 4;
            this.groupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseDown.Class = "";
            this.groupPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseOver.Class = "";
            this.groupPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel1.TabIndex = 40;
            this.groupPanel1.Text = "科目";
            // 
            // lvSubject
            // 
            // 
            // 
            // 
            this.lvSubject.Border.Class = "ListViewBorder";
            this.lvSubject.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvSubject.CheckBoxes = true;
            this.lvSubject.Location = new System.Drawing.Point(3, 7);
            this.lvSubject.Name = "lvSubject";
            this.lvSubject.Size = new System.Drawing.Size(253, 216);
            this.lvSubject.TabIndex = 0;
            this.lvSubject.UseCompatibleStateImageBehavior = false;
            this.lvSubject.View = System.Windows.Forms.View.List;
            // 
            // groupPanel2
            // 
            this.groupPanel2.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel2.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel2.Controls.Add(this.lvDomain);
            this.groupPanel2.Location = new System.Drawing.Point(305, 184);
            this.groupPanel2.Name = "groupPanel2";
            this.groupPanel2.Size = new System.Drawing.Size(265, 256);
            // 
            // 
            // 
            this.groupPanel2.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel2.Style.BackColorGradientAngle = 90;
            this.groupPanel2.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel2.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderBottomWidth = 1;
            this.groupPanel2.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel2.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderLeftWidth = 1;
            this.groupPanel2.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderRightWidth = 1;
            this.groupPanel2.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderTopWidth = 1;
            this.groupPanel2.Style.Class = "";
            this.groupPanel2.Style.CornerDiameter = 4;
            this.groupPanel2.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel2.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel2.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel2.StyleMouseDown.Class = "";
            this.groupPanel2.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel2.StyleMouseOver.Class = "";
            this.groupPanel2.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel2.TabIndex = 41;
            this.groupPanel2.Text = "領域";
            // 
            // lvDomain
            // 
            // 
            // 
            // 
            this.lvDomain.Border.Class = "ListViewBorder";
            this.lvDomain.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvDomain.CheckBoxes = true;
            this.lvDomain.Location = new System.Drawing.Point(3, 6);
            this.lvDomain.Name = "lvDomain";
            this.lvDomain.Size = new System.Drawing.Size(253, 216);
            this.lvDomain.TabIndex = 0;
            this.lvDomain.UseCompatibleStateImageBehavior = false;
            this.lvDomain.View = System.Windows.Forms.View.List;
            // 
            // groupPanel3
            // 
            this.groupPanel3.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel3.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel3.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel3.Controls.Add(this.txtFailScoreMark);
            this.groupPanel3.Controls.Add(this.txtReScoreMark);
            this.groupPanel3.Controls.Add(this.txtNeeedReScoreMark);
            this.groupPanel3.Controls.Add(this.labelX3);
            this.groupPanel3.Controls.Add(this.labelX4);
            this.groupPanel3.Controls.Add(this.labelX5);
            this.groupPanel3.Location = new System.Drawing.Point(24, 102);
            this.groupPanel3.Name = "groupPanel3";
            this.groupPanel3.Size = new System.Drawing.Size(542, 71);
            // 
            // 
            // 
            this.groupPanel3.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel3.Style.BackColorGradientAngle = 90;
            this.groupPanel3.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel3.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderBottomWidth = 1;
            this.groupPanel3.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel3.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderLeftWidth = 1;
            this.groupPanel3.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderRightWidth = 1;
            this.groupPanel3.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderTopWidth = 1;
            this.groupPanel3.Style.Class = "";
            this.groupPanel3.Style.CornerDiameter = 4;
            this.groupPanel3.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel3.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel3.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel3.StyleMouseDown.Class = "";
            this.groupPanel3.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel3.StyleMouseOver.Class = "";
            this.groupPanel3.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel3.TabIndex = 48;
            this.groupPanel3.Text = "成績未取得學分狀態標示設定";
            // 
            // txtFailScoreMark
            // 
            // 
            // 
            // 
            this.txtFailScoreMark.Border.Class = "TextBoxBorder";
            this.txtFailScoreMark.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtFailScoreMark.Location = new System.Drawing.Point(441, 9);
            this.txtFailScoreMark.Name = "txtFailScoreMark";
            this.txtFailScoreMark.Size = new System.Drawing.Size(48, 25);
            this.txtFailScoreMark.TabIndex = 2;
            // 
            // txtReScoreMark
            // 
            // 
            // 
            // 
            this.txtReScoreMark.Border.Class = "TextBoxBorder";
            this.txtReScoreMark.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtReScoreMark.Location = new System.Drawing.Point(276, 9);
            this.txtReScoreMark.Name = "txtReScoreMark";
            this.txtReScoreMark.Size = new System.Drawing.Size(48, 25);
            this.txtReScoreMark.TabIndex = 1;
            // 
            // txtNeeedReScoreMark
            // 
            // 
            // 
            // 
            this.txtNeeedReScoreMark.Border.Class = "TextBoxBorder";
            this.txtNeeedReScoreMark.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtNeeedReScoreMark.Location = new System.Drawing.Point(124, 9);
            this.txtNeeedReScoreMark.Name = "txtNeeedReScoreMark";
            this.txtNeeedReScoreMark.Size = new System.Drawing.Size(48, 25);
            this.txtNeeedReScoreMark.TabIndex = 0;
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(335, 11);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(101, 21);
            this.labelX3.TabIndex = 2;
            this.labelX3.Text = "不及格成績加註";
            // 
            // labelX4
            // 
            this.labelX4.AutoSize = true;
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.Class = "";
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Location = new System.Drawing.Point(188, 11);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(87, 21);
            this.labelX4.TabIndex = 1;
            this.labelX4.Text = "補考成績加註";
            // 
            // labelX5
            // 
            this.labelX5.AutoSize = true;
            // 
            // 
            // 
            this.labelX5.BackgroundStyle.Class = "";
            this.labelX5.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX5.Location = new System.Drawing.Point(21, 11);
            this.labelX5.Name = "labelX5";
            this.labelX5.Size = new System.Drawing.Size(101, 21);
            this.labelX5.TabIndex = 0;
            this.labelX5.Text = "需補考成績加註";
            // 
            // PrintForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(593, 500);
            this.Controls.Add(this.groupPanel3);
            this.Controls.Add(this.groupPanel2);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.lnkViewMapColumns);
            this.Controls.Add(this.btnSaveConfig);
            this.Controls.Add(this.lnkViewTemplate);
            this.Controls.Add(this.lnkChangeTemplate);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cboSemester);
            this.Controls.Add(this.cboSchoolYear);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.cboConfigure);
            this.Controls.Add(this.lnkDelConfig);
            this.Controls.Add(this.lnkCopyConfig);
            this.Controls.Add(this.labelX11);
            this.DoubleBuffered = true;
            this.Name = "PrintForm";
            this.Text = "班級學期成績通知單(固定排名)";
            this.Load += new System.EventHandler(this.PrintForm_Load);
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel2.ResumeLayout(false);
            this.groupPanel3.ResumeLayout(false);
            this.groupPanel3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ComboBoxEx cboConfigure;
        private System.Windows.Forms.LinkLabel lnkDelConfig;
        private System.Windows.Forms.LinkLabel lnkCopyConfig;
        private DevComponents.DotNetBar.LabelX labelX11;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSemester;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSchoolYear;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.LabelX labelX1;
        private System.Windows.Forms.LinkLabel lnkViewMapColumns;
        private DevComponents.DotNetBar.ButtonX btnSaveConfig;
        private System.Windows.Forms.LinkLabel lnkViewTemplate;
        private System.Windows.Forms.LinkLabel lnkChangeTemplate;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel2;
        private DevComponents.DotNetBar.Controls.ListViewEx lvSubject;
        private DevComponents.DotNetBar.Controls.ListViewEx lvDomain;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel3;
        private DevComponents.DotNetBar.Controls.TextBoxX txtFailScoreMark;
        private DevComponents.DotNetBar.Controls.TextBoxX txtReScoreMark;
        private DevComponents.DotNetBar.Controls.TextBoxX txtNeeedReScoreMark;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.DotNetBar.LabelX labelX5;
    }
}