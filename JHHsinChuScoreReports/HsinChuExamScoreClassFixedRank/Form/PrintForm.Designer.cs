﻿namespace HsinChuExamScoreClassFixedRank.Form
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
            this.cboExam = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cboSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cboSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
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
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.labelX5 = new DevComponents.DotNetBar.LabelX();
            this.cboParseNumber = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cboRefExam = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX6 = new DevComponents.DotNetBar.LabelX();
            this.labelX7 = new DevComponents.DotNetBar.LabelX();
            this.cboRefSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cboRefSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX8 = new DevComponents.DotNetBar.LabelX();
            this.checkOrder = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.labelX9 = new DevComponents.DotNetBar.LabelX();
            this.cboOrder = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.comboItem1 = new DevComponents.Editors.ComboItem();
            this.comboItem2 = new DevComponents.Editors.ComboItem();
            this.comboItem3 = new DevComponents.Editors.ComboItem();
            this.comboItem4 = new DevComponents.Editors.ComboItem();
            this.groupPanel1.SuspendLayout();
            this.groupPanel2.SuspendLayout();
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
            this.lnkCopyConfig.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkCopyConfig_LinkClicked);
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
            // cboExam
            // 
            this.cboExam.DisplayMember = "Text";
            this.cboExam.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboExam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboExam.FormattingEnabled = true;
            this.cboExam.ItemHeight = 19;
            this.cboExam.Location = new System.Drawing.Point(375, 67);
            this.cboExam.Name = "cboExam";
            this.cboExam.Size = new System.Drawing.Size(189, 25);
            this.cboExam.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboExam.TabIndex = 3;
            this.cboExam.SelectedIndexChanged += new System.EventHandler(this.cboExam_SelectedIndexChanged);
            // 
            // cboSemester
            // 
            this.cboSemester.DisplayMember = "Text";
            this.cboSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSemester.FormattingEnabled = true;
            this.cboSemester.ItemHeight = 19;
            this.cboSemester.Location = new System.Drawing.Point(179, 67);
            this.cboSemester.Name = "cboSemester";
            this.cboSemester.Size = new System.Drawing.Size(48, 25);
            this.cboSemester.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboSemester.TabIndex = 2;
            // 
            // cboSchoolYear
            // 
            this.cboSchoolYear.DisplayMember = "Text";
            this.cboSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSchoolYear.FormattingEnabled = true;
            this.cboSchoolYear.ItemHeight = 19;
            this.cboSchoolYear.Location = new System.Drawing.Point(71, 67);
            this.cboSchoolYear.Name = "cboSchoolYear";
            this.cboSchoolYear.Size = new System.Drawing.Size(62, 25);
            this.cboSchoolYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboSchoolYear.TabIndex = 1;
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(311, 69);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(60, 21);
            this.labelX3.TabIndex = 32;
            this.labelX3.Text = "試別名稱";
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
            this.labelX2.Location = new System.Drawing.Point(139, 69);
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
            this.labelX1.Location = new System.Drawing.Point(20, 69);
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
            this.lnkViewMapColumns.Location = new System.Drawing.Point(195, 473);
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
            this.btnSaveConfig.Location = new System.Drawing.Point(329, 470);
            this.btnSaveConfig.Name = "btnSaveConfig";
            this.btnSaveConfig.Size = new System.Drawing.Size(75, 23);
            this.btnSaveConfig.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSaveConfig.TabIndex = 6;
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
            this.lnkViewTemplate.Location = new System.Drawing.Point(16, 473);
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
            this.lnkChangeTemplate.Location = new System.Drawing.Point(105, 473);
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
            this.btnPrint.Location = new System.Drawing.Point(414, 470);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(67, 23);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 7;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(487, 470);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(67, 23);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "離開";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.lvSubject);
            this.groupPanel1.Location = new System.Drawing.Point(19, 180);
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
            this.lvSubject.HideSelection = false;
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
            this.groupPanel2.Location = new System.Drawing.Point(300, 180);
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
            this.lvDomain.HideSelection = false;
            this.lvDomain.Location = new System.Drawing.Point(3, 6);
            this.lvDomain.Name = "lvDomain";
            this.lvDomain.Size = new System.Drawing.Size(253, 216);
            this.lvDomain.TabIndex = 0;
            this.lvDomain.UseCompatibleStateImageBehavior = false;
            this.lvDomain.View = System.Windows.Forms.View.List;
            // 
            // labelX4
            // 
            this.labelX4.AutoSize = true;
            this.labelX4.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.Class = "";
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Location = new System.Drawing.Point(20, 105);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(127, 21);
            this.labelX4.TabIndex = 42;
            this.labelX4.Text = "平均計算至小數點後";
            // 
            // labelX5
            // 
            this.labelX5.AutoSize = true;
            this.labelX5.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX5.BackgroundStyle.Class = "";
            this.labelX5.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX5.Location = new System.Drawing.Point(204, 105);
            this.labelX5.Name = "labelX5";
            this.labelX5.Size = new System.Drawing.Size(74, 21);
            this.labelX5.TabIndex = 43;
            this.labelX5.Text = "位四捨五入";
            // 
            // cboParseNumber
            // 
            this.cboParseNumber.DisplayMember = "Text";
            this.cboParseNumber.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboParseNumber.FormattingEnabled = true;
            this.cboParseNumber.ItemHeight = 19;
            this.cboParseNumber.Location = new System.Drawing.Point(148, 103);
            this.cboParseNumber.Name = "cboParseNumber";
            this.cboParseNumber.Size = new System.Drawing.Size(40, 25);
            this.cboParseNumber.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboParseNumber.TabIndex = 5;
            this.cboParseNumber.Text = "2";
            // 
            // cboRefExam
            // 
            this.cboRefExam.DisplayMember = "Text";
            this.cboRefExam.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboRefExam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRefExam.FormattingEnabled = true;
            this.cboRefExam.ItemHeight = 19;
            this.cboRefExam.Location = new System.Drawing.Point(375, 137);
            this.cboRefExam.Name = "cboRefExam";
            this.cboRefExam.Size = new System.Drawing.Size(189, 25);
            this.cboRefExam.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboRefExam.TabIndex = 4;
            this.cboRefExam.SelectedIndexChanged += new System.EventHandler(this.cboRefExam_SelectedIndexChanged);
            // 
            // labelX6
            // 
            this.labelX6.AutoSize = true;
            this.labelX6.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX6.BackgroundStyle.Class = "";
            this.labelX6.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX6.Location = new System.Drawing.Point(311, 139);
            this.labelX6.Name = "labelX6";
            this.labelX6.Size = new System.Drawing.Size(60, 21);
            this.labelX6.TabIndex = 45;
            this.labelX6.Text = "參考試別";
            // 
            // labelX7
            // 
            this.labelX7.AutoSize = true;
            this.labelX7.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX7.BackgroundStyle.Class = "";
            this.labelX7.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX7.Location = new System.Drawing.Point(311, 104);
            this.labelX7.Name = "labelX7";
            this.labelX7.Size = new System.Drawing.Size(87, 21);
            this.labelX7.TabIndex = 46;
            this.labelX7.Text = "參考學年學期";
            // 
            // cboRefSchoolYear
            // 
            this.cboRefSchoolYear.DisplayMember = "Text";
            this.cboRefSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboRefSchoolYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRefSchoolYear.FormattingEnabled = true;
            this.cboRefSchoolYear.ItemHeight = 19;
            this.cboRefSchoolYear.Location = new System.Drawing.Point(404, 103);
            this.cboRefSchoolYear.Name = "cboRefSchoolYear";
            this.cboRefSchoolYear.Size = new System.Drawing.Size(83, 25);
            this.cboRefSchoolYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboRefSchoolYear.TabIndex = 47;
            // 
            // cboRefSemester
            // 
            this.cboRefSemester.DisplayMember = "Text";
            this.cboRefSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboRefSemester.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRefSemester.FormattingEnabled = true;
            this.cboRefSemester.ItemHeight = 19;
            this.cboRefSemester.Location = new System.Drawing.Point(494, 103);
            this.cboRefSemester.Name = "cboRefSemester";
            this.cboRefSemester.Size = new System.Drawing.Size(71, 25);
            this.cboRefSemester.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboRefSemester.TabIndex = 48;
            // 
            // labelX8
            // 
            this.labelX8.AutoSize = true;
            this.labelX8.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX8.BackgroundStyle.Class = "";
            this.labelX8.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX8.ForeColor = System.Drawing.Color.Red;
            this.labelX8.Location = new System.Drawing.Point(19, 444);
            this.labelX8.Name = "labelX8";
            this.labelX8.Size = new System.Drawing.Size(238, 21);
            this.labelX8.TabIndex = 49;
            this.labelX8.Text = "若勾選領域， 總分將以領域成績計算。";
            // 
            // checkOrder
            // 
            this.checkOrder.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.checkOrder.BackgroundStyle.Class = "";
            this.checkOrder.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.checkOrder.Location = new System.Drawing.Point(19, 138);
            this.checkOrder.Name = "checkOrder";
            this.checkOrder.Size = new System.Drawing.Size(55, 23);
            this.checkOrder.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.checkOrder.TabIndex = 50;
            this.checkOrder.Text = "依照";
            this.checkOrder.CheckedChanged += new System.EventHandler(this.checkOrder_CheckedChanged);
            // 
            // labelX9
            // 
            this.labelX9.AutoSize = true;
            this.labelX9.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX9.BackgroundStyle.Class = "";
            this.labelX9.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX9.Location = new System.Drawing.Point(188, 139);
            this.labelX9.Name = "labelX9";
            this.labelX9.Size = new System.Drawing.Size(101, 21);
            this.labelX9.TabIndex = 51;
            this.labelX9.Text = "班排名排序學生";
            // 
            // cboOrder
            // 
            this.cboOrder.DisplayMember = "Text";
            this.cboOrder.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboOrder.Enabled = false;
            this.cboOrder.FormattingEnabled = true;
            this.cboOrder.ItemHeight = 19;
            this.cboOrder.Items.AddRange(new object[] {
            this.comboItem1,
            this.comboItem2,
            this.comboItem3,
            this.comboItem4});
            this.cboOrder.Location = new System.Drawing.Point(70, 137);
            this.cboOrder.Name = "cboOrder";
            this.cboOrder.Size = new System.Drawing.Size(118, 25);
            this.cboOrder.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboOrder.TabIndex = 52;
            // 
            // comboItem1
            // 
            this.comboItem1.Text = "定期_加權總分";
            // 
            // comboItem2
            // 
            this.comboItem2.Text = "定期_加權平均";
            // 
            // comboItem3
            // 
            this.comboItem3.Text = "定期_總分";
            // 
            // comboItem4
            // 
            this.comboItem4.Text = "定期_平均";
            // 
            // PrintForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(593, 523);
            this.Controls.Add(this.cboOrder);
            this.Controls.Add(this.labelX9);
            this.Controls.Add(this.checkOrder);
            this.Controls.Add(this.labelX8);
            this.Controls.Add(this.cboRefSemester);
            this.Controls.Add(this.cboRefSchoolYear);
            this.Controls.Add(this.labelX7);
            this.Controls.Add(this.cboRefExam);
            this.Controls.Add(this.labelX6);
            this.Controls.Add(this.cboParseNumber);
            this.Controls.Add(this.labelX5);
            this.Controls.Add(this.labelX4);
            this.Controls.Add(this.groupPanel2);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.lnkViewMapColumns);
            this.Controls.Add(this.btnSaveConfig);
            this.Controls.Add(this.lnkViewTemplate);
            this.Controls.Add(this.lnkChangeTemplate);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cboExam);
            this.Controls.Add(this.cboSemester);
            this.Controls.Add(this.cboSchoolYear);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.cboConfigure);
            this.Controls.Add(this.lnkDelConfig);
            this.Controls.Add(this.lnkCopyConfig);
            this.Controls.Add(this.labelX11);
            this.DoubleBuffered = true;
            this.Name = "PrintForm";
            this.Text = "班級評量成績通知單(固定排名)";
            this.Load += new System.EventHandler(this.PrintForm_Load);
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ComboBoxEx cboConfigure;
        private System.Windows.Forms.LinkLabel lnkDelConfig;
        private System.Windows.Forms.LinkLabel lnkCopyConfig;
        private DevComponents.DotNetBar.LabelX labelX11;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboExam;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSemester;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSchoolYear;
        private DevComponents.DotNetBar.LabelX labelX3;
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
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.DotNetBar.LabelX labelX5;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboParseNumber;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboRefExam;
        private DevComponents.DotNetBar.LabelX labelX6;
        private DevComponents.DotNetBar.LabelX labelX7;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboRefSchoolYear;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboRefSemester;
        private DevComponents.DotNetBar.LabelX labelX8;
        private DevComponents.DotNetBar.Controls.CheckBoxX checkOrder;
        private DevComponents.DotNetBar.LabelX labelX9;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboOrder;
        private DevComponents.Editors.ComboItem comboItem1;
        private DevComponents.Editors.ComboItem comboItem2;
        private DevComponents.Editors.ComboItem comboItem3;
        private DevComponents.Editors.ComboItem comboItem4;
    }
}