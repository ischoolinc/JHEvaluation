﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Framework;
using JHSchool.Evaluation.Editor;
using System.Threading;
using JHSchool.Data;

namespace JHSchool.Evaluation.CourseExtendControls.Ribbon.AssessmentSetupManagerControls
{
    internal partial class AssessmentNameEditor : FISCA.Presentation.Controls.BaseForm
    {
        private JHAssessmentSetupRecord _record;
        private string _orig_name;

        public AssessmentNameEditor()
        {
            InitializeComponent();

            _record = null;
            _orig_name = string.Empty;
            Text = "新增評分樣版";

            foreach (JHAssessmentSetupRecord item in JHAssessmentSetup.SelectAll())
                cboExistTemplates.Items.Add(item);
            //foreach (Template tpl in templates)
            //    cboExistTemplates.Items.Add(tpl);
            cboExistTemplates.SelectedIndex = 0;
        }

        public AssessmentNameEditor(JHAssessmentSetupRecord record)
        {
            InitializeComponent();

            _record = record;
            _orig_name = record.Name;
            txtTemplateName.Text = record.Name;
            txtTemplateName.SelectAll();

            Text = "重新命名評分樣版";
            lblCopyExist.Enabled = lblCopyExist.Visible = false;
            cboExistTemplates.Enabled = cboExistTemplates.Visible = false;
            Size = new Size(373, 104);
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtTemplateName.Text.Trim()))
            {
                MsgBox.Show("您必須輸入名稱。");
                DialogResult = DialogResult.None;
                return;
            }

            if (ContainsTemplateName(txtTemplateName.Text.Trim()) && txtTemplateName.Text.Trim() != _orig_name)
            {
                MsgBox.Show("名稱重覆，請選擇其他名稱。");
                DialogResult = DialogResult.None;
                txtTemplateName.SelectAll();
                return;
            }

            if (txtTemplateName.Text.Trim() != _orig_name)
            {
                //AssessmentSetupRecordEditor editor;
                //if (_record == null) //Insert
                //    editor = AssessmentSetup.Instance.AddAssessmentSetup();
                //else
                //    editor = _record.GetEditor();
                //editor.Name = txtTemplateName.Text.Trim();

                if (_record == null)
                {
                    JHAssessmentSetup.AfterInsert += new EventHandler<K12.Data.DataChangedEventArgs>(JHAssessmentSetup_AfterInsert);
                    _record = new JHAssessmentSetupRecord();
                    _record.Name = txtTemplateName.Text.Trim();
                    JHAssessmentSetup.Insert(_record);
                }
                else
                {
                    _record.Name = txtTemplateName.Text.Trim();
                    JHAssessmentSetup.Update(_record);
                }

                DialogResult = DialogResult.OK;
            }
        }

        private void JHAssessmentSetup_AfterInsert(object sender, K12.Data.DataChangedEventArgs e)
        {
            if (e.PrimaryKeys.Count > 0)
            {
                if (cboExistTemplates.Enabled == true && cboExistTemplates.SelectedItem is JHAssessmentSetupRecord)
                {
                    JHAssessmentSetupRecord source = cboExistTemplates.SelectedItem as JHAssessmentSetupRecord;
                    bool executeRequired = false;

                    List<JHAEIncludeRecord> aeincludes = new List<JHAEIncludeRecord>();
                    foreach (JHAEIncludeRecord each in JHAEInclude.SelectByAssessmentSetupID(source.ID))
                    {
                        JHAEIncludeRecord aeNew = new JHAEIncludeRecord();
                        aeNew.RefAssessmentSetupID = e.PrimaryKeys[0];
                        aeNew.RefExamID = each.RefExamID;
                        aeNew.UseScore = each.UseScore;
                        aeNew.UseText = each.UseText;
                        aeNew.UseEffort = each.UseEffort;
                        aeNew.Weight = each.Weight;
                        aeNew.StartTime = each.StartTime;
                        aeNew.EndTime = each.EndTime;

                        aeincludes.Add(aeNew);
                        executeRequired = true;
                    }
                    if (executeRequired)
                        JHAEInclude.Insert(aeincludes);
                }
            }
            JHAssessmentSetup.AfterInsert -= new EventHandler<K12.Data.DataChangedEventArgs>(JHAssessmentSetup_AfterInsert);
        }

        private bool ContainsTemplateName(string p)
        {
            foreach (JHAssessmentSetupRecord item in cboExistTemplates.Items)
            {
                if (item.Name == p)
                    return true;
            }
            return false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}