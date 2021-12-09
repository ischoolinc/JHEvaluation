using FISCA.Presentation.Controls;
using Framework;
using JHSchool;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace HsinChu.JHEvaluation.ConfigControls.Ribbon
{
    public partial class ScoreValueManager : BaseForm
    {
        private bool _has_deleted;
        private ErrorProvider errorProvider1;

        public ScoreValueManager()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ResetDirty()
        {
            foreach (DataGridViewRow row in dgvScoreConfig.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                    cell.Tag = cell.Value;
            }
            //cboScoreSource.Tag = cboScoreSource.SelectedIndex;
            //txtStartTime.Tag = txtStartTime.Text;
            //txtEndTime.Tag = txtEndTime.Text;

            _has_deleted = false;
            errorProvider1.Clear();
        }

        private bool IsDirty()
        {
            bool dirty = false;
            foreach (DataGridViewRow row in dgvScoreConfig.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                    dirty = dirty || (cell.Tag + string.Empty != cell.Value + string.Empty);
            }

            return dirty || _has_deleted;
        }

        private void dataview_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            _has_deleted = true;
        }

        private void dgvScoreConfig_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }

        private void dgvScoreConfig_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgvScoreConfig.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgvScoreConfig_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && dgvScoreConfig.SelectedCells.Count == 1)
            {
                dgvScoreConfig.BeginEdit(true);
                if (dgvScoreConfig.CurrentCell != null && dgvScoreConfig.CurrentCell.GetType().ToString() == "System.Windows.Forms.DataGridViewComboBoxCell")
                    (dgvScoreConfig.EditingControl as ComboBox).DroppedDown = true;  //自動拉下清單
            }
        }

        private void dgvScoreConfig_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dgvScoreConfig.CurrentCell = null;
            //dgvScoreConfig.Rows[e.RowIndex].Selected = true;
        }

        private void dgvScoreConfig_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dgvScoreConfig.CurrentCell = null;
        }

        private void dgvScoreConfig_MouseClick(object sender, MouseEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            DataGridView.HitTestInfo hit = dgv.HitTest(e.X, e.Y);

            if (hit.Type == DataGridViewHitTestType.TopLeftHeader)
            {
                dgvScoreConfig.CurrentCell = null;
                //dgvScoreConfig.SelectAll();
            }
        }

        private void dgvScoreConfig_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgvScoreConfig.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgvScoreConfig_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                if (dgvScoreConfig.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "True")
                {
                    dgvScoreConfig.Rows[e.RowIndex].Cells[2].ReadOnly = false;
                }
                else
                {
                    dgvScoreConfig.Rows[e.RowIndex].Cells[2].Value = string.Empty;
                    dgvScoreConfig.Rows[e.RowIndex].Cells[2].ReadOnly = true;
                }
            }
        }

        private void languageChange(Object sender, InputLanguageChangedEventArgs e)
        {
            dgvScoreConfig.ImeMode = ImeMode.OnHalf;
            dgvScoreConfig.ImeMode = ImeMode.Off;
        }

        private void dgvScoreConfig_EditingControlShowing(Object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            dgvScoreConfig.ImeMode = ImeMode.OnHalf;
            dgvScoreConfig.ImeMode = ImeMode.Off;
        }

        private void ScoreValueManager_Load(object sender, EventArgs e)
        {
            errorProvider1 = new ErrorProvider();

            //  DataGridView 事件
            this.dgvScoreConfig.DataError += new DataGridViewDataErrorEventHandler(dgvScoreConfig_DataError);
            this.dgvScoreConfig.CurrentCellDirtyStateChanged += new EventHandler(dgvScoreConfig_CurrentCellDirtyStateChanged);
            this.dgvScoreConfig.CellEnter += new DataGridViewCellEventHandler(dgvScoreConfig_CellEnter);
            this.dgvScoreConfig.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvScoreConfig_ColumnHeaderMouseClick);
            this.dgvScoreConfig.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvScoreConfig_RowHeaderMouseClick);
            this.dgvScoreConfig.MouseClick += new System.Windows.Forms.MouseEventHandler(this.dgvScoreConfig_MouseClick);
            dgvScoreConfig.CellContentClick += new DataGridViewCellEventHandler(dgvScoreConfig_CellContentClick);
            dgvScoreConfig.CellValueChanged += new DataGridViewCellEventHandler(dgvScoreConfig_CellValueChanged);
            dgvScoreConfig.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(dgvScoreConfig_EditingControlShowing);
            this.InputLanguageChanged += new InputLanguageChangedEventHandler(languageChange);

            //dgvScoreConfig.Columns[1].ReadOnly = true;
            ///  <Setting >
            ///      <UseText>使用文字</UseText>
            ///      <AllowCalculation>是否計算成績</AllowCalculation>
            ///      <Score>計算分數</Score>
            ///      <Active>是否使用</Active>
            ///      <UseValue>替代分數</UseValue>
            /// </Setting>
            ConfigData cd = School.Configuration["評量成績缺考暨免試設定"];
            if (!string.IsNullOrEmpty(cd["評量成績缺考暨免試設定"]))
            {
                XmlElement element = XmlHelper.LoadXml(cd["評量成績缺考暨免試設定"]);

                foreach (XmlElement each in element.SelectNodes("Setting"))
                {
                    List<object> sources = new List<object>();

                    var UseText = each.SelectSingleNode("UseText").InnerText;
                    var AllowCalculation = each.SelectSingleNode("AllowCalculation").InnerText;
                    var Score = each.SelectSingleNode("Score").InnerText;
                    var Active = each.SelectSingleNode("Active").InnerText;
                    var UseValue = each.SelectSingleNode("UseValue").InnerText;
                    var ReportValue = each.SelectSingleNode("ReportValue").InnerText;

                    sources.Add(UseText);
                    sources.Add(AllowCalculation);
                    sources.Add(Score);
                    sources.Add(Active);
                    sources.Add(UseValue);
                    sources.Add(ReportValue);

                    int idx = this.dgvScoreConfig.Rows.Add(sources.ToArray());
                    if (AllowCalculation.ToLower() == "true")
                    {
                        this.dgvScoreConfig.Rows[idx].Cells[2].ReadOnly = false;
                    }
                    else
                    {
                        this.dgvScoreConfig.Rows[idx].Cells[2].ReadOnly = true;
                    }
                    this.dgvScoreConfig.Rows[idx].Cells[0].Tag = UseText;
                    this.dgvScoreConfig.Rows[idx].Cells[1].Tag = AllowCalculation;
                    this.dgvScoreConfig.Rows[idx].Cells[2].Tag = Score;
                    this.dgvScoreConfig.Rows[idx].Cells[3].Tag = Active;
                    this.dgvScoreConfig.Rows[idx].Cells[4].Tag = UseValue;
                    this.dgvScoreConfig.Rows[idx].Cells[6].Tag = ReportValue;
                }
            } 
            else
            {
                {
                    List<object> sources = new List<object>();

                    sources.Add("缺");
                    sources.Add(true);
                    sources.Add(0);
                    sources.Add(false);
                    sources.Add(int.MinValue);
                    sources.Add("0(缺)");

                    int idx = this.dgvScoreConfig.Rows.Add(sources.ToArray());
                    this.dgvScoreConfig.Rows[idx].Cells[0].Tag = "缺";
                    this.dgvScoreConfig.Rows[idx].Cells[1].Tag = true;
                    this.dgvScoreConfig.Rows[idx].Cells[2].Tag = 0;
                    this.dgvScoreConfig.Rows[idx].Cells[2].ReadOnly = false;
                    this.dgvScoreConfig.Rows[idx].Cells[3].Tag = false;
                    this.dgvScoreConfig.Rows[idx].Cells[4].Tag = int.MinValue;
                    this.dgvScoreConfig.Rows[idx].Cells[6].Tag = "0(缺)";
                }

                {
                    List<object> sources = new List<object>();

                    sources.Add("免");
                    sources.Add(false);
                    sources.Add("");
                    sources.Add(false);
                    sources.Add(int.MinValue + 1);
                    sources.Add("免");

                    int idx = this.dgvScoreConfig.Rows.Add(sources.ToArray());
                    this.dgvScoreConfig.Rows[idx].Cells[0].Tag = "免";
                    this.dgvScoreConfig.Rows[idx].Cells[1].Tag = false;
                    this.dgvScoreConfig.Rows[idx].Cells[2].Tag = ""; 
                    this.dgvScoreConfig.Rows[idx].Cells[2].ReadOnly = true;
                    this.dgvScoreConfig.Rows[idx].Cells[3].Tag = false;
                    this.dgvScoreConfig.Rows[idx].Cells[4].Tag = (int.MinValue + 1);
                    this.dgvScoreConfig.Rows[idx].Cells[6].Tag = "免";
                }

                cd["評量成績缺考暨免試設定"] = "<Settings><Setting>" +
                                "<UseText>缺</UseText>" +
                                "<AllowCalculation>true</AllowCalculation>" +
                                "<Score>0</Score>" +
                                "<Active>false</Active>" +
                                "<UseValue>" + int.MinValue + "</UseValue>" +
                                "<ReportValue>0(缺)</ReportValue>" +
                            "</Setting><Setting>" +
                                "<UseText>免</UseText>" +
                                "<AllowCalculation>false</AllowCalculation>" +
                                "<Score></Score>" +
                                "<Active>false</Active>" +
                                "<UseValue>" + (int.MinValue + 1) + "</UseValue>" +
                                "<ReportValue>免</ReportValue>" +
                            "</Setting></Settings>";

                cd.Save();
            }
        }

        private bool IsValid()
        {
            dgvScoreConfig.EndEdit();
            errorProvider1.Clear();

            foreach (DataGridViewRow row in dgvScoreConfig.Rows)
            {
                if (row.IsNewRow)
                    continue;

                row.Cells[0].ErrorText = "";
                row.Cells[1].ErrorText = "";
                row.Cells[2].ErrorText = "";
                row.Cells[3].ErrorText = "";

                if (string.IsNullOrEmpty((row.Cells[0].Value + "").Trim()))
                {
                    row.Cells[0].ErrorText = "必填";
                }
                decimal score;
                if (bool.Parse(row.Cells[1].Value + "") && !decimal.TryParse((row.Cells[2].Value + "").Trim(), out score))
                {
                    row.Cells[2].ErrorText = "僅允許數字";
                }
                if (bool.Parse(row.Cells[1].Value + "") && string.IsNullOrEmpty((row.Cells[2].Value + "").Trim()))
                {
                    row.Cells[2].ErrorText = "勾選「是否計算成績」時，此欄必填";
                }
            }
            if ((!string.IsNullOrEmpty((dgvScoreConfig.Rows[0].Cells[0].Value + "").Trim())) && ((!string.IsNullOrEmpty((dgvScoreConfig.Rows[1].Cells[0].Value + "").Trim())) && (dgvScoreConfig.Rows[0].Cells[0].Value + "").Trim() == (dgvScoreConfig.Rows[1].Cells[0].Value + "").Trim()))
            {
                dgvScoreConfig.Rows[0].Cells[0].ErrorText = "項目不可重複";
                dgvScoreConfig.Rows[1].Cells[0].ErrorText = "項目不可重複";
            }
            if (!string.IsNullOrEmpty(errorProvider1.GetError(dgvScoreConfig)))
            {
                return false;
            }
            foreach (DataGridViewRow row in dgvScoreConfig.Rows)
            {
                if (row.IsNewRow)
                    continue;

                if (!string.IsNullOrEmpty(row.Cells[0].ErrorText) || !string.IsNullOrEmpty(row.Cells[1].ErrorText) || !string.IsNullOrEmpty(row.Cells[2].ErrorText) || !string.IsNullOrEmpty(row.Cells[3].ErrorText))
                {
                    return false;
                }
            }

            return true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            dgvScoreConfig.EndEdit();

            if (!IsValid())
            {
                MessageBox.Show("請修正錯誤再儲存。");
                return;
            }

            ConfigData cd = School.Configuration["評量成績缺考暨免試設定"];
            StringBuilder sb = new StringBuilder("<Settings>");
            foreach (DataGridViewRow row in dgvScoreConfig.Rows)
            {
                if (row.IsNewRow) continue;

                sb.Append("<Setting>");
                sb.Append(string.Format("<UseText>{0}</UseText>", row.Cells["colUseText"].Value + ""));
                sb.Append(string.Format("<AllowCalculation>{0}</AllowCalculation>", row.Cells["colAllowCalculation"].Value + ""));
                sb.Append(string.Format("<Score>{0}</Score>", row.Cells["colScore"].Value + ""));
                sb.Append(string.Format("<Active>{0}</Active>", row.Cells["colActive"].Value + ""));
                sb.Append(string.Format("<UseValue>{0}</UseValue>", row.Cells["colUseValue"].Value + ""));
                sb.Append(string.Format("<ReportValue>{0}</ReportValue>", row.Cells["colReportValue"].Value + ""));
                sb.Append("</Setting>");
            }
            sb.Append("</Settings>");

            cd["評量成績缺考暨免試設定"] = sb.ToString();
            cd.Save();

            MessageBox.Show("儲存成功");
        }
    }
}
