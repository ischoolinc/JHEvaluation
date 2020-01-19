using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using JHSchool.Data;
using System.IO;
using HsinChu.JHEvaluation.Data;
using Aspose.Words;
using JHSchool.Evaluation.Calculation;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using FISCA.Data;


namespace HsinChuSemesterClassFixedRank
{
    public partial class PrintForm : BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        List<string> _ClassIDList = new List<string>();


        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();
        public Configure _Configure { get; private set; }

        // 畫面上所勾選領域科目        
        Dictionary<string, List<string>> SelDomainSubjectDict = new Dictionary<string, List<string>>();

        BackgroundWorker bgWorkerLoadTemplate;
        BackgroundWorker bgWorkerReport;

        // 紀錄樣板設定
        List<DAO.UDT_ScoreConfig> _UDTConfigList;


        public PrintForm()
        {
            InitializeComponent();
            bgWorkerLoadTemplate = new BackgroundWorker();
            bgWorkerLoadTemplate.DoWork += BgWorkerLoadTemplate_DoWork;
            bgWorkerLoadTemplate.RunWorkerCompleted += BgWorkerLoadTemplate_RunWorkerCompleted;
            bgWorkerLoadTemplate.ProgressChanged += BgWorkerLoadTemplate_ProgressChanged;
            bgWorkerLoadTemplate.WorkerReportsProgress = true;

            bgWorkerReport = new BackgroundWorker();
            bgWorkerReport.DoWork += BgWorkerReport_DoWork;
            bgWorkerReport.RunWorkerCompleted += BgWorkerReport_RunWorkerCompleted;
            bgWorkerReport.ProgressChanged += BgWorkerReport_ProgressChanged;
            bgWorkerReport.WorkerReportsProgress = true;
        }

        private void BgWorkerReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 100)
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學期成績完成");
            else
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學期成績產生中...", e.ProgressPercentage);
        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 產生
            try
            {
                object[] objArray = (object[])e.Result;

                btnSaveConfig.Enabled = true;
                btnPrint.Enabled = true;

                //if (_ErrorList.Count > 0)
                //{
                //    StringBuilder sb = new StringBuilder();
                //    //sb.AppendLine("樣板內科目合併欄位不足，請新增：");
                //    //sb.AppendLine(string.Join(",", _ErrorList.ToArray()));
                //    sb.AppendLine("1.樣板內科目合併欄位不足，請檢查樣板。");
                //    sb.AppendLine("2.如果使用只有領域樣板，請忽略此訊息。");
                //    //if (_ErrorDomainNameList.Count > 0)
                //        sb.AppendLine(string.Join(",", _ErrorDomainNameList.ToArray()));

                //    FISCA.Presentation.Controls.MsgBox.Show(sb.ToString(), "樣板內科目合併欄位不足", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                //}

                #region 儲存檔案

                string reportName = "" + Global._SelSchoolYear + "學年度第" + Global._SelSemester + "學期 班級學期成績單";

                string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                path = Path.Combine(path, reportName + ".doc");

                if (File.Exists(path))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                    }
                }

                Document document = new Document();
                try
                {
                    document = (Document)objArray[0];
                    document.Save(path, SaveFormat.Doc);
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = reportName + ".doc";
                    sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);

                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                #endregion

                FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績報表產生完成");
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
            }
        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void BgWorkerLoadTemplate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 100)
                FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            else
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學期成績設定資料讀取中...", e.ProgressPercentage);

        }

        private void BgWorkerLoadTemplate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_Configure == null)
                _Configure = new Configure();

            cboConfigure.Items.Clear();
            foreach (var item in _ConfigureList)
            {
                cboConfigure.Items.Add(item);
            }
            cboConfigure.Items.Add(new Configure() { Name = "新增" });

            if (_ConfigureList.Count > 0)
            {
                cboConfigure.SelectedIndex = 0;
            }
            else
            {
                cboConfigure.SelectedIndex = -1;
            }

            string userSelectConfigName = "";
            // 檢查畫面上是否有使用者選的
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    userSelectConfigName = conf.Name;
                    break;
                }


            if (!string.IsNullOrEmpty(_Configure.SelSetConfigName))
                cboConfigure.Text = userSelectConfigName;

        //    ReloadCanSelectDomainSubject(cboSchoolYear.Text, cboSemester.Text, _ClassIDList, _Configure.ExamRecordID);

            btnSaveConfig.Enabled = btnPrint.Enabled = true;
        }

        /// <summary>
        /// 重新讀取可選科目領域
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="_ClassIDList"></param>
        /// <param name="ExamID"></param>
        private void ReloadCanSelectDomainSubject(string SchoolYear, string Semester, List<string> _ClassIDList)
        {
            
            LoadDomainSubject();
            // 解析科目
            foreach (ListViewItem lvi in lvSubject.Items)
            {
                if (_Configure.PrintSubjectList != null && _Configure.PrintSubjectList.Contains(lvi.Text))
                {
                    lvi.Checked = true;
                }
            }

            // 解析領域
            // 解析科目
            foreach (ListViewItem lvi in lvDomain.Items)
            {
                if (_Configure.PrintDomainList != null && _Configure.PrintDomainList.Contains(lvi.Text))
                {
                    lvi.Checked = true;
                }
            }
        }

        private void LoadDomainSubject()
        {
            // 清空畫面上領域科目
            lvDomain.Items.Clear();
            lvSubject.Items.Clear();

            // 待修
            //if (CanSelectExamDomainSubjectDict.ContainsKey(SelExamID))
            //{
            //    foreach (string dName in CanSelectExamDomainSubjectDict[SelExamID].Keys)
            //    {
            //        ListViewItem lvi = new ListViewItem();
            //        lvi.Tag = dName;
            //        lvi.Name = dName;
            //        lvi.Text = dName;
            //        lvDomain.Items.Add(lvi);

            //        foreach (string sName in CanSelectExamDomainSubjectDict[SelExamID][dName])
            //        {
            //            ListViewItem lvis = new ListViewItem();
            //            lvis.Name = dName + "_" + sName;
            //            lvis.Text = sName;
            //            lvis.Tag = dName; // 放領域
            //            lvSubject.Items.Add(lvis);
            //        }
            //    }
            //}
        }

        private void BgWorkerLoadTemplate_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerLoadTemplate.ReportProgress(1);


            // 檢查預設樣板是否存在
            _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);

            bgWorkerLoadTemplate.ReportProgress(30);

          


            bgWorkerLoadTemplate.ReportProgress(50);

            bgWorkerLoadTemplate.ReportProgress(80);
            // 取的設定資料
            _ConfigureList = _AccessHelper.Select<Configure>();
            bgWorkerLoadTemplate.ReportProgress(100);
        }

        public void SetClassIDList(List<string> classIDList)
        {
            _ClassIDList = classIDList;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            //儲存設定檔
            SaveConfig();

            Global.SelOnlyDomainList.Clear();
            Global.SelOnlySubjectList.Clear();

            // 取得畫面上勾選領域科目
            SelDomainSubjectDict.Clear();
            // 領域
            foreach (ListViewItem lvi in lvDomain.CheckedItems)
            {
                string key = lvi.Name;
                if (!SelDomainSubjectDict.ContainsKey(key))
                {
                    SelDomainSubjectDict.Add(key, new List<string>());
                    if (!Global.SelOnlyDomainList.Contains(key))
                        Global.SelOnlyDomainList.Add(key);
                }
            }

            foreach (ListViewItem lvi in lvSubject.CheckedItems)
            {
                string key = lvi.Name;

                // 檢查自己領域是否有加入
                string keyD = lvi.Tag.ToString();

                // 只有勾科目沒有領域，只算科目
                if (!Global.SelOnlyDomainList.Contains(keyD))
                {
                    Global.SelOnlySubjectList.Add(lvi.Text);
                }

                if (!SelDomainSubjectDict.ContainsKey(keyD))
                    SelDomainSubjectDict.Add(keyD, new List<string>());

                if (!SelDomainSubjectDict[keyD].Contains(lvi.Text))
                    SelDomainSubjectDict[keyD].Add(lvi.Text);
            }

            // 列印報表
            bgWorkerReport.RunWorkerAsync();

        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            SaveConfig();
        }

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "新竹學期成績單樣板(" + _Configure.Name + ")";

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                _Configure.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
                stream.Write(Properties.Resources.新竹班級學期成績單樣板, 0, Properties.Resources.新竹班級學期成績單樣板.Length);
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.新竹班級學期成績單樣板, 0, Properties.Resources.新竹班級學期成績單樣板.Length);
                        stream.Flush();
                        stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkViewTemplate.Enabled = true;
            #endregion
        }

        /// <summary>
        /// 儲存設定檔
        /// </summary>
        public void SaveConfig()
        {
            //儲存設定
            if (_Configure == null) return;
            _Configure.SchoolYear = cboSchoolYear.Text;
            _Configure.Semester = cboSemester.Text;
            _Configure.SelSetConfigName = cboConfigure.Text;
            _Configure.ReScoreMark = txtReScoreMark.Text;
            _Configure.NeeedReScoreMark = txtNeeedReScoreMark.Text;
            _Configure.FailScoreMark = txtFailScoreMark.Text;


            // 科目
            foreach (ListViewItem item in lvSubject.Items)
            {
                if (item.Checked)
                {
                    if (!_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Add(item.Text);
                }
                else
                {
                    if (_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Remove(item.Text);
                }
            }

            // 領域
            foreach (ListViewItem item in lvDomain.Items)
            {
                if (item.Checked)
                {
                    if (!_Configure.PrintDomainList.Contains(item.Text))
                        _Configure.PrintDomainList.Add(item.Text);
                }
                else
                {
                    if (_Configure.PrintDomainList.Contains(item.Text))
                        _Configure.PrintDomainList.Remove(item.Text);
                }
            }

            Global._SelSchoolYear = cboSchoolYear.Text;
            Global._SelSemester = cboSemester.Text;

            _Configure.Encode();
            _Configure.Save();

            #region 樣板設定檔記錄用

            // 記錄使用這選的專案            
            List<DAO.UDT_ScoreConfig> uList = new List<DAO.UDT_ScoreConfig>();
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    conf.Name = cboConfigure.Text;
                    uList.Add(conf);
                    break;
                }

            if (uList.Count > 0)
            {
                DAO.UDTTransfer.UpdateConfigData(uList);
            }
            else
            {
                // 新增
                List<DAO.UDT_ScoreConfig> iList = new List<DAO.UDT_ScoreConfig>();
                DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                conf.Name = cboConfigure.Text;
                conf.ProjectName = Global._ProjectName;
                conf.Type = Global._UserConfTypeName;
                conf.UDTTableName = Global._UDTTableName;
                iList.Add(conf);
                DAO.UDTTransfer.InsertConfigData(iList);
            }
            #endregion
        }

        private void lnkChangeTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            lnkChangeTemplate.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    _Configure.Template = new Aspose.Words.Document(dialog.FileName);
                    List<string> fields = new List<string>(_Configure.Template.MailMerge.GetFieldNames());
                    _Configure.SubjectLimit = 0;
                    while (fields.Contains("科目名稱" + (_Configure.SubjectLimit + 1)))
                    {
                        _Configure.SubjectLimit++;
                    }

                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
            lnkChangeTemplate.Enabled = true;
        }

        private void lnkViewMapColumns_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            
            // 產生合併欄位總表
            lnkViewMapColumns.Enabled = false;

            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            cboSemester.Text = K12.Data.School.DefaultSemester;
          
            this.MaximumSize = this.MinimumSize = this.Size;
            bgWorkerLoadTemplate.RunWorkerAsync();
        }
    }
}
