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

namespace HsinChuExamScoreClassFixedRank.Form
{
    public partial class PrintForm : BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        List<string> _ClassIDList = new List<string>();
        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();
        public Configure _Configure { get; private set; }

        List<ExamRecord> _exams = new List<ExamRecord>();

        BackgroundWorker bgWorkerLoadTemplate;
        BackgroundWorker bgWorkerReport;
        string SelSchoolYear = "";
        string SelSemester = "";
        string SelExamID = "";
        // 紀錄樣板設定
        List<DAO.UDT_ScoreConfig> _UDTConfigList;
        Dictionary<string, Dictionary<string, List<string>>> CanSelectExamDomainSubjectDict = new Dictionary<string, Dictionary<string, List<string>>>();

        public PrintForm()
        {
            InitializeComponent();
            bgWorkerLoadTemplate = new BackgroundWorker();
            bgWorkerLoadTemplate.DoWork += BgWorkerLoadTemplate_DoWork;
            bgWorkerLoadTemplate.ProgressChanged += BgWorkerLoadTemplate_ProgressChanged;
            bgWorkerLoadTemplate.WorkerReportsProgress = true;
            bgWorkerLoadTemplate.RunWorkerCompleted += BgWorkerLoadTemplate_RunWorkerCompleted;
            bgWorkerReport = new BackgroundWorker();
            bgWorkerReport.DoWork += BgWorkerReport_DoWork;
            bgWorkerReport.ProgressChanged += BgWorkerReport_ProgressChanged;
            bgWorkerReport.WorkerReportsProgress = true;
            bgWorkerReport.RunWorkerCompleted += BgWorkerReport_RunWorkerCompleted;



        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void BgWorkerReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {
            throw new NotImplementedException();
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
            cboExam.Items.Clear();
            foreach (ExamRecord exName in _exams)
                cboExam.Items.Add(exName.Name);

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



        }

        private void BgWorkerLoadTemplate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 100)
                FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            else
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級評量成績設定資料讀取中...", e.ProgressPercentage);
        }

        private void BgWorkerLoadTemplate_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerLoadTemplate.ReportProgress(1);

            //試別清單
            _exams.Clear();
            _exams = K12.Data.Exam.SelectAll();


            // 檢查預設樣板是否存在
            _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);

            // 沒有設定檔，建立預設設定檔
            if (_UDTConfigList.Count < 2)
            {
                bgWorkerLoadTemplate.ReportProgress(30);

                foreach (string name in Global.DefaultConfigNameList())
                {
                    Configure cn = new Configure();
                    cn.Name = name;
                    cn.SchoolYear = K12.Data.School.DefaultSchoolYear;
                    cn.Semester = K12.Data.School.DefaultSemester;
                    DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                    conf.Name = name;
                    conf.UDTTableName = Global._UDTTableName;
                    conf.ProjectName = Global._ProjectName;
                    conf.Type = Global._DefaultConfTypeName;
                    _UDTConfigList.Add(conf);

                    // 設預設樣板
                    switch (name)
                    {
                        //case "領域成績單":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_領域成績單));
                        //    break;

                        //case "科目成績單":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目成績單));
                        //    break;

             
                    }

                    if (cn.Template == null)
                        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹班級評量成績單樣板));
                    cn.Encode();
                    cn.Save();
                }
                if (_UDTConfigList.Count > 0)
                    DAO.UDTTransfer.InsertConfigData(_UDTConfigList);
            }


            bgWorkerLoadTemplate.ReportProgress(50);

            CanSelectExamDomainSubjectDict = DAO.DataAccess.GetExamDomainSubjectDictByClass(K12.Data.School.DefaultSchoolYear, K12.Data.School.DefaultSemester, _ClassIDList);

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

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            //儲存設定
        }



        private void lnkDelConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;

            // 檢查是否是預設設定檔名稱，如果是無法刪除
            if (Global.DefaultConfigNameList().Contains(_Configure.Name))
            {
                FISCA.Presentation.Controls.MsgBox.Show("系統預設設定檔案無法刪除");
                return;
            }

            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _ConfigureList.Remove(_Configure);
                if (_Configure.UID != "")
                {
                    _Configure.Deleted = true;
                    _Configure.Save();
                }
                var conf = _Configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }
        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            // 2019/10/22 與佳樺討論，學年度學期使用系統預設，先鎖定不讓使用者更動。
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            cboSemester.Text = K12.Data.School.DefaultSemester;
            cboSchoolYear.Enabled = false;
            cboSemester.Enabled = false;
            this.MaximumSize = this.MinimumSize = this.Size;
            bgWorkerLoadTemplate.RunWorkerAsync();
        }

        private void lnkCopyConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = _Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;
                conf.ExamRecord = _Configure.ExamRecord;
                conf.PrintSubjectList.AddRange(_Configure.PrintSubjectList);
                conf.SchoolYear = _Configure.SchoolYear;
                conf.Semester = _Configure.Semester;
                conf.SubjectLimit = _Configure.SubjectLimit;
                conf.Template = _Configure.Template;
                conf.BeginDate = _Configure.BeginDate;
                conf.EndDate = _Configure.EndDate;
                conf.ScoreEditDate = _Configure.ScoreEditDate;
              
                conf.Encode();
                conf.Save();
                _ConfigureList.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "新竹評量成績單樣板(" + _Configure.Name + ")";

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
                stream.Write(Properties.Resources.新竹班級評量成績單樣板, 0, Properties.Resources.新竹班級評量成績單樣板.Length);
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
                        stream.Write(Properties.Resources.新竹班級評量成績單樣板, 0, Properties.Resources.新竹班級評量成績單樣板.Length);
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
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // 列印報表
        }

        private void cboExam_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelExamID = "";
            foreach(ExamRecord er in _exams)
            {
                if (er.Name == cboExam.Text)
                {
                    SelExamID = er.ID;
                }
            }
            
            LoadDomainSubject();
        }

        private void LoadDomainSubject ()
        {
            // 清空畫面上領域科目
            lvDomain.Items.Clear();
            lvSubject.Items.Clear();
            List<string> dNameList = new List<string>();
            List<string> sNameList = new List<string>();
            if (CanSelectExamDomainSubjectDict.ContainsKey(SelExamID))
            {
                foreach(string dName in CanSelectExamDomainSubjectDict[SelExamID].Keys)
                {
                    dNameList.Add(dName);

                    foreach(string sName in CanSelectExamDomainSubjectDict[SelExamID][dName])
                    {
                        if (!sNameList.Contains(sName))
                            sNameList.Add(sName);
                    }
                }                
            }

            foreach (string dName in dNameList)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Name = dName;
                lvi.Text = dName;
                lvDomain.Items.Add(lvi);
            }

            foreach (string sName in sNameList)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Name = sName;
                lvi.Text = sName;
                lvSubject.Items.Add(lvi);
            }

        }


        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = btnPrint.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    _Configure = new Configure();
                    _Configure.Name = dialog.ConfigName;
                    _Configure.Template = dialog.Template;
                    _Configure.SubjectLimit = dialog.SubjectLimit;
                    _Configure.SchoolYear = K12.Data.School.DefaultSchoolYear;
                    _Configure.Semester = K12.Data.School.DefaultSemester;
                    if (_Configure.PrintSubjectList == null)
                        _Configure.PrintSubjectList = new List<string>();

                    if (cboExam.Items.Count > 0)
                    {
                        string exName = cboExam.Items[0].ToString();
                        foreach (ExamRecord rec in _exams)
                        {
                            if (exName == rec.Name)
                            {
                                _Configure.ExamRecord = rec;
                                break;
                            }
                        }

                    }
                    _ConfigureList.Add(_Configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, _Configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;
                    _Configure.Encode();
                    _Configure.Save();
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnSaveConfig.Enabled = btnPrint.Enabled = true;
                    _Configure = _ConfigureList[cboConfigure.SelectedIndex];
                    if (_Configure.Template == null)
                        _Configure.Decode();
                    if (!cboSchoolYear.Items.Contains(_Configure.SchoolYear))
                        cboSchoolYear.Items.Add(_Configure.SchoolYear);
                    cboSchoolYear.Text = _Configure.SchoolYear;
                    cboSemester.Text = _Configure.Semester;
                    if (_Configure.ExamRecord != null)
                    {
                        int idx = 0;
                        foreach (string sitm in cboExam.Items)
                        {
                            if (sitm == _Configure.ExamRecord.Name)
                            {
                                cboExam.SelectedIndex = idx;
                                break;
                            }
                            idx++;
                        }

                    }

                    // 解析科目
                    foreach (ListViewItem lvi in lvSubject.Items)
                    {
                        if (_Configure.PrintSubjectList.Contains(lvi.Text))
                        {
                            lvi.Checked = true;
                        }
                    }                                       
                }
                else
                {
                    _Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;
                    cboExam.SelectedIndex = -1;
               
                }
            }
        }
    }
}
