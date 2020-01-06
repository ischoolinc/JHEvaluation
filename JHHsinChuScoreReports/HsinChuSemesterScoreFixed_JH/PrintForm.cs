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
using Campus.ePaperCloud;
using HsinChuSemesterScoreFixed_JH.DAO;

namespace HsinChuSemesterScoreFixed_JH
{
    public partial class PrintForm : BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();


        List<string> _StudentIDList;

        DataTable dt = new DataTable();
        private List<string> typeList = new List<string>();

        BackgroundWorker _bgWorkReport;
        DocumentBuilder _builder;
        BackgroundWorker _bgWorkerLoadData;

        string NeeedReScoreMark = "";
        string ReScoreMark = "";
        string FailScoreMark = "";

        // 錯誤訊息
        List<string> _ErrorList = new List<string>();

        // 領域錯誤訊息
        List<string> _ErrorDomainNameList = new List<string>();

        // 樣板內有科目名稱
        List<string> _TemplateSubjectNameList = new List<string>();

        // 存檔路徑
        string pathW = "";

        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();


        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        private int _SelSchoolYear;
        private int _SelSemester;

        ScoreMappingConfig _ScoreMappingConfig = new ScoreMappingConfig();

        // 紀錄樣板設定
        List<DAO.UDT_ScoreConfig> _UDTConfigList;

        public PrintForm(List<string> StudIDList)
        {
            InitializeComponent();
            _bgWorkerLoadData = new BackgroundWorker();
            _bgWorkReport = new BackgroundWorker();
            _bgWorkerLoadData.DoWork += _bgWorkerLoadData_DoWork;
            _bgWorkerLoadData.RunWorkerCompleted += _bgWorkerLoadData_RunWorkerCompleted;
            _bgWorkerLoadData.WorkerReportsProgress = true;
            _bgWorkerLoadData.ProgressChanged += _bgWorkerLoadData_ProgressChanged;
            _bgWorkReport.DoWork += _bgWorkReport_DoWork;
            _bgWorkReport.RunWorkerCompleted += _bgWorkReport_RunWorkerCompleted;
            _bgWorkReport.WorkerReportsProgress = true;
            _bgWorkReport.ProgressChanged += _bgWorkReport_ProgressChanged;

            _StudentIDList = StudIDList;

        }

        private void _bgWorkReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績報表產生中...", e.ProgressPercentage);
        }

        private void _bgWorkReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
           
        }

        private void _bgWorkReport_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorkReport.ReportProgress(1);
            dt.Clear();

            // 建立合併欄位
            dt.Columns.Add("列印日期");
            dt.Columns.Add("學校名稱");
            dt.Columns.Add("學年度");
            dt.Columns.Add("學期");
            dt.Columns.Add("系統編號");
            dt.Columns.Add("姓名");
            dt.Columns.Add("英文姓名");
            dt.Columns.Add("班級");
            dt.Columns.Add("座號");
            dt.Columns.Add("學號");
            dt.Columns.Add("大功");
            dt.Columns.Add("小功");
            dt.Columns.Add("嘉獎");
            dt.Columns.Add("大過");
            dt.Columns.Add("小過");
            dt.Columns.Add("警告");
            dt.Columns.Add("上課天數");
            dt.Columns.Add("學習領域成績");
            dt.Columns.Add("學習領域原始成績");
            dt.Columns.Add("課程學習成績");
            dt.Columns.Add("課程學習原始成績");
            dt.Columns.Add("班導師");
            dt.Columns.Add("教務主任");
            dt.Columns.Add("校長");
            dt.Columns.Add("服務學習時數");
            dt.Columns.Add("文字描述");




        }

        private void _bgWorkerLoadData_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績報表資料讀取中...", e.ProgressPercentage);
        }

        private void _bgWorkerLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_Configure == null)
                _Configure = new Configure();

            _DefalutSchoolYear = K12.Data.School.DefaultSchoolYear;
            _DefaultSemester = K12.Data.School.DefaultSemester;

            cboConfigure.Items.Clear();
            foreach (var item in _ConfigureList)
            {
                cboConfigure.Items.Add(item);
            }
            cboConfigure.Items.Add(new Configure() { Name = "新增" });
            int i;

            if (int.TryParse(_DefalutSchoolYear, out i))
            {
                for (int j = 5; j > 0; j--)
                {
                    cboSchoolYear.Items.Add("" + (i - j));
                }

                for (int j = 0; j < 3; j++)
                {
                    cboSchoolYear.Items.Add("" + (i + j));
                }

            }

            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");


            string userSelectConfigName = "";
            // 檢查畫面上是否有使用者選的
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    userSelectConfigName = conf.Name;
                    break;
                }

           if (!string.IsNullOrEmpty(userSelectConfigName))
                cboConfigure.Text = userSelectConfigName;
            
            EnableSelectItem(true);
        }

        private void _bgWorkerLoadData_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorkerLoadData.ReportProgress(1);
            // 檢查預設樣板是否存在
            _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);

            // 沒有設定檔，建立預設設定檔
            if (_UDTConfigList.Count < 2)
            {
                _bgWorkerLoadData.ReportProgress(20);
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

                        //case "科目及領域成績單_領域組距":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_領域組距));
                        //    break;
                        //case "科目及領域成績單_科目組距":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_科目組距));
                        //    break;
                    }

                    if (cn.Template == null)
                        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹學期成績單樣板_固定排名_科目_領域));
                    cn.Encode();
                    cn.Save();
                }
                if (_UDTConfigList.Count > 0)
                    DAO.UDTTransfer.InsertConfigData(_UDTConfigList);
            }
            _bgWorkerLoadData.ReportProgress(70);
            // 取的設定資料
            _ConfigureList = _AccessHelper.Select<Configure>();

            _bgWorkerLoadData.ReportProgress(100);
        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            EnableSelectItem(false);
            _SelSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _SelSemester = int.Parse(K12.Data.School.DefaultSemester);
            _ScoreMappingConfig.LoadData();
            _bgWorkerLoadData.RunWorkerAsync();

        }


        // 啟用項目
        private void EnableSelectItem(bool enable)
        {
            cboConfigure.Enabled = enable;
            cboSchoolYear.Enabled = enable;
            cboSemester.Enabled = enable;
            lnkViewTemplate.Enabled = enable;
            lnkChangeTemplate.Enabled = enable;
            lnkViewMapColumns.Enabled = enable;
            btnSaveConfig.Enabled = enable;
            btnPrint.Enabled = enable;
        }



        private void lnkCopyConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = _Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;
                conf.PrintSubjectList.AddRange(_Configure.PrintSubjectList);
                conf.SchoolYear = _Configure.SchoolYear;
                conf.Semester = _Configure.Semester;
                conf.SubjectLimit = _Configure.SubjectLimit;
                conf.Template = _Configure.Template;
                conf.NeeedReScoreMark = _Configure.NeeedReScoreMark;
                conf.ReScoreMark = _Configure.ReScoreMark;
                conf.FailScoreMark = _Configure.FailScoreMark;

                if (conf.PrintAttendanceList == null)
                    conf.PrintAttendanceList = new List<string>();
                conf.PrintAttendanceList.AddRange(_Configure.PrintAttendanceList);
                conf.Encode();
                conf.Save();
                _ConfigureList.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }

        public Configure _Configure { get; private set; }

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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void btnPrint_Click(object sender, EventArgs e)
        {


            int sc, ss;
            if (int.TryParse(cboSchoolYear.Text, out sc))
            {
                _SelSchoolYear = sc;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學年度必填!");
                return;
            }

            if (int.TryParse(cboSemester.Text, out ss))
            {
                _SelSemester = ss;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學期必填!");
                return;
            }


            SaveTemplate(null, null);

            btnSaveConfig.Enabled = false;
           

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
            // 執行報表
            _bgWorkReport.RunWorkerAsync();
        }

        // 儲存樣板
        private void SaveTemplate(object sender, EventArgs e)
        {
            if (_Configure == null) return;
            _Configure.SchoolYear = cboSchoolYear.Text;
            _Configure.Semester = cboSemester.Text;
            _Configure.SelSetConfigName = cboConfigure.Text;
            _Configure.NeeedReScoreMark = txtNeeedReScoreMark.Text;
            _Configure.ReScoreMark = txtReScoreMark.Text;
            _Configure.FailScoreMark = txtFailScoreMark.Text;

            if (_Configure.PrintAttendanceList == null)
                _Configure.PrintAttendanceList = new List<string>();

            _Configure.PrintAttendanceList.Clear();

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
                    _Configure.SchoolYear = _DefalutSchoolYear;
                    _Configure.Semester = _DefaultSemester;
                    if (_Configure.PrintAttendanceList == null)
                        _Configure.PrintAttendanceList = new List<string>();
                    if (_Configure.PrintSubjectList == null)
                        _Configure.PrintSubjectList = new List<string>();


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
                    txtFailScoreMark.Text = _Configure.FailScoreMark;
                    txtReScoreMark.Text = _Configure.ReScoreMark;
                    txtNeeedReScoreMark.Text = _Configure.NeeedReScoreMark;
                }
                else
                {
                    _Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;

                }
            }
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
                        stream.Write(Properties.Resources.新竹學期成績單樣板_固定排名_科目_領域, 0, Properties.Resources.新竹學期成績單樣板_固定排名_科目_領域.Length);
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

        private void lnkChangeTemplate_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
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

        private void lnkViewMapColumns_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            int sc, ss;
            if (int.TryParse(cboSchoolYear.Text, out sc))
            {
                _SelSchoolYear = sc;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學年度必填!");
                return;
            }

            if (int.TryParse(cboSemester.Text, out ss))
            {
                _SelSemester = ss;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學期必填!");
                return;
            }

            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
        }


    }
}
