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
using HsinChuSemesterClassFixedRank.DAO;

namespace HsinChuSemesterClassFixedRank
{
    public partial class PrintForm : BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        List<string> _ClassIDList = new List<string>();
        // 一般狀態學生與班級ID
        Dictionary<string, string> StudentClassIDDict = new Dictionary<string, string>();

        // 班級學生ID
        Dictionary<string, List<StudentRecord>> ClassStudentIDDict = new Dictionary<string, List<StudentRecord>>();

        // 學生學期成績
        Dictionary<string, JHSemesterScoreRecord> StudentSemesterScoreRecordDict = new Dictionary<string, JHSemesterScoreRecord>();

        // 班級領域名稱
        Dictionary<string, List<string>> ClassDomainNameDict = new Dictionary<string, List<string>>();
        // 班級科目名稱
        Dictionary<string, List<string>> ClassSubjectNameDict = new Dictionary<string, List<string>>();

        Dictionary<string, int> SelDomainIdxDict = new Dictionary<string, int>();
        Dictionary<string, int> SelSubjectIdxDict = new Dictionary<string, int>();

        // 畫面上可選領域與科目
        Dictionary<string, List<string>> CanSelectDomainSubjectDict = new Dictionary<string, List<string>>();

        // 錯誤訊息
        List<string> _ErrorList = new List<string>();


        // 領域錯誤訊息
        List<string> _ErrorDomainNameList = new List<string>();

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

                if (_ErrorList.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    //sb.AppendLine("樣板內科目合併欄位不足，請新增：");
                    //sb.AppendLine(string.Join(",", _ErrorList.ToArray()));
                    sb.AppendLine("1.樣板內科目合併欄位不足，請檢查樣板。");
                    sb.AppendLine("2.如果使用只有領域樣板，請忽略此訊息。");
                    //if (_ErrorDomainNameList.Count > 0)
                    sb.AppendLine(string.Join(",", _ErrorDomainNameList.ToArray()));

                    FISCA.Presentation.Controls.MsgBox.Show(sb.ToString(), "樣板內科目合併欄位不足", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                }

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
            DataTable dtTable = new DataTable();

            // 取得相關資料
            bgWorkerReport.ReportProgress(1);

            // 每次合併後放入，最後再合成一張
            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.新竹班級學期成績單樣板));

            _ErrorList.Clear();
            _ErrorDomainNameList.Clear();

            // 校名
            string SchoolName = K12.Data.School.ChineseName;
            // 校長
            string ChancellorChineseName = JHSchool.Data.JHSchoolInfo.ChancellorChineseName;
            // 教務主任
            string EduDirectorName = JHSchool.Data.JHSchoolInfo.EduDirectorName;

            // 取得班導師
            Dictionary<string, string> ClassTeacherNameDict = DataAccess.GetClassTeacherNameDictByClassID(_ClassIDList);

            // 取得學生學期成績排名、五標、分數區間
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> SemsScoreRankMatrixDataValueDict = DataAccess.GetSemsScoreRankMatrixDataValue(_Configure.SchoolYear, _Configure.Semester, StudentClassIDDict.Keys.ToList());

            StreamWriter sw = new StreamWriter(Application.StartupPath + "\\學生五標排名組距.txt");
            List<string> tmpList = new List<string>();
            foreach (string key in SemsScoreRankMatrixDataValueDict.Keys)
            {
                foreach (string key1 in SemsScoreRankMatrixDataValueDict[key].Keys)
                {
                    if (!tmpList.Contains(key1))
                        tmpList.Add(key1);
                }
            }
            foreach (string key in tmpList)
                sw.WriteLine(key);
            sw.Close();

            #region 產生合併欄位 Columns


            List<string> r2List = new List<string>();
            r2List.Add("rank");
            r2List.Add("matrix_count");
            r2List.Add("pr");
            r2List.Add("percentile");
            r2List.Add("avg_top_25");
            r2List.Add("avg_top_50");
            r2List.Add("avg");
            r2List.Add("avg_bottom_50");
            r2List.Add("avg_bottom_25");
            r2List.Add("level_gte100");
            r2List.Add("level_90");
            r2List.Add("level_80");
            r2List.Add("level_70");
            r2List.Add("level_60");
            r2List.Add("level_50");
            r2List.Add("level_40");
            r2List.Add("level_30");
            r2List.Add("level_20");
            r2List.Add("level_10");
            r2List.Add("level_lt10");

            List<string> r1List = new List<string>();
            r1List.Add("課程學習總成績");
            r1List.Add("學習領域總成績");
            r1List.Add("課程學習總成績(原始)");
            r1List.Add("學習領域總成績(原始)");

            List<string> rkList = new List<string>();
            rkList.Add("班排名");
            rkList.Add("年排名");
            rkList.Add("類別1排名");
            rkList.Add("類別2排名");

            dtTable.Columns.Add("學校名稱");
            dtTable.Columns.Add("學年度");
            dtTable.Columns.Add("學期");
            dtTable.Columns.Add("班級");
            dtTable.Columns.Add("班導師");


            int maxStudent = 30;

            foreach (string key in ClassStudentIDDict.Keys)
            {
                if (ClassStudentIDDict[key].Count > maxStudent)
                    maxStudent = ClassStudentIDDict[key].Count;
            }


            // 學生合併欄位
            for (int studIdx = 1; studIdx <= maxStudent; studIdx++)
            {
                dtTable.Columns.Add("姓名" + studIdx);
                dtTable.Columns.Add("座號" + studIdx);
                // 總計成績合併欄位
                // 學生1_課程學習總成績_班排名_rank
                foreach (string r1 in r1List)
                {
                    dtTable.Columns.Add("學生" + studIdx + "_" + r1 + "_成績");
                    foreach (string rk in rkList)
                    {
                        // 五標、PR、百分比、組距
                        foreach (string r2 in r2List)
                        {
                            dtTable.Columns.Add("學生" + studIdx + "_" + r1 + "_" + rk + "_" + r2);
                        }
                    }

                }

                // 領域1,2 合併欄位
                for (int dIdx = 1; dIdx <= SelDomainIdxDict.Keys.Count; dIdx++)
                {
                    // 學生1_領域1_成績
                    dtTable.Columns.Add("學生" + studIdx + "_領域" + dIdx + "_名稱");
                    dtTable.Columns.Add("學生" + studIdx + "_領域" + dIdx + "_學分");
                    dtTable.Columns.Add("學生" + studIdx + "_領域" + dIdx + "_成績");
                    dtTable.Columns.Add("學生" + studIdx + "_領域(原始)" + dIdx + "_成績");

                    foreach (string rk in rkList)
                    {
                        // 五標、PR、百分比、組距
                        foreach (string r2 in r2List)
                        {
                            dtTable.Columns.Add("學生" + studIdx + "_領域" + dIdx + "_" + rk + "_" + r2);
                            dtTable.Columns.Add("學生" + studIdx + "_領域(原始)" + dIdx + "_" + rk + "_" + r2);
                        }
                    }

                }

                // 科目1,2 合併欄位
                for (int sIdx = 1; sIdx <= SelSubjectIdxDict.Keys.Count; sIdx++)
                {
                    // 學生1_科目1_成績
                    dtTable.Columns.Add("學生" + studIdx + "_科目" + sIdx + "_名稱");
                    dtTable.Columns.Add("學生" + studIdx + "_科目" + sIdx + "_學分");
                    dtTable.Columns.Add("學生" + studIdx + "_科目" + sIdx + "_成績");
                    dtTable.Columns.Add("學生" + studIdx + "_科目(原始)" + sIdx + "_成績");

                    foreach (string rk in rkList)
                    {
                        // 五標、PR、百分比、組距
                        foreach (string r2 in r2List)
                        {
                            dtTable.Columns.Add("學生" + studIdx + "_科目" + sIdx + "_" + rk + "_" + r2);
                            dtTable.Columns.Add("學生" + studIdx + "_科目(原始)" + sIdx + "_" + rk + "_" + r2);
                        }
                    }

                }
            }


            // 班級合併欄位
            foreach (string r1 in r1List)
            {
                foreach (string rk in rkList)
                {
                    // 五標、PR、百分比、組距
                    foreach (string r2 in r2List)
                    {
                        dtTable.Columns.Add("班級_" + r1 + "_" + rk + "_" + r2);
                    }
                }

            }

            // 領域1,2 合併欄位
            for (int dIdx = 1; dIdx <= SelDomainIdxDict.Keys.Count; dIdx++)
            {
                // 班級1_領域1_成績
                dtTable.Columns.Add("班級_領域" + dIdx + "_名稱");

                foreach (string rk in rkList)
                {
                    // 五標、PR、百分比、組距
                    foreach (string r2 in r2List)
                    {
                        dtTable.Columns.Add("班級_領域" + dIdx + "_" + rk + "_" + r2);
                        dtTable.Columns.Add("班級_領域(原始)" + dIdx + "_" + rk + "_" + r2);
                    }
                }

            }

            // 科目1,2 合併欄位
            for (int sIdx = 1; sIdx <= SelSubjectIdxDict.Keys.Count; sIdx++)
            {
                // 班級1_科目1_成績
                dtTable.Columns.Add("班級_科目" + sIdx + "_名稱");

                foreach (string rk in rkList)
                {
                    // 五標、PR、百分比、組距
                    foreach (string r2 in r2List)
                    {
                        dtTable.Columns.Add("班級_科目" + sIdx + "_" + rk + "_" + r2);
                        dtTable.Columns.Add("班級_科目(原始)" + sIdx + "_" + rk + "_" + r2);
                    }
                }

            }



            StreamWriter sw1 = new StreamWriter(Application.StartupPath + "\\合併欄位.txt");
            StringBuilder sb1 = new StringBuilder();
            foreach (DataColumn dc in dtTable.Columns)
                sb1.AppendLine(dc.Caption);

            sw1.Write(sb1.ToString());
            sw1.Close();

            #endregion


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

            int sc = 0, ss = 0;
            int.TryParse(_Configure.SchoolYear, out sc);
            int.TryParse(_Configure.Semester, out ss);

            if (sc > 0 && ss > 0 && StudentClassIDDict.Keys.Count > 0)
            {
                LoadStudentSemesterScore(sc, ss, StudentClassIDDict.Keys.ToList());
            }

            ReloadCanSelectDomainSubject();

            btnSaveConfig.Enabled = btnPrint.Enabled = true;
        }


        private void ReloadCanSelectDomainSubject()
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

            List<string> tmpDomainList = CanSelectDomainSubjectDict.Keys.ToList();
            tmpDomainList.Remove("彈性課程");
            tmpDomainList.Sort(new StringComparer("語文", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"));
            tmpDomainList.Add("彈性課程");

            foreach (string dName in tmpDomainList)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Tag = dName;
                lvi.Name = dName;
                lvi.Text = dName;
                lvDomain.Items.Add(lvi);

                if (CanSelectDomainSubjectDict.ContainsKey(dName))
                {
                    foreach (string sName in CanSelectDomainSubjectDict[dName])
                    {
                        ListViewItem lvis = new ListViewItem();
                        lvis.Name = dName + "_" + sName;
                        lvis.Text = sName;
                        lvis.Tag = dName; // 放領域
                        lvSubject.Items.Add(lvis);
                    }
                }
            }

        }

        private void BgWorkerLoadTemplate_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerLoadTemplate.ReportProgress(1);

            // 檢查預設樣板是否存在
            _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);

            bgWorkerLoadTemplate.ReportProgress(30);

            //  取得班級一般狀態學生ID
            StudentClassIDDict.Clear();
            ClassStudentIDDict.Clear();
            List<JHStudentRecord> studRecList = JHStudent.SelectByClassIDs(_ClassIDList);
            foreach (JHStudentRecord rec in studRecList)
            {
                if (rec.Status == StudentRecord.StudentStatus.一般)
                {
                    if (!StudentClassIDDict.ContainsKey(rec.ID))
                        StudentClassIDDict.Add(rec.ID, rec.RefClassID);

                    if (!ClassStudentIDDict.ContainsKey(rec.RefClassID))
                        ClassStudentIDDict.Add(rec.RefClassID,new List<StudentRecord>());

                    ClassStudentIDDict[rec.RefClassID].Add(rec);

                }
            }


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

        /// <summary>
        /// 依學年度學期取得學期成績
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="StudentIDList"></param>
        private void LoadStudentSemesterScore(int SchoolYear, int Semester, List<string> StudentIDList)
        {
            List<JHSemesterScoreRecord> SemesterScoreRecordLList = JHSemesterScore.SelectBySchoolYearAndSemester(StudentIDList, SchoolYear, Semester);

            StudentSemesterScoreRecordDict.Clear();
            ClassDomainNameDict.Clear();
            ClassSubjectNameDict.Clear();
            CanSelectDomainSubjectDict.Clear();
            CanSelectDomainSubjectDict.Add("彈性課程", new List<string>());

            foreach (JHSemesterScoreRecord rec in SemesterScoreRecordLList)
            {
                if (!StudentSemesterScoreRecordDict.ContainsKey(rec.RefStudentID))
                    StudentSemesterScoreRecordDict.Add(rec.RefStudentID, rec);

                if (StudentClassIDDict.ContainsKey(rec.RefStudentID))
                {
                    if (!ClassDomainNameDict.ContainsKey(StudentClassIDDict[rec.RefStudentID]))
                        ClassDomainNameDict.Add(StudentClassIDDict[rec.RefStudentID], new List<string>());

                    foreach (string key in rec.Domains.Keys)
                    {
                        if (!ClassDomainNameDict[StudentClassIDDict[rec.RefStudentID]].Contains(key))
                            ClassDomainNameDict[StudentClassIDDict[rec.RefStudentID]].Add(key);

                        if (!CanSelectDomainSubjectDict.ContainsKey(key))
                            CanSelectDomainSubjectDict.Add(key, new List<string>());
                    }

                    if (!ClassSubjectNameDict.ContainsKey(StudentClassIDDict[rec.RefStudentID]))
                        ClassSubjectNameDict.Add(StudentClassIDDict[rec.RefStudentID], new List<string>());

                    foreach (string key in rec.Subjects.Keys)
                    {
                        if (!ClassSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].Contains(key))
                            ClassSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].Add(key);
                    }

                    foreach (SubjectScore ss in rec.Subjects.Values)
                    {
                        if (ss.Domain == "")
                        {
                            if (!CanSelectDomainSubjectDict["彈性課程"].Contains(ss.Subject))
                                CanSelectDomainSubjectDict["彈性課程"].Add(ss.Subject);
                        }
                        else
                        {
                            if (CanSelectDomainSubjectDict.ContainsKey(ss.Domain))
                            {
                                if (!CanSelectDomainSubjectDict[ss.Domain].Contains(ss.Subject))
                                    CanSelectDomainSubjectDict[ss.Domain].Add(ss.Subject);
                            }
                        }

                    }
                }

            }

        }



        private void btnPrint_Click(object sender, EventArgs e)
        {
            //儲存設定檔
            SaveConfig();

            Global.SelOnlyDomainList.Clear();
            Global.SelOnlySubjectList.Clear();

            // 取得畫面上勾選領域科目
            SelDomainSubjectDict.Clear();
            SelDomainIdxDict.Clear();
            SelSubjectIdxDict.Clear();

            int sd = 1;
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

                if (!SelDomainIdxDict.ContainsKey(key))
                {
                    SelDomainIdxDict.Add(key, sd);
                    sd++;
                }
            }

            int ss = 1;
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

                if (!SelSubjectIdxDict.ContainsKey(lvi.Text))
                {
                    SelSubjectIdxDict.Add(lvi.Text, ss);
                    ss++;
                }
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

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 清空預設
            foreach (ListViewItem lvi in lvSubject.Items)
            {
                lvi.Checked = false;
            }


            // 清空預設
            foreach (ListViewItem lvi in lvDomain.Items)
            {
                lvi.Checked = false;
            }


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
                    _Configure.FailScoreMark = "";
                    _Configure.ReScoreMark = "";
                    _Configure.NeeedReScoreMark = "";
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
                    // 相關標示
                    txtFailScoreMark.Text = _Configure.FailScoreMark;
                    txtNeeedReScoreMark.Text = _Configure.NeeedReScoreMark;
                    txtReScoreMark.Text = _Configure.ReScoreMark;

                    if (_Configure.PrintSubjectList == null)
                        _Configure.PrintSubjectList = new List<string>();

                    if (_Configure.PrintDomainList == null)
                        _Configure.PrintDomainList = new List<string>();

                    // 解析科目
                    foreach (ListViewItem lvi in lvSubject.Items)
                    {
                        if (_Configure.PrintSubjectList.Contains(lvi.Text))
                        {
                            lvi.Checked = true;
                        }
                    }

                    // 解析領域
                    // 解析科目
                    foreach (ListViewItem lvi in lvDomain.Items)
                    {
                        if (_Configure.PrintDomainList.Contains(lvi.Text))
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

                }
            }
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
