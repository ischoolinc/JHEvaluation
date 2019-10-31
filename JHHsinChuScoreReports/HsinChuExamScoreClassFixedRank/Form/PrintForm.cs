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
using HsinChuExamScoreClassFixedRank.DAO;

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

        // 評分樣板比例
        Dictionary<string, decimal> ScorePercentageHSDict = new Dictionary<string, decimal>();

        List<ClassInfo> ClassInfoList = new List<ClassInfo>();

        // 錯誤訊息
        List<string> _ErrorList = new List<string>();

        // 畫面上所勾選領域科目        
        Dictionary<string, List<string>> SelDomainSubjectDict = new Dictionary<string, List<string>>();

        // 領域錯誤訊息
        List<string> _ErrorDomainNameList = new List<string>();

        BackgroundWorker bgWorkerLoadTemplate;
        BackgroundWorker bgWorkerReport;
        string SelSchoolYear = "";
        string SelSemester = "";
        string SelExamID = "";
        string SelExamName = "";
        int parseNumber = 0;
        List<string> _StudentIDList = new List<string>();

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
                    if (_ErrorDomainNameList.Count > 0)
                        sb.AppendLine(string.Join(",", _ErrorDomainNameList.ToArray()));

                    FISCA.Presentation.Controls.MsgBox.Show(sb.ToString(), "樣板內科目合併欄位不足", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                }

                #region 儲存檔案

                string reportName = "" + SelSchoolYear + "學年度第" + SelSemester + "學期" + SelExamName + "新竹班級評量成績單";

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



                FISCA.Presentation.MotherForm.SetStatusBarMessage("評量成績報表產生完成");
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
            }

        }

        private void BgWorkerReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 100)
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級評量成績完成");
            else
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級評量成績產生中...", e.ProgressPercentage);
        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {

            DataTable dtTable = new DataTable();

            // 取得相關資料
            bgWorkerReport.ReportProgress(1);

            // 每次合併後放入，最後再合成一張
            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.新竹科目班級評量成績單));

            _ErrorList.Clear();
            _ErrorDomainNameList.Clear();

            Global.SetDomainList();

            // 校名
            string SchoolName = K12.Data.School.ChineseName;
            // 校長
            string ChancellorChineseName = JHSchool.Data.JHSchoolInfo.ChancellorChineseName;
            // 教務主任
            string EduDirectorName = JHSchool.Data.JHSchoolInfo.EduDirectorName;

            // 取得班導師
            Dictionary<string, string> ClassTeacherNameDict = DataAccess.GetClassTeacherNameDictByClassID(_ClassIDList);

            // 取得評分樣板
            ScorePercentageHSDict = DataAccess.GetScorePercentageHS();

            // 取得班級學生
            ClassInfoList = DataAccess.GetClassStudentsByClassID(_ClassIDList);

            // 取得段考成績
            ClassInfoList = DataAccess.LoadClassStudentScore(ClassInfoList, SelSchoolYear, SelSemester, SelExamID, ScorePercentageHSDict, _ClassIDList);

            // 取得學生評量固定排名
            DataTable dtStudentExamRankMatrix = DataAccess.GetStudentExamRankMatrix(SelSchoolYear, SelSemester, SelExamID, _ClassIDList);
            // debug
            //dtStudentExamRankMatrix.TableName = "dtStudentExamRankMatrix";
            //dtStudentExamRankMatrix.WriteXml(Application.StartupPath + "\\tmp.xml");

            Dictionary<string, Dictionary<string, int>> StudentExamRankMatrixDict = new Dictionary<string, Dictionary<string, int>>();
            // 建立排名索引
            if (dtStudentExamRankMatrix != null && dtStudentExamRankMatrix.Rows.Count > 0)
            {

                foreach (DataRow dr in dtStudentExamRankMatrix.Rows)
                {
                    string student_id = dr["ref_student_id"].ToString();
                    // 定期評量/總計成績 加權總分 班排名
                    string type = dr["item_type"].ToString() + dr["item_name"].ToString() + dr["rank_type"].ToString();
                    if (!StudentExamRankMatrixDict.ContainsKey(student_id))
                        StudentExamRankMatrixDict.Add(student_id, new Dictionary<string, int>());

                    if (!StudentExamRankMatrixDict[student_id].ContainsKey(type))
                    {
                        int rank = 0;
                        int.TryParse(dr["rank"].ToString(), out rank);
                        StudentExamRankMatrixDict[student_id].Add(type, rank);
                    }
                }
            }

            // 計算學生總成績並將排名放入
            foreach (ClassInfo ci in ClassInfoList)
            {
                foreach (StudentInfo si in ci.Students)
                {
                    // 計算成績
                    si.CalScore();

                    if (StudentExamRankMatrixDict.ContainsKey(si.StudentID))
                    {
                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分班排名"))
                            si.ClassSumRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分班排名"))
                            si.ClassSumRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均班排名"))
                            si.ClassAvgRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均班排名"))
                            si.ClassAvgRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分年排名"))
                            si.YearSumRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分年排名"))
                            si.YearSumRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均年排名"))
                            si.YearAvgRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均年排名"))
                            si.YearAvgRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均年排名"];
                    }
                }

                // 計算各項目及格人數
                ci.CalClassPassCount();
                ci.CalClassAvgScore();
                ci.CalCredits();
                ci.CalClassDomainSubjectName();
            }

            // 合併用 DataTable
            // 產生合併欄位

            dtTable.Columns.Add("學校名稱");
            dtTable.Columns.Add("學年度");
            dtTable.Columns.Add("學期");
            dtTable.Columns.Add("試別名稱");
            dtTable.Columns.Add("班級");
            dtTable.Columns.Add("班導師");

            // 標頭用學分數
            foreach (string domainName in Global.DomainNameList)
            {
                if (SelDomainSubjectDict.ContainsKey(domainName))
                {
                    dtTable.Columns.Add(domainName + "領域學分");
                    dtTable.Columns.Add(domainName + "領域平均");
                    dtTable.Columns.Add(domainName + "領域及格人數");
                    for (int subj = 1; subj <= 12; subj++)
                    {
                        dtTable.Columns.Add(domainName + "領域_科目" + subj + "名稱");
                        dtTable.Columns.Add(domainName + "領域_科目" + subj + "學分");
                        dtTable.Columns.Add(domainName + "領域_科目" + subj + "平均");
                        dtTable.Columns.Add(domainName + "領域_科目" + subj + "及格人數");
                    }
                }

            }

            // 讀取最大人數
            int MaxStudentCount = 50;

            foreach (ClassInfo ci in ClassInfoList)
            {
                if (ci.Students.Count > MaxStudentCount)
                    MaxStudentCount = ci.Students.Count;
            }

            // 學生領域科目
            for (int studCot = 1; studCot <= MaxStudentCount; studCot++)
            {
                dtTable.Columns.Add("姓名" + studCot);
                dtTable.Columns.Add("座號" + studCot);
                dtTable.Columns.Add("總分" + studCot);
                dtTable.Columns.Add("加權總分" + studCot);
                dtTable.Columns.Add("平均" + studCot);
                dtTable.Columns.Add("加權平均" + studCot);
                dtTable.Columns.Add("總分班排名" + studCot);
                dtTable.Columns.Add("加權總分班排名" + studCot);
                dtTable.Columns.Add("平均班排名" + studCot);
                dtTable.Columns.Add("加權平均班排名" + studCot);
                dtTable.Columns.Add("總分年排名" + studCot);
                dtTable.Columns.Add("加權總分年排名" + studCot);
                dtTable.Columns.Add("平均年排名" + studCot);
                dtTable.Columns.Add("加權平均年排名" + studCot);

                foreach (string dName in Global.DomainNameList)
                {
                    if (SelDomainSubjectDict.ContainsKey(dName))
                    {
                        dtTable.Columns.Add(dName + "領域成績" + studCot);
                        dtTable.Columns.Add(dName + "領域學分" + studCot);

                        for (int i = 1; i <= 12; i++)
                        {
                            dtTable.Columns.Add(dName + "領域_科目名稱" + studCot + "_" + i);
                            dtTable.Columns.Add(dName + "領域_科目成績" + studCot + "_" + i);
                            dtTable.Columns.Add(dName + "領域_科目學分" + studCot + "_" + i);

                        }
                    }
                }
            }

            // 各領域科目 及格人數與平均
            foreach (string dName in Global.DomainNameList)
            {
                if (SelDomainSubjectDict.ContainsKey(dName))
                {
                    dtTable.Columns.Add(dName + "班級領域平均");
                    dtTable.Columns.Add(dName + "班級領域及格人數");

                    for (int i = 1; i <= 12; i++)
                    {
                        dtTable.Columns.Add(dName + "班級領域_科目平均" + i);
                        dtTable.Columns.Add(dName + "班級領域_科目及格人數" + i);
                    }

                }
            }

            bgWorkerReport.ReportProgress(30);

            // 填入資料
            foreach (ClassInfo ci in ClassInfoList)
            {
                DataRow row = dtTable.NewRow();
                row["學校名稱"] = SchoolName;
                row["學年度"] = SelSchoolYear;
                row["學期"] = SelSemester;
                row["試別名稱"] = SelExamName;
                row["班級"] = ci.ClassName;

                if (ClassTeacherNameDict.ContainsKey(ci.ClassID))
                    row["班導師"] = ClassTeacherNameDict[ci.ClassID];

                foreach (string dName in ci.ClassDomainSubjectNameDict.Keys)
                {
                    if (SelDomainSubjectDict.ContainsKey(dName))
                    {
                        string key1 = dName + "領域學分";
                        if (ci.CreditsDict.ContainsKey(key1))
                        {
                            row[key1] = string.Join(",", ci.CreditsDict[key1].ToArray());
                        }

                        int ss = 1;
                        foreach (string ssName in SelDomainSubjectDict[dName])
                        {
                            foreach (string subjName in ci.ClassDomainSubjectNameDict[dName])
                            {
                                if (ssName == subjName)
                                {
                                    string key2 = dName + "領域_" + subjName + "科目學分";

                                    if (ci.CreditsDict.ContainsKey(key2))
                                    {
                                        string key2_1 = dName + "領域_科目" + ss + "學分";
                                        string key2_2 = dName + "領域_科目" + ss + "名稱";
                                        row[key2_1] = string.Join(",", ci.CreditsDict[key2].ToArray());
                                        row[key2_2] = subjName;
                                    }
                                }
                            }
                            ss++;
                        }
                    }

                }

                // 填入學生資料
                int studCot = 1;
                foreach (StudentInfo si in ci.Students)
                {
                    row["姓名" + studCot] = si.Name;
                    row["座號" + studCot] = si.SeatNo;
                    if (si.SumScore.HasValue)
                        row["總分" + studCot] = Math.Round(si.SumScore.Value, parseNumber, MidpointRounding.AwayFromZero);
                    if (si.SumScoreA.HasValue)
                        row["加權總分" + studCot] = Math.Round(si.SumScoreA.Value, parseNumber, MidpointRounding.AwayFromZero);

                    if (si.AvgScore.HasValue)
                        row["平均" + studCot] = Math.Round(si.AvgScore.Value, parseNumber, MidpointRounding.AwayFromZero);

                    if (si.AvgScoreA.HasValue)
                        row["加權平均" + studCot] = Math.Round(si.AvgScoreA.Value, parseNumber, MidpointRounding.AwayFromZero);

                    if (si.ClassSumRank.HasValue)
                        row["總分班排名" + studCot] = si.ClassSumRank.Value;

                    if (si.ClassSumRankA.HasValue)
                        row["加權總分班排名" + studCot] = si.ClassSumRankA.Value;

                    if (si.ClassAvgRank.HasValue)
                        row["平均班排名" + studCot] = si.ClassAvgRank.Value;

                    if (si.ClassAvgRankA.HasValue)
                        row["加權平均班排名" + studCot] = si.ClassAvgRankA.Value;

                    if (si.YearSumRank.HasValue)
                        row["總分年排名" + studCot] = si.YearSumRank.Value;

                    if (si.YearSumRankA.HasValue)
                        row["加權總分年排名" + studCot] = si.YearSumRankA.Value;

                    if (si.YearAvgRank.HasValue)
                        row["平均年排名" + studCot] = si.YearAvgRank.Value;

                    if (si.YearAvgRankA.HasValue)
                        row["加權平均年排名" + studCot] = si.YearAvgRankA.Value;

                    // 領域科目成績
                    foreach (DomainInfo di in si.DomainInfoList)
                    {
                        if (SelDomainSubjectDict.ContainsKey(di.Name))
                        {
                            string d1 = di.Name + "領域成績" + studCot;
                            if (dtTable.Columns.Contains(d1))
                            {
                                if (di.Score.HasValue)
                                    row[d1] = Math.Round(di.Score.Value, parseNumber, MidpointRounding.AwayFromZero);
                            }
                            string d2 = di.Name + "領域學分" + studCot;
                            if (dtTable.Columns.Contains(d2))
                            {
                                if (di.Credit.HasValue)
                                {
                                    row[d2] = di.Credit.Value;
                                }
                            }

                            int subjCot = 1;

                            foreach (string ssName in SelDomainSubjectDict[di.Name])
                            {
                                foreach (SubjectInfo subj in di.SubjectInfoList)
                                {
                                    if (subj.Name == ssName)
                                    {
                                        string s1Key = di.Name + "領域_科目名稱" + studCot + "_" + subjCot;
                                        string s2Key = di.Name + "領域_科目成績" + studCot + "_" + subjCot;
                                        string s3Key = di.Name + "領域_科目學分" + studCot + "_" + subjCot;

                                        if (dtTable.Columns.Contains(s1Key))
                                        {
                                            row[s1Key] = subj.Name;
                                        }

                                        if (dtTable.Columns.Contains(s2Key))
                                        {
                                            if (dtTable.Columns.Contains(s2Key))
                                                if (subj.Score.HasValue)
                                                    row[s2Key] = subj.Score.Value;
                                        }
                                        if (dtTable.Columns.Contains(s3Key))
                                        {
                                            if (subj.Credit.HasValue)
                                                row[s3Key] = subj.Credit.Value;
                                        }
                                    }
                                }
                                subjCot++;
                            }
                        }
                    }
                    studCot += 1;
                }

                // 處理班級平均
                // 處理班級及格人數
                foreach (string dName in SelDomainSubjectDict.Keys)
                {
                    if (ci.ClassAvgScoreDict.ContainsKey(dName + "平均"))
                        row[dName + "班級領域平均"] = ci.ClassAvgScoreDict[dName + "平均"];
                    if (ci.ClassScorePassCountDict.ContainsKey(dName))
                        row[dName + "班級領域及格人數"] = ci.ClassScorePassCountDict[dName];

                    int subjCot = 1;
                    foreach (string sName in SelDomainSubjectDict[dName])
                    {
                        string key = dName + "_" + sName;
                        if (ci.ClassAvgScoreDict.ContainsKey(key + "平均"))
                        {
                            row[dName + "班級領域_科目平均" + subjCot] = ci.ClassAvgScoreDict[key + "平均"];
                        }

                        if (ci.ClassScorePassCountDict.ContainsKey(key))
                            row[dName + "班級領域_科目及格人數" + subjCot] = ci.ClassScorePassCountDict[key];

                        subjCot++;
                    }
                }

                dtTable.Rows.Add(row);
            }

            docTemplate.MailMerge.Execute(dtTable);
            docTemplate.MailMerge.RemoveEmptyParagraphs = true;
            docTemplate.MailMerge.DeleteFields();

            #region Word 合併列印

            try
            {
                e.Result = new object[] { docTemplate };
            }
            catch (Exception exow)
            {
                throw exow;
            }

            #endregion


            //dtTable.TableName = "dtTable";
            //dtTable.WriteXml(Application.StartupPath + "\\dtT.xml");




            //           dtTable.WriteXmlSchema(Application.StartupPath + "\\dtTsc.xml");
            bgWorkerReport.ReportProgress(100);
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

            int idx = 0;
            foreach (ExamRecord exName in _exams)
            {
                if (exName.ID == _Configure.ExamRecordID)
                    cboExam.SelectedIndex = idx;

                idx++;
            }

            if (!string.IsNullOrEmpty(_Configure.SelSetConfigName))
                cboConfigure.Text = userSelectConfigName;

            btnSaveConfig.Enabled = btnPrint.Enabled = true;

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

            bgWorkerLoadTemplate.ReportProgress(30);

            // 沒有設定檔，建立預設設定檔
            if (_UDTConfigList.Count < 2)
            {

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
                        case "領域成績單":
                            cn.Template = new Document(new MemoryStream(Properties.Resources.新竹領域班級評量成績單));
                            break;

                        case "科目成績單":
                            cn.Template = new Document(new MemoryStream(Properties.Resources.新竹科目班級評量成績單));
                            break;
                    }

                    if (cn.Template == null)
                        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹班級評量成績單樣板));
                    cn.Encode();
                    cn.Save();
                }
                if (_UDTConfigList.Count > 0)
                    DAO.UDTTransfer.InsertConfigData(_UDTConfigList);
            }
            else
            {

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
            SaveConfig();
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

            // 四捨五入進位
            if (!int.TryParse(cboParseNumber.Text, out parseNumber))
            {
                MsgBox.Show("平均計算至小數點後需要輸入整數!");
                return;
            }

            _Configure.ParseNumber = parseNumber;

            foreach (ExamRecord exm in _exams)
            {
                if (exm.Name == cboExam.Text)
                {
                    _Configure.ExamRecord = exm;
                    break;
                }
            }

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
            for (int i = 0; i <= 3; i++)
                cboParseNumber.Items.Add(i);

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
            SelSchoolYear = cboSchoolYear.Text;
            SelSemester = cboSemester.Text;

            // 產生合併欄位總表
            lnkViewMapColumns.Enabled = false;

            if (string.IsNullOrEmpty(cboExam.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇試別!");
                lnkViewMapColumns.Enabled = true;
                return;
            }
            else
            {
                bool isEr = true;
                foreach (ExamRecord ex in _exams)
                    if (ex.Name == cboExam.Text)
                    {
                        SelExamID = ex.ID;

                        isEr = false;
                        break;
                    }

                if (isEr)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("試別錯誤，請重新選擇!");
                    return;
                }
            }

            SetDomainList();

            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
        }

        private void SetDomainList()
        {
            Global._SelSchoolYear = SelSchoolYear;
            Global._SelSemester = SelSemester; ;
            //   Global._SelStudentIDList = _StudentIDList;
            Global._SelExamID = SelExamID;
            Global._SelClassIDList = _ClassIDList;
            Global.SetDomainList();
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

            // 取得試別編號
            foreach (ExamRecord er in _exams)
            {
                if (er.Name == cboExam.Text)
                {
                    SelExamID = er.ID;
                    SelExamName = er.Name;
                }
            }

            SelSchoolYear = cboSchoolYear.Text;
            SelSemester = cboSemester.Text;

            // 四捨五入進位
            if (!int.TryParse(cboParseNumber.Text, out parseNumber))
            {
                MsgBox.Show("平均計算至小數點後需要輸入整數!");
                return;
            }

            Global.parseNumebr = parseNumber;
            Global._SelExamID = SelExamID;
            Global._SelSchoolYear = SelSchoolYear;
            Global._SelSemester = SelSemester;


            SetDomainList();

            // 列印報表
            bgWorkerReport.RunWorkerAsync();


        }

        private void cboExam_SelectedIndexChanged(object sender, EventArgs e)
        {

            foreach (ExamRecord er in _exams)
            {
                if (er.Name == cboExam.Text)
                {
                    SelExamID = er.ID;
                    SelExamName = er.Name;
                }
            }

            LoadDomainSubject();
        }

        private void LoadDomainSubject()
        {
            // 清空畫面上領域科目
            lvDomain.Items.Clear();
            lvSubject.Items.Clear();
            if (CanSelectExamDomainSubjectDict.ContainsKey(SelExamID))
            {
                foreach (string dName in CanSelectExamDomainSubjectDict[SelExamID].Keys)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Tag = dName;
                    lvi.Name = dName;
                    lvi.Text = dName;
                    lvDomain.Items.Add(lvi);

                    foreach (string sName in CanSelectExamDomainSubjectDict[SelExamID][dName])
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
                                _Configure.ExamRecordID = rec.ID;
                                break;
                            }
                        }

                    }
                    _Configure.ParseNumber = parseNumber;
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
                    cboExam.Text = "";
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

                    Global.parseNumebr = parseNumber = _Configure.ParseNumber;
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
                    cboExam.SelectedIndex = -1;

                }
            }
        }
    }
}
