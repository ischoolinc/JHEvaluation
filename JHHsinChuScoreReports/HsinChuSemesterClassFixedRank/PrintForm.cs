using Aspose.Words;
using FISCA.Presentation.Controls;
using HsinChuSemesterClassFixedRank.DAO;
using JHSchool.Data;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JHSchool.Evaluation.Calculation;

namespace HsinChuSemesterClassFixedRank
{
    public partial class PrintForm : BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        List<string> _ClassIDList = new List<string>();
        Dictionary<string, ClassRecord> ClassRecordDict = new Dictionary<string, ClassRecord>();
        // 一般狀態學生與班級ID
        Dictionary<string, string> StudentClassIDDict = new Dictionary<string, string>();

        // 班級學生ID
        Dictionary<string, List<StudentRecord>> ClassStudentDict = new Dictionary<string, List<StudentRecord>>();

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

        // 班級可選領域科目整理
        Dictionary<string, Dictionary<string, List<string>>> ClassDomainSubjectNameDict = new Dictionary<string, Dictionary<string, List<string>>>();


        // 錯誤訊息
        List<string> _ErrorList = new List<string>();

        /// <summary>
        /// 第幾位四捨五入
        /// </summary>
      //  int ParseNumber = 0;

        // 領域錯誤訊息
        List<string> _ErrorDomainNameList = new List<string>();

        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();
        public Configure _Configure { get; private set; }

        // 畫面上所勾選領域科目        
        Dictionary<string, Dictionary<string, List<string>>> SelDomainSubjectDict = new Dictionary<string, Dictionary<string, List<string>>>();

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
                Document doc = (Document)e.Result;

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
                    if (doc != null)
                        document = doc;
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

            try
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

                Dictionary<string, Dictionary<string, Dictionary<string, string>>> ClassSemsScoreRankMatrixDataValueDict = DataAccess.GetClassSemsScoreRankMatrixDataValue(_Configure.SchoolYear, _Configure.Semester, _ClassIDList);
                bgWorkerReport.ReportProgress(10);
                // debug
                //StreamWriter sw = new StreamWriter(Application.StartupPath + "\\學生五標排名組距.txt");
                //StreamWriter swClass = new StreamWriter(Application.StartupPath + "\\班級五標排名組距.txt");
                //List<string> tmpList = new List<string>();
                //List<string> tmpClassList = new List<string>();

                //foreach (string key in SemsScoreRankMatrixDataValueDict.Keys)
                //{
                //    foreach (string key1 in SemsScoreRankMatrixDataValueDict[key].Keys)
                //    {
                //        if (!tmpList.Contains(key1))
                //            tmpList.Add(key1);
                //    }
                //}

                //foreach (string key in ClassSemsScoreRankMatrixDataValueDict.Keys)
                //{
                //    foreach (string k1 in ClassSemsScoreRankMatrixDataValueDict[key].Keys)
                //    {
                //        if (!tmpClassList.Contains(k1))
                //            tmpClassList.Add(k1);
                //    }
                //}

                //foreach (string key in tmpList)
                //    sw.WriteLine(key);
                //sw.Close();

                //foreach (string key in tmpClassList)
                //    swClass.WriteLine(key);
                //swClass.Close();


                // 所領域科目排序
                foreach (string classID in SelDomainSubjectDict.Keys)
                {
                    foreach (string dName in SelDomainSubjectDict[classID].Keys)
                    {
                        SelDomainSubjectDict[classID][dName].Sort(new StringComparer("國文", "英文", "數學", "理化", "生物", "社會", "物理", "化學", "歷史", "地理", "公民"));
                    }
                }




                #region 產生合併欄位 Columns

                //// 全部
                //List<string> r2List = new List<string>();
                //r2List.Add("rank");
                //r2List.Add("matrix_count");
                //r2List.Add("pr");
                //r2List.Add("percentile");
                //r2List.Add("avg_top_25");
                //r2List.Add("avg_top_50");
                //r2List.Add("avg");
                //r2List.Add("avg_bottom_50");
                //r2List.Add("avg_bottom_25");
                //r2List.Add("level_gte100");
                //r2List.Add("level_90");
                //r2List.Add("level_80");
                //r2List.Add("level_70");
                //r2List.Add("level_60");
                //r2List.Add("level_50");
                //r2List.Add("level_40");
                //r2List.Add("level_30");
                //r2List.Add("level_20");
                //r2List.Add("level_10");
                //r2List.Add("level_lt10");

                // 加入學生常用 排名、PR、百分比
                List<string> r2List = new List<string>();
                r2List.Add("rank");
                r2List.Add("pr");
                r2List.Add("percentile");

                // 班級排名
                List<string> cr2List = new List<string>();
                cr2List.Add("matrix_count");
                cr2List.Add("avg_top_25");
                cr2List.Add("avg_top_50");
                cr2List.Add("avg");
                cr2List.Add("avg_bottom_50");
                cr2List.Add("avg_bottom_25");
                cr2List.Add("level_gte100");
                cr2List.Add("level_90");
                cr2List.Add("level_80");
                cr2List.Add("level_70");
                cr2List.Add("level_60");
                cr2List.Add("level_50");
                cr2List.Add("level_40");
                cr2List.Add("level_30");
                cr2List.Add("level_20");
                cr2List.Add("level_10");
                cr2List.Add("level_lt10");

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
                dtTable.Columns.Add("類別1排名名稱");
                dtTable.Columns.Add("類別2排名名稱");

                int maxStudent = 30;

                foreach (string key in ClassStudentDict.Keys)
                {
                    if (ClassStudentDict[key].Count > maxStudent)
                        maxStudent = ClassStudentDict[key].Count;
                }

                List<string> dNameList = new List<string>();
                foreach (string cid in SelDomainSubjectDict.Keys)
                {
                    foreach (string name in SelDomainSubjectDict[cid].Keys)
                    {
                        if (!dNameList.Contains(name))
                            dNameList.Add(name);
                    }
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
                            // 先放排名   --五標、PR、百分比、組距
                            foreach (string r2 in r2List)
                            {
                                dtTable.Columns.Add("學生" + studIdx + "_" + r1 + "_" + rk + "_" + r2);
                            }
                        }

                    }

                    // 產生領域_科目1,2,3合併欄位
                    // 領域1,2 合併欄位             
                    // 取各班最大聯集
                    foreach (string dName in dNameList)
                    {
                        dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域_學分");
                        dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域_成績");
                        dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域(原始)_成績");


                        for (int si = 1; si <= 12; si++)
                        {
                            dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域_科目" + si + "_名稱");
                            dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域_科目" + si + "_學分");
                            dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域_科目" + si + "_成績");
                            dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域(原始)_科目" + si + "_成績");
                        }

                        foreach (string rk in rkList)
                        {
                            // 排名 
                            foreach (string r2 in r2List)
                            {
                                dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域_" + rk + "_" + r2);
                                dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域(原始)_" + rk + "_" + r2);

                                for (int si = 1; si <= 12; si++)
                                {
                                    dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域_科目" + si + "_" + rk + "_" + r2);
                                    dtTable.Columns.Add("學生" + studIdx + "_" + dName + "領域(原始)_科目" + si + "_" + rk + "_" + r2);
                                }
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
                        dtTable.Columns.Add("學生" + studIdx + "_領域" + dIdx + "_需補考標示");
                        dtTable.Columns.Add("學生" + studIdx + "_領域" + dIdx + "_補考成績標示");
                        dtTable.Columns.Add("學生" + studIdx + "_領域" + dIdx + "_不及格標示");


                        foreach (string rk in rkList)
                        {
                            // 排名 --五標、PR、百分比、組距
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
                        dtTable.Columns.Add("學生" + studIdx + "_科目" + sIdx + "_需補考標示");
                        dtTable.Columns.Add("學生" + studIdx + "_科目" + sIdx + "_補考成績標示");
                        dtTable.Columns.Add("學生" + studIdx + "_科目" + sIdx + "_不及格標示");


                        foreach (string rk in rkList)
                        {
                            // 排名 -- 五標、PR、百分比、組距
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
                        foreach (string r2 in cr2List)
                        {
                            dtTable.Columns.Add("班級_" + r1 + "_" + rk + "_" + r2);
                        }
                    }

                }


                // 領域1,2 合併欄位
                for (int dIdx = 1; dIdx <= SelDomainIdxDict.Keys.Count; dIdx++)
                {
                    // 班級_領域1_成績
                    dtTable.Columns.Add("班級_領域" + dIdx + "_名稱");
                    dtTable.Columns.Add("班級_領域" + dIdx + "_學分");

                    // 班級_領域1_平均,及格人數
                    dtTable.Columns.Add("班級_領域" + dIdx + "_平均");
                    dtTable.Columns.Add("班級_領域(原始)" + dIdx + "_平均");
                    dtTable.Columns.Add("班級_領域" + dIdx + "_及格人數");
                    dtTable.Columns.Add("班級_領域(原始)" + dIdx + "_及格人數");

                    foreach (string rk in rkList)
                    {
                        // 五標、PR、百分比、組距
                        foreach (string r2 in cr2List)
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
                    dtTable.Columns.Add("班級_科目" + sIdx + "_學分");

                    // 班級_科目1_平均,及格人數
                    dtTable.Columns.Add("班級_科目" + sIdx + "_平均");
                    dtTable.Columns.Add("班級_科目(原始)" + sIdx + "_平均");
                    dtTable.Columns.Add("班級_科目" + sIdx + "_及格人數");
                    dtTable.Columns.Add("班級_科目(原始)" + sIdx + "_及格人數");

                    foreach (string rk in rkList)
                    {
                        // 五標、PR、百分比、組距
                        foreach (string r2 in cr2List)
                        {
                            dtTable.Columns.Add("班級_科目" + sIdx + "_" + rk + "_" + r2);
                            dtTable.Columns.Add("班級_科目(原始)" + sIdx + "_" + rk + "_" + r2);
                        }
                    }

                }

                // 班級領域 科目 1,2,3
                foreach (string dName in dNameList)
                {

                    foreach (string rk in rkList)
                    {
                        foreach (string r2 in cr2List)
                        {
                            dtTable.Columns.Add("班級_" + dName + "領域_" + rk + "_" + r2);
                            dtTable.Columns.Add("班級_" + dName + "領域(原始)_" + rk + "_" + r2);
                        }
                    }


                    for (int si = 1; si <= 12; si++)
                    {
                        dtTable.Columns.Add("班級_" + dName + "領域_科目" + si + "_名稱");
                        dtTable.Columns.Add("班級_" + dName + "領域_科目" + si + "_學分");
                        dtTable.Columns.Add("班級_" + dName + "領域_科目" + si + "_平均");
                        dtTable.Columns.Add("班級_" + dName + "領域_科目" + si + "_及格人數");
                        dtTable.Columns.Add("班級_" + dName + "領域(原始)_科目" + si + "_平均");
                        dtTable.Columns.Add("班級_" + dName + "領域(原始)_科目" + si + "_及格人數");

                        foreach (string rk in rkList)
                        {
                            foreach (string r2 in cr2List)
                            {
                                dtTable.Columns.Add("班級_" + dName + "領域_科目" + si + "_" + rk + "_" + r2);
                                dtTable.Columns.Add("班級_" + dName + "領域(原始)_科目" + si + "_" + rk + "_" + r2);
                            }
                        }

                    }


                }

                #region 各位學生即時計算加權總分、總分、加權平均、平均

                List<string> tmpStudScoreItemList = new List<string>();
                tmpStudScoreItemList.Add("領域總計_加權總分");
                tmpStudScoreItemList.Add("領域總計_總分");
                tmpStudScoreItemList.Add("領域總計_加權平均");
                tmpStudScoreItemList.Add("領域總計_平均");
                tmpStudScoreItemList.Add("科目總計_加權總分");
                tmpStudScoreItemList.Add("科目總計_總分");
                tmpStudScoreItemList.Add("科目總計_加權平均");
                tmpStudScoreItemList.Add("科目總計_平均");
                tmpStudScoreItemList.Add("領域(原始)總計_加權總分");
                tmpStudScoreItemList.Add("領域(原始)總計_總分");
                tmpStudScoreItemList.Add("領域(原始)總計_加權平均");
                tmpStudScoreItemList.Add("領域(原始)總計_平均");
                tmpStudScoreItemList.Add("科目(原始)總計_加權總分");
                tmpStudScoreItemList.Add("科目(原始)總計_總分");
                tmpStudScoreItemList.Add("科目(原始)總計_加權平均");
                tmpStudScoreItemList.Add("科目(原始)總計_平均");

                for (int studIdx = 1; studIdx <= maxStudent; studIdx++)
                {
                    foreach (string item in tmpStudScoreItemList)
                    {
                        dtTable.Columns.Add("學生" + studIdx + "_" + item);
                    }
                }
                #endregion



                //StreamWriter sw1 = new StreamWriter(Application.StartupPath + "\\合併欄位.txt");
                //StringBuilder sb1 = new StringBuilder();
                //foreach (DataColumn dc in dtTable.Columns)
                //    sb1.AppendLine(dc.Caption);

                //sw1.Write(sb1.ToString());
                //sw1.Close();

                bgWorkerReport.ReportProgress(30);
                #endregion

                #region 填入資料

                // 各領域、各科目 學分數、平均、及格人數
                Dictionary<string, List<decimal>> tmpCreditDict = new Dictionary<string, List<decimal>>();
                Dictionary<string, ScoreItemR> tmpScoreItemRDict = new Dictionary<string, ScoreItemR>();

                Dictionary<string, ScoreItemC> tmpScoreItemCDict = new Dictionary<string, ScoreItemC>();


                #region 取得學生成績計算規則
                ScoreCalculator defaultScoreCalculator = new ScoreCalculator(null);

                //key: ScoreCalcRuleID
                Dictionary<string, ScoreCalculator> calcCache = new Dictionary<string, ScoreCalculator>();
                //key: StudentID, val: ScoreCalcRuleID
                Dictionary<string, string> calcIDCache = new Dictionary<string, string>();
                List<string> scoreCalcRuleIDList = new List<string>();
                foreach (string class_id in ClassTeacherNameDict.Keys)
                {
                    foreach (StudentRecord student in ClassStudentDict[class_id])
                    {
                        //calcCache.Add(student.ID, new ScoreCalculator(student.ScoreCalcRule));
                        string calcID = string.Empty;
                        if (!string.IsNullOrEmpty(student.OverrideScoreCalcRuleID))
                            calcID = student.OverrideScoreCalcRuleID;
                        else if (student.Class != null && !string.IsNullOrEmpty(student.Class.RefScoreCalcRuleID))
                            calcID = student.Class.RefScoreCalcRuleID;

                        if (!string.IsNullOrEmpty(calcID))
                            calcIDCache.Add(student.ID, calcID);
                    }
                    foreach (JHScoreCalcRuleRecord record in JHScoreCalcRule.SelectByIDs(calcIDCache.Values))
                    {
                        if (!calcCache.ContainsKey(record.ID))
                            calcCache.Add(record.ID, new ScoreCalculator(record));
                    }
                }



                #endregion


                bgWorkerReport.ReportProgress(40);
                // 排序後班級
                foreach (string class_id in ClassTeacherNameDict.Keys)
                {
                    DataRow row = dtTable.NewRow();

                    tmpCreditDict.Clear();
                    tmpScoreItemRDict.Clear();
                    tmpScoreItemCDict.Clear();

                    row["學校名稱"] = K12.Data.School.ChineseName;
                    row["學年度"] = _Configure.SchoolYear;
                    row["學期"] = _Configure.Semester;
                    if (ClassRecordDict.ContainsKey(class_id))
                        row["班級"] = ClassRecordDict[class_id].Name;

                    row["班導師"] = ClassTeacherNameDict[class_id];

                    if (DataAccess.ClassTag1Dict.ContainsKey(class_id))
                        row["類別1排名名稱"] = DataAccess.ClassTag1Dict[class_id];

                    if (DataAccess.ClassTag2Dict.ContainsKey(class_id))
                        row["類別2排名名稱"] = DataAccess.ClassTag2Dict[class_id];


                    // 填入學生資料
                    if (ClassStudentDict.ContainsKey(class_id))
                    {
                        int studIdx = 1;
                        foreach (StudentRecord studRec in ClassStudentDict[class_id])
                        {

                            // 成績計算規則
                            ScoreCalculator studentCalculator = defaultScoreCalculator;
                            if (calcIDCache.ContainsKey(studRec.ID) && calcCache.ContainsKey(calcIDCache[studRec.ID]))
                                studentCalculator = calcCache[calcIDCache[studRec.ID]];


                            row["姓名" + studIdx] = studRec.Name;
                            if (studRec.SeatNo.HasValue)
                                row["座號" + studIdx] = studRec.SeatNo.Value;

                            if (StudentSemesterScoreRecordDict.ContainsKey(studRec.ID))
                            {
                                JHSemesterScoreRecord studSemsScoreRec = StudentSemesterScoreRecordDict[studRec.ID];
                                // 處理學期總計成績、排名、五標、組距
                                // 學生1_課程學習總成績_成績
                                if (studSemsScoreRec.CourseLearnScore.HasValue)
                                    row["學生" + studIdx + "_課程學習總成績_成績"] = studSemsScoreRec.CourseLearnScore.Value;

                                if (studSemsScoreRec.CourseLearnScoreOrigin.HasValue)
                                    row["學生" + studIdx + "_課程學習總成績(原始)_成績"] = studSemsScoreRec.CourseLearnScoreOrigin.Value;

                                // 學生1_學習領域總成績_成績
                                if (studSemsScoreRec.LearnDomainScore.HasValue)
                                    row["學生" + studIdx + "_學習領域總成績_成績"] = studSemsScoreRec.LearnDomainScore.Value;

                                if (studSemsScoreRec.LearnDomainScoreOrigin.HasValue)
                                    row["學生" + studIdx + "_學習領域總成績(原始)_成績"] = studSemsScoreRec.LearnDomainScoreOrigin.Value;

                                // 處理固定排名相關
                                // 學生1_課程學習總成績_班排名_rank
                                if (SemsScoreRankMatrixDataValueDict.ContainsKey(studRec.ID))
                                {
                                    // 班、年..
                                    foreach (string rk in rkList)
                                    {
                                        // 學期/總計成績_課程學習總成績_年排名
                                        if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/總計成績_課程學習總成績_" + rk))
                                        {
                                            foreach (string r2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績_課程學習總成績_" + rk].ContainsKey(r2))
                                                {
                                                    row["學生" + studIdx + "_課程學習總成績_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績_課程學習總成績_" + rk][r2];
                                                }
                                            }
                                        }

                                        // 學期/總計成績_課程學習總成績_年排名
                                        if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/總計成績(原始)_課程學習總成績_" + rk))
                                        {
                                            foreach (string r2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績(原始)_課程學習總成績_" + rk].ContainsKey(r2))
                                                {
                                                    row["學生" + studIdx + "_課程學習總成績(原始)_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績(原始)_課程學習總成績_" + rk][r2];
                                                }
                                            }
                                        }

                                        // 學期/總計成績_學習領域總成績_年排名
                                        if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/總計成績_學習領域總成績_" + rk))
                                        {
                                            foreach (string r2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績_學習領域總成績_" + rk].ContainsKey(r2))
                                                {
                                                    row["學生" + studIdx + "_學習領域總成績_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績_學習領域總成績_" + rk][r2];
                                                }
                                            }
                                        }

                                        // 學期/總計成績_學習領域總成績_年排名
                                        if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/總計成績(原始)_學習領域總成績_" + rk))
                                        {
                                            foreach (string r2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績(原始)_學習領域總成績_" + rk].ContainsKey(r2))
                                                {
                                                    row["學生" + studIdx + "_學習領域總成績(原始)_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/總計成績(原始)_學習領域總成績_" + rk][r2];
                                                }
                                            }
                                        }

                                    }
                                }

                                // 領域 科目1,2,3
                                if (SelDomainSubjectDict.ContainsKey(class_id))
                                {
                                    foreach (string dName in SelDomainSubjectDict[class_id].Keys)
                                    {
                                        if (studSemsScoreRec.Domains.ContainsKey(dName))
                                        {
                                            // 學生5_領域2_名稱

                                            if (studSemsScoreRec.Domains[dName].Credit.HasValue)
                                                row["學生" + studIdx + "_" + dName + "領域_學分"] = studSemsScoreRec.Domains[dName].Credit.Value;

                                            if (studSemsScoreRec.Domains[dName].Score.HasValue)
                                                row["學生" + studIdx + "_" + dName + "領域_成績"] = studSemsScoreRec.Domains[dName].Score.Value;

                                            if (studSemsScoreRec.Domains[dName].ScoreOrigin.HasValue)
                                                row["學生" + studIdx + "_" + dName + "領域(原始)_成績"] = studSemsScoreRec.Domains[dName].ScoreOrigin.Value;

                                        }
                                        // 科目
                                        int subjIdx = 1;
                                        foreach (string sName in SelDomainSubjectDict[class_id][dName])
                                        {
                                            if (studSemsScoreRec.Subjects.ContainsKey(sName))
                                            {
                                                // 學生5_科目2_名稱
                                                row["學生" + studIdx + "_" + dName + "領域_科目" + subjIdx + "_名稱"] = sName;
                                                if (studSemsScoreRec.Subjects[sName].Credit.HasValue)
                                                {
                                                    row["學生" + studIdx + "_" + dName + "領域_科目" + subjIdx + "_學分"] = studSemsScoreRec.Subjects[sName].Credit.Value;

                                                    // 班級_語文領域_科目1_學分
                                                    string kk = "班級_" + dName + "領域_科目" + subjIdx + "_學分";

                                                    if (!tmpCreditDict.ContainsKey(kk))
                                                        tmpCreditDict.Add(kk, new List<decimal>());

                                                    if (!tmpCreditDict[kk].Contains(studSemsScoreRec.Subjects[sName].Credit.Value))
                                                        tmpCreditDict[kk].Add(studSemsScoreRec.Subjects[sName].Credit.Value);

                                                }


                                                if (studSemsScoreRec.Subjects[sName].Score.HasValue)
                                                {
                                                    row["學生" + studIdx + "_" + dName + "領域_科目" + subjIdx + "_成績"] = studSemsScoreRec.Subjects[sName].Score.Value;

                                                    // 班級_語文領域_科目1_平均
                                                    // 班級_語文領域_科目1_及格人數
                                                    // 班級_語文領域(原始)_科目1_平均
                                                    // 班級_語文領域(原始)_科目1_及格人數

                                                    string kka1 = "班級_" + dName + "領域_科目" + subjIdx + "_平均";
                                                    if (!tmpScoreItemRDict.ContainsKey(kka1))
                                                    {
                                                        ScoreItemR sr = new ScoreItemR();
                                                        sr.ScoreKey = kka1;
                                                        sr.ScoreOriginKey = "班級_" + dName + "領域(原始)_科目" + subjIdx + "_平均";
                                                        sr.PassCountKey = "班級_" + dName + "領域_科目" + subjIdx + "_及格人數";
                                                        sr.PassOriginKey = "班級_" + dName + "領域(原始)_科目" + subjIdx + "_及格人數";

                                                        tmpScoreItemRDict.Add(kka1, sr);
                                                    }

                                                    tmpScoreItemRDict[kka1].AddScore(studSemsScoreRec.Subjects[sName].Score.Value);




                                                    // 加入原始成績
                                                    if (studSemsScoreRec.Subjects[sName].ScoreOrigin.HasValue)
                                                    {
                                                        tmpScoreItemRDict[kka1].AddScoreOrigin(studSemsScoreRec.Subjects[sName].ScoreOrigin.Value);
                                                    }
                                                }

                                                if (studSemsScoreRec.Subjects[sName].ScoreOrigin.HasValue)
                                                    row["學生" + studIdx + "_" + dName + "領域(原始)_科目" + subjIdx + "_成績"] = studSemsScoreRec.Subjects[sName].ScoreOrigin.Value;



                                            }

                                            subjIdx++;
                                        }
                                    }
                                }



                                // 處理學期領域成績、排名、五標、組距
                                foreach (string dName in SelDomainIdxDict.Keys)
                                {
                                    if (studSemsScoreRec.Domains.ContainsKey(dName))
                                    {
                                        // 學生5_領域2_名稱
                                        row["學生" + studIdx + "_領域" + SelDomainIdxDict[dName] + "_名稱"] = dName;
                                        if (studSemsScoreRec.Domains[dName].Credit.HasValue)
                                        {

                                            // 加入班級領域1,2,3 學分
                                            string kk = "班級_領域" + SelDomainIdxDict[dName] + "_學分";

                                            if (!tmpCreditDict.ContainsKey(kk))
                                                tmpCreditDict.Add(kk, new List<decimal>());

                                            if (!tmpCreditDict[kk].Contains(studSemsScoreRec.Domains[dName].Credit.Value))
                                                tmpCreditDict[kk].Add(studSemsScoreRec.Domains[dName].Credit.Value);


                                            row["學生" + studIdx + "_領域" + SelDomainIdxDict[dName] + "_學分"] = studSemsScoreRec.Domains[dName].Credit.Value;
                                        }


                                        if (studSemsScoreRec.Domains[dName].Score.HasValue)
                                        {
                                            row["學生" + studIdx + "_領域" + SelDomainIdxDict[dName] + "_成績"] = studSemsScoreRec.Domains[dName].Score.Value;

                                            string kka1 = "班級_領域" + SelDomainIdxDict[dName] + "_平均";
                                            if (!tmpScoreItemRDict.ContainsKey(kka1))
                                            {
                                                ScoreItemR sr = new ScoreItemR();
                                                sr.ScoreKey = kka1;
                                                sr.ScoreOriginKey = "班級_領域(原始)" + SelDomainIdxDict[dName] + "_平均";
                                                sr.PassCountKey = "班級_領域" + SelDomainIdxDict[dName] + "_及格人數";
                                                sr.PassOriginKey = "班級_領域(原始)" + SelDomainIdxDict[dName] + "_及格人數";

                                                tmpScoreItemRDict.Add(kka1, sr);
                                            }

                                            tmpScoreItemRDict[kka1].AddScore(studSemsScoreRec.Domains[dName].Score.Value);

                                            // 加入原始成績
                                            if (studSemsScoreRec.Domains[dName].ScoreOrigin.HasValue)
                                            {
                                                tmpScoreItemRDict[kka1].AddScoreOrigin(studSemsScoreRec.Domains[dName].ScoreOrigin.Value);
                                            }

                                            // 處理領域總計成績計算
                                            string keyS = "學生" + studIdx + "_領域總計成績";

                                            if (!tmpScoreItemCDict.ContainsKey(keyS))
                                            {
                                                ScoreItemC sic = new ScoreItemC();
                                                sic.ItemName = keyS;
                                                tmpScoreItemCDict.Add(keyS, sic);
                                            }

                                            tmpScoreItemCDict[keyS].AddCredit(studSemsScoreRec.Domains[dName].Credit.Value);

                                            tmpScoreItemCDict[keyS].AddScore(studSemsScoreRec.Domains[dName].Score.Value);
                                            tmpScoreItemCDict[keyS].AddScore(studSemsScoreRec.Domains[dName].Score.Value, studSemsScoreRec.Domains[dName].Credit.Value);

                                            if (studSemsScoreRec.Domains[dName].ScoreOrigin.HasValue)
                                            {
                                                tmpScoreItemCDict[keyS].AddOriginScore(studSemsScoreRec.Domains[dName].ScoreOrigin.Value);
                                                tmpScoreItemCDict[keyS].AddOriginScore(studSemsScoreRec.Domains[dName].ScoreOrigin.Value, studSemsScoreRec.Domains[dName].Credit.Value);
                                            }

                                            tmpScoreItemCDict[keyS].CalScore(studentCalculator, "domain");
                                        }

                                        if (studSemsScoreRec.Domains[dName].ScoreOrigin.HasValue)
                                        {


                                            row["學生" + studIdx + "_領域(原始)" + SelDomainIdxDict[dName] + "_成績"] = studSemsScoreRec.Domains[dName].ScoreOrigin.Value;

                                        }

                                        if (studSemsScoreRec.Domains[dName].Score.HasValue && studSemsScoreRec.Domains[dName].Score.Value < 60)
                                        {
                                            row["學生" + studIdx + "_領域" + SelDomainIdxDict[dName] + "_需補考標示"] = _Configure.NeeedReScoreMark;
                                            row["學生" + studIdx + "_領域" + SelDomainIdxDict[dName] + "_不及格標示"] = _Configure.FailScoreMark;
                                        }

                                        if (studSemsScoreRec.Domains[dName].ScoreMakeup.HasValue)
                                            row["學生" + studIdx + "_領域" + SelDomainIdxDict[dName] + "_補考成績標示"] = _Configure.ReScoreMark;


                                        foreach (string rk in rkList)
                                        {
                                            // 學期/總計成績_課程學習總成績_年排名
                                            if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/領域成績_" + dName + "_" + rk))
                                            {
                                                // 學生5_領域2_班排名_rank
                                                foreach (string r2 in r2List)
                                                {
                                                    if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/領域成績_" + dName + "_" + rk].ContainsKey(r2))
                                                    {
                                                        row["學生" + studIdx + "_領域" + SelDomainIdxDict[dName] + "_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/領域成績_" + dName + "_" + rk][r2];
                                                    }
                                                }
                                            }


                                            // 學期/總計成績_課程學習總成績_年排名  原始
                                            if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/領域成績(原始)_" + dName + "_" + rk))
                                            {
                                                // 學生5_領域2_班排名_rank
                                                foreach (string r2 in r2List)
                                                {
                                                    if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/領域成績(原始)_" + dName + "_" + rk].ContainsKey(r2))
                                                    {
                                                        row["學生" + studIdx + "_領域(原始)" + SelDomainIdxDict[dName] + "_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/領域成績(原始)_" + dName + "_" + rk][r2];
                                                    }
                                                }
                                            }

                                        }

                                    }
                                }


                                // 處理學期科目成績、排名、五標、組距
                                foreach (string sName in SelSubjectIdxDict.Keys)
                                {
                                    if (studSemsScoreRec.Subjects.ContainsKey(sName))
                                    {
                                        // 學生5_科目2_名稱
                                        row["學生" + studIdx + "_科目" + SelSubjectIdxDict[sName] + "_名稱"] = sName;

                                        if (studSemsScoreRec.Subjects[sName].Credit.HasValue)
                                        {
                                            row["學生" + studIdx + "_科目" + SelSubjectIdxDict[sName] + "_學分"] = studSemsScoreRec.Subjects[sName].Credit.Value;

                                            // 加入班級科目1,2,3 學分
                                            string kk = "班級_科目" + SelSubjectIdxDict[sName] + "_學分";
                                            if (!tmpCreditDict.ContainsKey(kk))
                                                tmpCreditDict.Add(kk, new List<decimal>());

                                            if (!tmpCreditDict[kk].Contains(studSemsScoreRec.Subjects[sName].Credit.Value))
                                                tmpCreditDict[kk].Add(studSemsScoreRec.Subjects[sName].Credit.Value);
                                        }


                                        if (studSemsScoreRec.Subjects[sName].Score.HasValue)
                                        {
                                            row["學生" + studIdx + "_科目" + SelSubjectIdxDict[sName] + "_成績"] = studSemsScoreRec.Subjects[sName].Score.Value;

                                            string kka1 = "班級_科目" + SelSubjectIdxDict[sName] + "_平均";
                                            if (!tmpScoreItemRDict.ContainsKey(kka1))
                                            {
                                                ScoreItemR sr = new ScoreItemR();
                                                sr.ScoreKey = kka1;
                                                sr.ScoreOriginKey = "班級_科目(原始)" + SelSubjectIdxDict[sName] + "_平均";
                                                sr.PassCountKey = "班級_科目" + SelSubjectIdxDict[sName] + "_及格人數";
                                                sr.PassOriginKey = "班級_科目(原始)" + SelSubjectIdxDict[sName] + "_及格人數";

                                                tmpScoreItemRDict.Add(kka1, sr);
                                            }

                                            tmpScoreItemRDict[kka1].AddScore(studSemsScoreRec.Subjects[sName].Score.Value);

                                            // 原始成績
                                            if (studSemsScoreRec.Subjects[sName].ScoreOrigin.HasValue)
                                            {
                                                tmpScoreItemRDict[kka1].AddScoreOrigin(studSemsScoreRec.Subjects[sName].ScoreOrigin.Value);
                                            }

                                            // 處理領域總計成績計算
                                            string keyS = "學生" + studIdx + "_科目總計成績";
                                            if (!tmpScoreItemCDict.ContainsKey(keyS))
                                            {
                                                ScoreItemC sic = new ScoreItemC();
                                                sic.ItemName = keyS;
                                                tmpScoreItemCDict.Add(keyS, sic);
                                            }

                                            tmpScoreItemCDict[keyS].AddCredit(studSemsScoreRec.Subjects[sName].Credit.Value);
                                            tmpScoreItemCDict[keyS].AddScore(studSemsScoreRec.Subjects[sName].Score.Value);
                                            tmpScoreItemCDict[keyS].AddScore(studSemsScoreRec.Subjects[sName].Score.Value, studSemsScoreRec.Subjects[sName].Credit.Value);

                                            if (studSemsScoreRec.Subjects[sName].ScoreOrigin.HasValue)
                                            {
                                                tmpScoreItemCDict[keyS].AddOriginScore(studSemsScoreRec.Subjects[sName].ScoreOrigin.Value);
                                                tmpScoreItemCDict[keyS].AddOriginScore(studSemsScoreRec.Subjects[sName].ScoreOrigin.Value, studSemsScoreRec.Subjects[sName].Credit.Value);
                                            }

                                            tmpScoreItemCDict[keyS].CalScore(studentCalculator, "subject");
                                        }


                                        if (studSemsScoreRec.Subjects[sName].ScoreOrigin.HasValue)
                                        {
                                            row["學生" + studIdx + "_科目(原始)" + SelSubjectIdxDict[sName] + "_成績"] = studSemsScoreRec.Subjects[sName].ScoreOrigin.Value;

                                        }


                                        if (studSemsScoreRec.Subjects[sName].ScoreMakeup.HasValue)
                                            row["學生" + studIdx + "_科目" + SelSubjectIdxDict[sName] + "_補考成績標示"] = _Configure.ReScoreMark;

                                        if (studSemsScoreRec.Subjects[sName].Score.HasValue && studSemsScoreRec.Subjects[sName].Score.Value < 60)
                                        {
                                            row["學生" + studIdx + "_科目" + SelSubjectIdxDict[sName] + "_需補考標示"] = _Configure.NeeedReScoreMark;
                                            row["學生" + studIdx + "_科目" + SelSubjectIdxDict[sName] + "_不及格標示"] = _Configure.FailScoreMark;
                                        }


                                        foreach (string rk in rkList)
                                        {
                                            // 學期/總計成績_課程學習總成績_年排名
                                            if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/科目成績_" + sName + "_" + rk))
                                            {
                                                // 學生5_科目2_班排名_rank
                                                foreach (string r2 in r2List)
                                                {
                                                    if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/科目成績_" + sName + "_" + rk].ContainsKey(r2))
                                                    {
                                                        row["學生" + studIdx + "_科目" + SelSubjectIdxDict[sName] + "_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/科目成績_" + sName + "_" + rk][r2];
                                                    }
                                                }
                                            }


                                            // 學期/科目成績_xx_年排名  原始
                                            if (SemsScoreRankMatrixDataValueDict[studRec.ID].ContainsKey("學期/科目成績(原始)_" + sName + "_" + rk))
                                            {
                                                // 學生5_科目2_班排名_rank
                                                foreach (string r2 in r2List)
                                                {
                                                    if (SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/科目成績(原始)_" + sName + "_" + rk].ContainsKey(r2))
                                                    {
                                                        row["學生" + studIdx + "_科目(原始)" + SelSubjectIdxDict[sName] + "_" + rk + "_" + r2] = SemsScoreRankMatrixDataValueDict[studRec.ID]["學期/科目成績(原始)_" + sName + "_" + rk][r2];
                                                    }
                                                }
                                            }

                                        }

                                    }
                                }

                            }

                            studIdx++;
                        }
                    }

                   
                    // 填入班級 五標、組距
                    if (ClassSemsScoreRankMatrixDataValueDict.ContainsKey(class_id))
                    {
                        // 總計成績 五標、組距
                        foreach (string rk in rkList)
                        {
                            if (ClassSemsScoreRankMatrixDataValueDict.ContainsKey("學期/總計成績_學習領域總成績_" + rk))
                            {
                                // 班級_課程學習總成績_班排名_matrix_count
                                foreach (string r2 in cr2List)
                                {
                                    if (ClassSemsScoreRankMatrixDataValueDict["學期/總計成績_學習領域總成績_" + rk].ContainsKey(r2))
                                    {
                                        row["班級_課程學習總成績_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict["學期/總計成績_學習領域總成績_" + rk][r2];
                                    }
                                }
                            }

                            if (ClassSemsScoreRankMatrixDataValueDict.ContainsKey("學期/總計成績(原始)_學習領域總成績_" + rk))
                            {
                                // 班級_課程學習總成績_班排名_matrix_count
                                foreach (string r2 in cr2List)
                                {
                                    if (ClassSemsScoreRankMatrixDataValueDict["學期/總計成績(原始)_學習領域總成績_" + rk].ContainsKey(r2))
                                    {
                                        row["班級_學習領域總成績(原始)_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict["學期/總計成績(原始)_學習領域總成績_" + rk][r2];
                                    }
                                }
                            }

                            if (ClassSemsScoreRankMatrixDataValueDict.ContainsKey("學期/總計成績_課程學習總成績_" + rk))
                            {
                                // 班級_課程學習總成績_班排名_matrix_count
                                foreach (string r2 in cr2List)
                                {
                                    if (ClassSemsScoreRankMatrixDataValueDict["學期/總計成績_課程學習總成績_" + rk].ContainsKey(r2))
                                    {
                                        row["班級_課程學習總成績_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict["學期/總計成績_課程學習總成績_" + rk][r2];
                                    }
                                }
                            }

                            if (ClassSemsScoreRankMatrixDataValueDict.ContainsKey("學期/總計成績(原始)_課程學習總成績_" + rk))
                            {
                                // 班級_課程學習總成績_班排名_matrix_count
                                foreach (string r2 in cr2List)
                                {
                                    if (ClassSemsScoreRankMatrixDataValueDict["學期/總計成績(原始)_課程學習總成績_" + rk].ContainsKey(r2))
                                    {
                                        row["班級_課程學習總成績(原始)_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict["學期/總計成績(原始)_課程學習總成績_" + rk][r2];
                                    }
                                }
                            }



                        }


                        // 班級領域-科目1,2,3  五標、組距
                        if (SelDomainSubjectDict.ContainsKey(class_id))
                        {
                            // 領域
                            foreach (string dName in SelDomainSubjectDict[class_id].Keys)
                            {
                                foreach (string rk in rkList)
                                {
                                    // 學期/總計成績_課程學習總成績_年排名
                                    if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/領域成績_" + dName + "_" + rk))
                                    {
                                        // 學生5_領域2_班排名_rank
                                        foreach (string r2 in cr2List)
                                        {
                                            if (ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績_" + dName + "_" + rk].ContainsKey(r2))
                                            {
                                                row["班級_" + dName + "領域_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績_" + dName + "_" + rk][r2];

                                            }
                                        }
                                    }

                                    // 科目
                                    int subjIdx = 1;
                                    foreach (string sName in SelDomainSubjectDict[class_id][dName])
                                    {
                                        // 學期/總計成績_課程學習總成績_年排名
                                        if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/科目成績_" + sName + "_" + rk))
                                        {
                                            row["班級_" + dName + "領域_科目" + subjIdx + "_名稱"] = sName;
                                            foreach (string r2 in cr2List)
                                            {
                                                row["班級_" + dName + "領域_科目" + subjIdx + "_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/科目成績_" + sName + "_" + rk][r2];
                                            }
                                        }
                                        subjIdx++;
                                    }


                                    // 學期/總計成績_課程學習總成績_年排名  原始
                                    if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/領域成績(原始)_" + dName + "_" + rk))
                                    {
                                        // 班級_領域2_班排名_rank
                                        foreach (string r2 in cr2List)
                                        {
                                            if (ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績(原始)_" + dName + "_" + rk].ContainsKey(r2))
                                            {
                                                row["班級_" + dName + "領域(原始)_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績(原始)_" + dName + "_" + rk][r2];
                                            }
                                        }

                                        // 科目
                                        subjIdx = 1;
                                        foreach (string sName in SelDomainSubjectDict[class_id][dName])
                                        {
                                            // 學期/總計成績_課程學習總成績_年排名
                                            if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/科目成績(原始)_" + sName + "_" + rk))
                                            {
                                                foreach (string r2 in cr2List)
                                                {
                                                    row["班級_" + dName + "領域(原始)_科目" + subjIdx + "_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/科目成績(原始)_" + sName + "_" + rk][r2];
                                                }
                                            }
                                            subjIdx++;
                                        }
                                    }
                                }
                            }
                        }

                        // 班級領域成績 五標、組距 
                        foreach (string dName in SelDomainIdxDict.Keys)
                        {
                            // 班級_領域2_名稱
                            row["班級_領域" + SelDomainIdxDict[dName] + "_名稱"] = dName;

                            foreach (string rk in rkList)
                            {
                                // 學期/總計成績_課程學習總成績_年排名
                                if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/領域成績_" + dName + "_" + rk))
                                {
                                    // 學生5_領域2_班排名_rank
                                    foreach (string r2 in r2List)
                                    {
                                        if (ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績_" + dName + "_" + rk].ContainsKey(r2))
                                        {
                                            row["班級_領域" + SelDomainIdxDict[dName] + "_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績_" + dName + "_" + rk][r2];
                                        }
                                    }
                                }


                                // 學期/總計成績_課程學習總成績_年排名  原始
                                if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/領域成績(原始)_" + dName + "_" + rk))
                                {
                                    // 班級_領域2_班排名_rank
                                    foreach (string r2 in r2List)
                                    {
                                        if (ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績(原始)_" + dName + "_" + rk].ContainsKey(r2))
                                        {
                                            row["班級_領域(原始)" + SelDomainIdxDict[dName] + "_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/領域成績(原始)_" + dName + "_" + rk][r2];
                                        }
                                    }
                                }
                            }
                        }

                        // 班級 科目成績 五標、組距                   
                        foreach (string sName in SelSubjectIdxDict.Keys)
                        {

                            // 班級_科目2_名稱
                            row["班級_科目" + SelSubjectIdxDict[sName] + "_名稱"] = sName;

                            foreach (string rk in rkList)
                            {
                                // 學期/總計成績_課程學習總成績_年排名
                                if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/科目成績_" + sName + "_" + rk))
                                {
                                    // 學生5_科目2_班排名_rank
                                    foreach (string r2 in r2List)
                                    {
                                        if (ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/科目成績_" + sName + "_" + rk].ContainsKey(r2))
                                        {
                                            row["班級_科目" + SelSubjectIdxDict[sName] + "_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/科目成績_" + sName + "_" + rk][r2];
                                        }
                                    }
                                }


                                // 學期/總計成績_課程學習總成績_年排名  原始
                                if (ClassSemsScoreRankMatrixDataValueDict[class_id].ContainsKey("學期/科目成績(原始)_" + sName + "_" + rk))
                                {
                                    // 班級_科目2_班排名_rank
                                    foreach (string r2 in r2List)
                                    {
                                        if (ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/科目成績(原始)_" + sName + "_" + rk].ContainsKey(r2))
                                        {
                                            row["班級_科目(原始)" + SelSubjectIdxDict[sName] + "_" + rk + "_" + r2] = ClassSemsScoreRankMatrixDataValueDict[class_id]["學期/科目成績(原始)_" + sName + "_" + rk][r2];
                                        }
                                    }
                                }
                            }
                        }

                    }

                    #region 填入班級學分、及格人數、平均
                    // 填入班級學分、及格人數、平均
                    foreach (string key in tmpCreditDict.Keys)
                    {
                        if (dtTable.Columns.Contains(key))
                            row[key] = string.Join(",", tmpCreditDict[key].ToArray());
                    }

                    foreach (string key in tmpScoreItemRDict.Keys)
                    {
                        ScoreItemR sr = tmpScoreItemRDict[key];
                        if (dtTable.Columns.Contains(sr.ScoreKey))
                        {
                            row[sr.ScoreKey] = sr.GetAvgScore(defaultScoreCalculator, "domain");
                        }

                        if (dtTable.Columns.Contains(sr.ScoreOriginKey))
                        {
                            row[sr.ScoreOriginKey] = sr.GetAvgScoreOrigin(defaultScoreCalculator, "domain");
                        }

                        if (dtTable.Columns.Contains(sr.PassCountKey))
                        {
                            row[sr.PassCountKey] = sr.GetPassScoreCount();
                        }

                        if (dtTable.Columns.Contains(sr.PassOriginKey))
                        {
                            row[sr.PassOriginKey] = sr.GetPassScoreOriginCount();
                        }
                    }
                    #endregion

                    #region 填入學生各類總計
                    for (int studIdx = 1; studIdx <= maxStudent; studIdx++)
                    {
                        string keyD = "學生" + studIdx + "_領域總計成績";
                        string keyS = "學生" + studIdx + "_科目總計成績";

                        if (tmpScoreItemCDict.ContainsKey(keyD))
                        {
                            row["學生" + studIdx + "_領域總計_加權總分"] = tmpScoreItemCDict[keyD].SumScoreA;
                            row["學生" + studIdx + "_領域總計_總分"] = tmpScoreItemCDict[keyD].SumScore;
                            row["學生" + studIdx + "_領域總計_加權平均"] = tmpScoreItemCDict[keyD].AvgScoreA;
                            row["學生" + studIdx + "_領域總計_平均"] = tmpScoreItemCDict[keyD].AvgScore;
                            row["學生" + studIdx + "_領域(原始)總計_加權總分"] = tmpScoreItemCDict[keyD].SumScoreOriginA;
                            row["學生" + studIdx + "_領域(原始)總計_總分"] = tmpScoreItemCDict[keyD].SumScoreOrigin;
                            row["學生" + studIdx + "_領域(原始)總計_加權平均"] = tmpScoreItemCDict[keyD].AvgScoreOriginA;
                            row["學生" + studIdx + "_領域(原始)總計_平均"] = tmpScoreItemCDict[keyD].AvgScoreOrigin;
                        }

                        if (tmpScoreItemCDict.ContainsKey(keyS))
                        {
                            row["學生" + studIdx + "_科目總計_加權總分"] = tmpScoreItemCDict[keyS].SumScoreA;
                            row["學生" + studIdx + "_科目總計_總分"] = tmpScoreItemCDict[keyS].SumScore;
                            row["學生" + studIdx + "_科目總計_加權平均"] = tmpScoreItemCDict[keyS].AvgScoreA;
                            row["學生" + studIdx + "_科目總計_平均"] = tmpScoreItemCDict[keyS].AvgScore;
                            row["學生" + studIdx + "_科目(原始)總計_加權總分"] = tmpScoreItemCDict[keyS].SumScoreOriginA;
                            row["學生" + studIdx + "_科目(原始)總計_總分"] = tmpScoreItemCDict[keyS].SumScoreOrigin;
                            row["學生" + studIdx + "_科目(原始)總計_加權平均"] = tmpScoreItemCDict[keyS].AvgScoreOriginA;
                            row["學生" + studIdx + "_科目(原始)總計_平均"] = tmpScoreItemCDict[keyS].AvgScoreOrigin;
                        }
                    }

                    #endregion


                    dtTable.Rows.Add(row);
                }

                bgWorkerReport.ReportProgress(80);

                #endregion



                // debug
                //dtTable.TableName = "debug";
                //dtTable.WriteXml(Application.StartupPath + "\\debug.xml");

                //return;


                Document doc = _Configure.Template;
                doc.MailMerge.Execute(dtTable);
                doc.MailMerge.DeleteFields();
                e.Result = doc;
                bgWorkerReport.ReportProgress(100);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
          
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

            userControlEnable(true);
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
            ClassStudentDict.Clear();
            ClassRecordDict.Clear();
            List<JHStudentRecord> studRecList = JHStudent.SelectByClassIDs(_ClassIDList);
            List<JHClassRecord> classRecList = JHClass.SelectByIDs(_ClassIDList);
            foreach (JHClassRecord rec in classRecList)
            {
                if (!ClassRecordDict.ContainsKey(rec.ID))
                    ClassRecordDict.Add(rec.ID, rec);
            }

            foreach (JHStudentRecord rec in studRecList)
            {
                if (rec.Status == StudentRecord.StudentStatus.一般)
                {
                    if (!StudentClassIDDict.ContainsKey(rec.ID))
                        StudentClassIDDict.Add(rec.ID, rec.RefClassID);

                    if (!ClassStudentDict.ContainsKey(rec.RefClassID))
                        ClassStudentDict.Add(rec.RefClassID, new List<StudentRecord>());

                    ClassStudentDict[rec.RefClassID].Add(rec);

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
            ClassDomainSubjectNameDict.Clear();


            CanSelectDomainSubjectDict.Add("彈性課程", new List<string>());

            foreach (JHSemesterScoreRecord rec in SemesterScoreRecordLList)
            {
                if (!StudentSemesterScoreRecordDict.ContainsKey(rec.RefStudentID))
                    StudentSemesterScoreRecordDict.Add(rec.RefStudentID, rec);

                if (StudentClassIDDict.ContainsKey(rec.RefStudentID))
                {
                    if (!ClassDomainNameDict.ContainsKey(StudentClassIDDict[rec.RefStudentID]))
                        ClassDomainNameDict.Add(StudentClassIDDict[rec.RefStudentID], new List<string>());

                    if (!ClassDomainSubjectNameDict.ContainsKey(StudentClassIDDict[rec.RefStudentID]))
                        ClassDomainSubjectNameDict.Add(StudentClassIDDict[rec.RefStudentID], new Dictionary<string, List<string>>());


                    foreach (string key in rec.Domains.Keys)
                    {
                        if (!ClassDomainNameDict[StudentClassIDDict[rec.RefStudentID]].Contains(key))
                            ClassDomainNameDict[StudentClassIDDict[rec.RefStudentID]].Add(key);

                        if (!CanSelectDomainSubjectDict.ContainsKey(key))
                            CanSelectDomainSubjectDict.Add(key, new List<string>());


                        if (!ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].ContainsKey(key))
                            ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].Add(key, new List<string>());
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

                            if (!ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].ContainsKey("彈性課程"))
                                ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].Add("彈性課程", new List<string>());

                            if (!ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]]["彈性課程"].Contains(ss.Subject))
                                ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]]["彈性課程"].Add(ss.Subject);

                        }
                        else
                        {
                            if (CanSelectDomainSubjectDict.ContainsKey(ss.Domain))
                            {
                                if (!CanSelectDomainSubjectDict[ss.Domain].Contains(ss.Subject))
                                    CanSelectDomainSubjectDict[ss.Domain].Add(ss.Subject);
                            }

                            if (!ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].ContainsKey(ss.Domain))
                                ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]].Add(ss.Domain, new List<string>());

                            if (!ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]][ss.Domain].Contains(ss.Subject))
                                ClassDomainSubjectNameDict[StudentClassIDDict[rec.RefStudentID]][ss.Domain].Add(ss.Subject);
                        }

                    }
                }

            }

        }



        private void btnPrint_Click(object sender, EventArgs e)
        {
            //儲存設定檔
            SaveConfig();


            // 取得畫面上勾選領域科目
            SelDomainSubjectDict.Clear();
            SelDomainIdxDict.Clear();
            SelSubjectIdxDict.Clear();

            // 各班畫面上勾選領域科目
            foreach (string cid in _ClassIDList)
            {
                SelDomainSubjectDict.Add(cid, new Dictionary<string, List<string>>());
                if (ClassDomainSubjectNameDict.ContainsKey(cid))
                {
                    foreach (ListViewItem lvi in lvDomain.CheckedItems)
                    {
                        if (ClassDomainSubjectNameDict[cid].ContainsKey(lvi.Name))
                        {
                            SelDomainSubjectDict[cid].Add(lvi.Name, new List<string>());
                        }
                    }

                    foreach (ListViewItem lvi in lvSubject.CheckedItems)
                    {
                        string key = lvi.Name;

                        // 檢查自己領域是否有加入
                        string keyD = lvi.Tag.ToString();

                        if (!SelDomainSubjectDict[cid].ContainsKey(keyD))
                            SelDomainSubjectDict[cid].Add(keyD, new List<string>());

                        if (ClassDomainSubjectNameDict[cid].ContainsKey(keyD))
                        {
                            if (ClassDomainSubjectNameDict[cid][keyD].Contains(lvi.Text))
                                SelDomainSubjectDict[cid][keyD].Add(lvi.Text);
                        }
                    }

                }
            }

            int sd = 1;
            List<string> tmpDomainNameList = new List<string>();
            // 領域
            foreach (ListViewItem lvi in lvDomain.CheckedItems)
            {
                string key = lvi.Name;

                if (!tmpDomainNameList.Contains(key))
                    tmpDomainNameList.Add(key);

            }

            tmpDomainNameList.Sort(new StringComparer("語文", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"));

            foreach (string name in tmpDomainNameList)
            {
                SelDomainIdxDict.Add(name, sd);
                sd++;
            }

            int ss = 1;
            List<string> tmpSubjNameList = new List<string>();
            foreach (ListViewItem lvi in lvSubject.CheckedItems)
            {
                string key = lvi.Name;

                // 檢查自己領域是否有加入
                string keyD = lvi.Tag.ToString();

                if (!tmpSubjNameList.Contains(lvi.Text))
                    tmpSubjNameList.Add(lvi.Text);
            }

            tmpSubjNameList.Sort(new StringComparer("國文", "英文", "數學", "理化", "生物", "社會", "物理", "化學", "歷史", "地理", "公民"));


            foreach (string name in tmpSubjNameList)
            {
                SelSubjectIdxDict.Add(name, ss);
                ss++;
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
            Global.domainNameList.Clear();
            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學期成績合併欄位總表產生中...");
            foreach (ListViewItem lvi in lvDomain.Items)
            {
                if (lvi.Checked)
                    Global.domainNameList.Add(lvi.Text);
            }

            // 產生合併欄位總表
            lnkViewMapColumns.Enabled = false;

            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            cboSemester.Text = K12.Data.School.DefaultSemester;

            userControlEnable(false);
            this.MaximumSize = this.MinimumSize = this.Size;
            bgWorkerLoadTemplate.RunWorkerAsync();
        }

        private void userControlEnable(bool value)
        {
            cboSchoolYear.Enabled = value;
            cboSemester.Enabled = value;
            txtNeeedReScoreMark.Enabled = value;
            txtReScoreMark.Enabled = value;
            txtFailScoreMark.Enabled = value;
            lnkViewTemplate.Enabled = value;
            lnkChangeTemplate.Enabled = value;
            lnkViewMapColumns.Enabled = value;
            btnSaveConfig.Enabled = value;
            btnPrint.Enabled = value;


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
