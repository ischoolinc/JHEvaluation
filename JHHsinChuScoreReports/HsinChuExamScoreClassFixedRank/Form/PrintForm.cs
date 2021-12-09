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
using System.Xml;

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
        string RefSelSchoolYear = "";
        string RefSelSemester = "";
        string SelExamID = "";
        string SelExamName = "";
        string SelRefExamID = "";
        string SelRefExamName = "";
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

                string reportName = "" + SelSchoolYear + "學年度第" + SelSemester + "學期" + SelExamName + "班級評量成績單(固定排名)";

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
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級評量成績單產生完成");
            else
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級評量成績單產生中...", e.ProgressPercentage);
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

            // 取得取得學生評量固定排名(參考試別)
            DataTable dtStudentRefExamRankMatrix = DataAccess.GetStudentExamRankMatrix(RefSelSchoolYear, RefSelSemester, SelRefExamID, _ClassIDList);

            // 取得固定排名 班級、年級 五標與組距
            DataTable dtClassExamRankMatrix = DataAccess.GetClassExamRankMatrix(SelSchoolYear, SelSemester, SelExamID, _ClassIDList);

            // debug
            //dtStudentExamRankMatrix.TableName = "dtStudentExamRankMatrix";
            //dtStudentExamRankMatrix.WriteXml(Application.StartupPath + "\\tmp.xml");

            Dictionary<string, Dictionary<string, int>> StudentExamRankMatrixDict = new Dictionary<string, Dictionary<string, int>>();


            // 班級五標與組距
            Dictionary<string, Dictionary<string, DataRow>> ClassExamRankMatrixDict = new Dictionary<string, Dictionary<string, DataRow>>();


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


            // 學生評量參考試別 Dict		
            Dictionary<string, Dictionary<string, int>> StudentRefExamRankMatrixDict = new Dictionary<string, Dictionary<string, int>>();
            // 建立排名索引
            if (dtStudentRefExamRankMatrix != null && dtStudentRefExamRankMatrix.Rows.Count > 0)
            {

                foreach (DataRow dr in dtStudentRefExamRankMatrix.Rows)
                {
                    string student_id = dr["ref_student_id"].ToString();
                    // 定期評量/總計成績 加權總分 班排名
                    string type = dr["item_type"].ToString() + dr["item_name"].ToString() + dr["rank_type"].ToString();
                    if (!StudentRefExamRankMatrixDict.ContainsKey(student_id))
                        StudentRefExamRankMatrixDict.Add(student_id, new Dictionary<string, int>());

                    if (!StudentRefExamRankMatrixDict[student_id].ContainsKey(type))
                    {
                        int rank = 0;
                        int.TryParse(dr["rank"].ToString(), out rank);
                        StudentRefExamRankMatrixDict[student_id].Add(type, rank);
                    }
                }
            }


            // 建立索引
            if (dtClassExamRankMatrix != null && dtClassExamRankMatrix.Rows.Count > 0)
            {
                foreach (DataRow dr in dtClassExamRankMatrix.Rows)
                {
                    string key = dr["rank_name"].ToString();// 班級 or 年排名
                    string type = dr["item_type"].ToString() + dr["item_name"].ToString() + dr["rank_type"].ToString();
                    if (!ClassExamRankMatrixDict.ContainsKey(key))
                        ClassExamRankMatrixDict.Add(key, new Dictionary<string, DataRow>());

                    if (!ClassExamRankMatrixDict[key].ContainsKey(type))
                        ClassExamRankMatrixDict[key].Add(type, dr);
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

                        #region 學生總計班排名
                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分班排名"))
                            si.ClassSumRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分班排名"))
                            si.ClassSumRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均班排名"))
                            si.ClassAvgRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均班排名"))
                            si.ClassAvgRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分班排名"))
                            si.ClassSumRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分班排名"))
                            si.ClassSumRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均班排名"))
                            si.ClassAvgRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均班排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均班排名"))
                            si.ClassAvgRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均班排名"];

                        #endregion

                        #region 學生總計年排名
                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分年排名"))
                            si.YearSumRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分年排名"))
                            si.YearSumRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均年排名"))
                            si.YearAvgRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均年排名"))
                            si.YearAvgRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均年排名"];


                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分年排名"))
                            si.YearSumRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分年排名"))
                            si.YearSumRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均年排名"))
                            si.YearAvgRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均年排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均年排名"))
                            si.YearAvgRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均年排名"];
                        #endregion

                        #region 學生總計類別1排名
                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分類別1排名"))
                            si.ClassType1SumRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分類別1排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分類別1排名"))
                            si.ClassType1SumRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分類別1排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均類別1排名"))
                            si.ClassType1AvgRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均類別1排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均類別1排名"))
                            si.ClassType1AvgRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均類別1排名"];


                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分類別1排名"))
                            si.ClassType1SumRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分類別1排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分類別1排名"))
                            si.ClassType1SumRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分類別1排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均類別1排名"))
                            si.ClassType1AvgRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均類別1排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均類別1排名"))
                            si.ClassType1AvgRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均類別1排名"];
                        #endregion

                        #region 學生總計類別2排名
                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分類別2排名"))
                            si.ClassType2SumRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分類別2排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分類別2排名"))
                            si.ClassType2SumRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分類別2排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均類別2排名"))
                            si.ClassType2AvgRankA = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均類別2排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均類別2排名"))
                            si.ClassType2AvgRank = StudentExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均類別2排名"];


                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分類別2排名"))
                            si.ClassType2SumRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分類別2排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分類別2排名"))
                            si.ClassType2SumRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分類別2排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均類別2排名"))
                            si.ClassType2AvgRankAF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均類別2排名"];

                        if (StudentExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均類別2排名"))
                            si.ClassType2AvgRankF = StudentExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均類別2排名"];
                        #endregion


                    }

                    if (StudentRefExamRankMatrixDict.ContainsKey(si.StudentID))
                    {
                        #region 學生總計班排名 (參考試別)
                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分班排名"))
                            si.ClassSumRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分班排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分班排名"))
                            si.ClassSumRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分班排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均班排名"))
                            si.ClassAvgRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均班排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均班排名"))
                            si.ClassAvgRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均班排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分班排名"))
                            si.ClassSumRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分班排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分班排名"))
                            si.ClassSumRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分班排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均班排名"))
                            si.ClassAvgRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均班排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均班排名"))
                            si.ClassAvgRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均班排名"];

                        #endregion

                        #region 學生總計年排名(參考試別)
                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分年排名"))
                            si.YearSumRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分年排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分年排名"))
                            si.YearSumRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分年排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均年排名"))
                            si.YearAvgRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均年排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均年排名"))
                            si.YearAvgRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均年排名"];


                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分年排名"))
                            si.YearSumRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分年排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分年排名"))
                            si.YearSumRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分年排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均年排名"))
                            si.YearAvgRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均年排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均年排名"))
                            si.YearAvgRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均年排名"];
                        #endregion

                        #region 學生總計類別1排名 (參考試別)
                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分類別1排名"))
                            si.ClassType1SumRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分類別1排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分類別1排名"))
                            si.ClassType1SumRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分類別1排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均類別1排名"))
                            si.ClassType1AvgRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均類別1排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均類別1排名"))
                            si.ClassType1AvgRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均類別1排名"];


                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分類別1排名"))
                            si.ClassType1SumRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分類別1排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分類別1排名"))
                            si.ClassType1SumRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分類別1排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均類別1排名"))
                            si.ClassType1AvgRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均類別1排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均類別1排名"))
                            si.ClassType1AvgRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均類別1排名"];
                        #endregion

                        #region 學生總計類別2排名  (參考試別)
                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權總分類別2排名"))
                            si.ClassType2SumRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權總分類別2排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績總分類別2排名"))
                            si.ClassType2SumRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績總分類別2排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績加權平均類別2排名"))
                            si.ClassType2AvgRefRankA = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績加權平均類別2排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量/總計成績平均類別2排名"))
                            si.ClassType2AvgRefRank = StudentRefExamRankMatrixDict[si.StudentID]["定期評量/總計成績平均類別2排名"];


                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權總分類別2排名"))
                            si.ClassType2SumRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權總分類別2排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績總分類別2排名"))
                            si.ClassType2SumRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績總分類別2排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績加權平均類別2排名"))
                            si.ClassType2AvgRefRankAF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績加權平均類別2排名"];

                        if (StudentRefExamRankMatrixDict[si.StudentID].ContainsKey("定期評量_定期/總計成績平均類別2排名"))
                            si.ClassType2AvgRefRankF = StudentRefExamRankMatrixDict[si.StudentID]["定期評量_定期/總計成績平均類別2排名"];
                        #endregion
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
            dtTable.Columns.Add("參考試別");
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

            // 學生單領域與科目成績欄位
            for (int ss = 1; ss <= 30; ss++)
            {
                dtTable.Columns.Add("學生_科目名稱" + ss);
                dtTable.Columns.Add("學生_科目學分" + ss);
                dtTable.Columns.Add("學生_領域名稱" + ss);
                dtTable.Columns.Add("學生_領域學分" + ss);
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
                dtTable.Columns.Add("總分_定期" + studCot);
                dtTable.Columns.Add("加權總分_定期" + studCot);
                dtTable.Columns.Add("平均_定期" + studCot);
                dtTable.Columns.Add("加權平均_定期" + studCot);

                dtTable.Columns.Add("總分班排名" + studCot);
                dtTable.Columns.Add("加權總分班排名" + studCot);
                dtTable.Columns.Add("平均班排名" + studCot);
                dtTable.Columns.Add("加權平均班排名" + studCot);
                dtTable.Columns.Add("總分年排名" + studCot);
                dtTable.Columns.Add("加權總分年排名" + studCot);
                dtTable.Columns.Add("平均年排名" + studCot);
                dtTable.Columns.Add("加權平均年排名" + studCot);

                dtTable.Columns.Add("總分_定期班排名" + studCot);
                dtTable.Columns.Add("加權總分_定期班排名" + studCot);
                dtTable.Columns.Add("平均_定期班排名" + studCot);
                dtTable.Columns.Add("加權平均_定期班排名" + studCot);
                dtTable.Columns.Add("總分_定期年排名" + studCot);
                dtTable.Columns.Add("加權總分_定期年排名" + studCot);
                dtTable.Columns.Add("平均_定期年排名" + studCot);
                dtTable.Columns.Add("加權平均_定期年排名" + studCot);

                dtTable.Columns.Add("總分類別1排名" + studCot);
                dtTable.Columns.Add("加權總分類別1排名" + studCot);
                dtTable.Columns.Add("平均類別1排名" + studCot);
                dtTable.Columns.Add("加權平均類別1排名" + studCot);
                dtTable.Columns.Add("總分_定期類別1排名" + studCot);
                dtTable.Columns.Add("加權總分_定期類別1排名" + studCot);
                dtTable.Columns.Add("平均_定期類別1排名" + studCot);
                dtTable.Columns.Add("加權平均_定期類別1排名" + studCot);

                dtTable.Columns.Add("總分類別2排名" + studCot);
                dtTable.Columns.Add("加權總分類別2排名" + studCot);
                dtTable.Columns.Add("平均類別2排名" + studCot);
                dtTable.Columns.Add("加權平均類別2排名" + studCot);
                dtTable.Columns.Add("總分_定期類別2排名" + studCot);
                dtTable.Columns.Add("加權總分_定期類別2排名" + studCot);
                dtTable.Columns.Add("平均_定期類別2排名" + studCot);
                dtTable.Columns.Add("加權平均_定期類別2排名" + studCot);

                dtTable.Columns.Add("總分班排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分班排名_參考試別" + studCot);
                dtTable.Columns.Add("平均班排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均班排名_參考試別" + studCot);
                dtTable.Columns.Add("總分年排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分年排名_參考試別" + studCot);
                dtTable.Columns.Add("平均年排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均年排名_參考試別" + studCot);
                dtTable.Columns.Add("總分_定期班排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分_定期班排名_參考試別" + studCot);
                dtTable.Columns.Add("平均_定期班排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均_定期班排名_參考試別" + studCot);
                dtTable.Columns.Add("總分_定期年排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分_定期年排名_參考試別" + studCot);
                dtTable.Columns.Add("平均_定期年排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均_定期年排名_參考試別" + studCot);
                dtTable.Columns.Add("總分類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("平均類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("總分_定期類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分_定期類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("平均_定期類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均_定期類別1排名_參考試別" + studCot);
                dtTable.Columns.Add("總分類別2排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分類別2排名_參考試別" + studCot);
                dtTable.Columns.Add("平均類別2排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均類別2排名_參考試別" + studCot);
                dtTable.Columns.Add("總分_定期類別2排名_參考試別" + studCot);
                dtTable.Columns.Add("加權總分_定期類別2排名_參考試別" + studCot);
                dtTable.Columns.Add("平均_定期類別2排名_參考試別" + studCot);
                dtTable.Columns.Add("加權平均_定期類別2排名_參考試別" + studCot);

                dtTable.Columns.Add("總分班排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分班排名_進退步" + studCot);
                dtTable.Columns.Add("平均班排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均班排名_進退步" + studCot);
                dtTable.Columns.Add("總分年排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分年排名_進退步" + studCot);
                dtTable.Columns.Add("平均年排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均年排名_進退步" + studCot);
                dtTable.Columns.Add("總分_定期班排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分_定期班排名_進退步" + studCot);
                dtTable.Columns.Add("平均_定期班排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均_定期班排名_進退步" + studCot);
                dtTable.Columns.Add("總分_定期年排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分_定期年排名_進退步" + studCot);
                dtTable.Columns.Add("平均_定期年排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均_定期年排名_進退步" + studCot);
                dtTable.Columns.Add("總分類別1排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分類別1排名_進退步" + studCot);
                dtTable.Columns.Add("平均類別1排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均類別1排名_進退步" + studCot);
                dtTable.Columns.Add("總分_定期類別1排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分_定期類別1排名_進退步" + studCot);
                dtTable.Columns.Add("平均_定期類別1排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均_定期類別1排名_進退步" + studCot);
                dtTable.Columns.Add("總分類別2排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分類別2排名_進退步" + studCot);
                dtTable.Columns.Add("平均類別2排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均類別2排名_進退步" + studCot);
                dtTable.Columns.Add("總分_定期類別2排名_進退步" + studCot);
                dtTable.Columns.Add("加權總分_定期類別2排名_進退步" + studCot);
                dtTable.Columns.Add("平均_定期類別2排名_進退步" + studCot);
                dtTable.Columns.Add("加權平均_定期類別2排名_進退步" + studCot);


                foreach (string dName in Global.DomainNameList)
                {
                    if (SelDomainSubjectDict.ContainsKey(dName))
                    {
                        dtTable.Columns.Add(dName + "領域成績" + studCot);
                        dtTable.Columns.Add(dName + "領域_定期成績" + studCot);
                        dtTable.Columns.Add(dName + "領域學分" + studCot);

                        for (int i = 1; i <= 12; i++)
                        {
                            dtTable.Columns.Add(dName + "領域_科目名稱" + studCot + "_" + i);
                            dtTable.Columns.Add(dName + "領域_科目成績" + studCot + "_" + i);
                            dtTable.Columns.Add(dName + "領域_科目_定期成績" + studCot + "_" + i);
                            dtTable.Columns.Add(dName + "領域_科目_平時成績" + studCot + "_" + i);
                            dtTable.Columns.Add(dName + "領域_科目學分" + studCot + "_" + i);

                        }
                    }
                }

                // 學生單領域與科目成績欄位
                for (int ss = 1; ss <= 30; ss++)
                {
                    dtTable.Columns.Add("學生_科目名稱" + studCot + "_" + ss);
                    dtTable.Columns.Add("學生_科目成績" + studCot + "_" + ss);
                    dtTable.Columns.Add("學生_科目_定期成績" + studCot + "_" + ss);
                    dtTable.Columns.Add("學生_科目_平時成績" + studCot + "_" + ss);

                    dtTable.Columns.Add("學生_科目學分" + studCot + "_" + ss);

                    dtTable.Columns.Add("學生_領域名稱" + studCot + "_" + ss);
                    dtTable.Columns.Add("學生_領域成績" + studCot + "_" + ss);
                    dtTable.Columns.Add("學生_領域_定期成績" + studCot + "_" + ss);
                    dtTable.Columns.Add("學生_領域學分" + studCot + "_" + ss);
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

            List<string> rankTypeList = new List<string>();
            rankTypeList.Add("matrix_count");
            rankTypeList.Add("level_gte100");
            rankTypeList.Add("level_90");
            rankTypeList.Add("level_80");
            rankTypeList.Add("level_70");
            rankTypeList.Add("level_60");
            rankTypeList.Add("level_50");
            rankTypeList.Add("level_40");
            rankTypeList.Add("level_30");
            rankTypeList.Add("level_20");
            rankTypeList.Add("level_10");
            rankTypeList.Add("level_lt10");
            rankTypeList.Add("avg_top_25");
            rankTypeList.Add("avg_top_50");
            rankTypeList.Add("avg");
            rankTypeList.Add("avg_bottom_50");
            rankTypeList.Add("avg_bottom_25");

            // 建立班級、年組距合併欄位 領域
            List<int> grList = new List<int>();
            // 班級
            int clCount = 1;
            foreach (ClassInfo ci in ClassInfoList)
            {
                if (!grList.Contains(ci.GradeYear))
                    grList.Add(ci.GradeYear);

                dtTable.Columns.Add("班級" + clCount + "_名稱");

                foreach (string dName in SelDomainSubjectDict.Keys)
                {
                    foreach (string rt in rankTypeList)
                    {
                        dtTable.Columns.Add("班級" + clCount + "_" + dName + "_領域成績_" + rt);
                        dtTable.Columns.Add("班級" + clCount + "_" + dName + "_領域定期成績_" + rt);
                    }
                }

                clCount++;
            }

            // 年級
            int grCount = 1;
            foreach (int gr in grList)
            {
                dtTable.Columns.Add("年級" + grCount + "_名稱");
                foreach (string dName in SelDomainSubjectDict.Keys)
                {
                    foreach (string rt in rankTypeList)
                    {
                        dtTable.Columns.Add("年級" + grCount + "_" + dName + "_領域成績_" + rt);
                        dtTable.Columns.Add("年級" + grCount + "_" + dName + "_領域定期成績_" + rt);
                    }
                }
                grCount++;
            }

            // 建立班級、年組距合併欄位 科目
            // 班級
            int clsCount = 1;
            // 科目數
            int selSubjCount = 0;
            foreach (string dName in SelDomainSubjectDict.Keys)
            {
                selSubjCount += SelDomainSubjectDict[dName].Count;
            }

            foreach (ClassInfo ci in ClassInfoList)
            {
                if (!grList.Contains(ci.GradeYear))
                    grList.Add(ci.GradeYear);

                for (int su = 1; su <= selSubjCount; su++)
                {
                    dtTable.Columns.Add("班級" + clsCount + "_科目名稱_" + su);
                    foreach (string rt in rankTypeList)
                    {
                        dtTable.Columns.Add("班級" + clsCount + "_科目成績_" + rt + su);
                        dtTable.Columns.Add("班級" + clsCount + "_科目定期成績_" + rt + su);
                    }
                }


                clsCount++;
            }

            // 年級
            int grsCount = 1;
            foreach (int gr in grList)
            {
                for (int su = 1; su <= selSubjCount; su++)
                {
                    dtTable.Columns.Add("年級" + grsCount + "_科目名稱_" + su);
                    foreach (string rt in rankTypeList)
                    {
                        dtTable.Columns.Add("年級" + grsCount + "_科目成績_" + rt + su);
                        dtTable.Columns.Add("年級" + grsCount + "_科目定期成績_" + rt + su);
                    }
                }

                grsCount++;
            }

            //using (StreamWriter fi = new StreamWriter(Application.StartupPath + "\\test.txt", false))
            //{
            //    foreach (DataColumn dc in dtTable.Columns)
            //    {
            //        fi.WriteLine(dc.Caption);
            //    }
            //}


            bgWorkerReport.ReportProgress(30);


            // 單存科目或領域 學分數使用
            Dictionary<string, List<decimal>> tmpSubjectCreditDict = new Dictionary<string, List<decimal>>();
            Dictionary<string, List<decimal>> tmpDomainCreditDict = new Dictionary<string, List<decimal>>();

            Dictionary<int, List<string>> grClassNameDict = new Dictionary<int, List<string>>();

            foreach (ClassInfo ci in ClassInfoList)
            {
                if (!grClassNameDict.ContainsKey(ci.GradeYear))
                    grClassNameDict.Add(ci.GradeYear, new List<string>());

                grClassNameDict[ci.GradeYear].Add(ci.ClassName);
            }

            // 所選只有科目
            List<string> SelDomainNameList = new List<string>();
            List<string> SelSubjectNameList = new List<string>();

            foreach (string dName in SelDomainSubjectDict.Keys)
            {
                if (!SelDomainNameList.Contains(dName))
                    SelDomainNameList.Add(dName);

                foreach (string sName in SelDomainSubjectDict[dName])
                {
                    if (!SelSubjectNameList.Contains(sName))
                        SelSubjectNameList.Add(sName);
                }
            }

            List<string> ddList = DataAccess.GetDomainConfigSortName();

            SelDomainNameList.Sort(new StringComparer(ddList.ToArray()));

            //SelDomainNameList.Sort(new StringComparer("語文", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"));
            
            //科目排序
            SelSubjectNameList.Sort(new StringComparer(Utility.GetSubjectOrder().ToArray()));

            // 所選單一排序
            //SelSubjectNameList.Sort(new StringComparer("國文"
            //                    , "英文"
            //                    , "數學"
            //                    , "理化"
            //                    , "生物"
            //                    , "社會"
            //                    , "物理"
            //                    , "化學"
            //                    , "歷史"
            //                    , "地理"
            //                    , "公民"));


            //foreach (string dName in SelDomainSubjectDict.Keys)
            //{
            //    SelDomainSubjectDict[dName].Sort(new StringComparer("國文"
            //                    , "英文"
            //                    , "數學"
            //                    , "理化"
            //                    , "生物"
            //                    , "社會"
            //                    , "物理"
            //                    , "化學"
            //                    , "歷史"
            //                    , "地理"
            //                    , "公民"));
            //}
            foreach (string dName in SelDomainSubjectDict.Keys)
            {
                SelDomainSubjectDict[dName].Sort(new StringComparer(Utility.GetSubjectOrder().ToArray()));
            }


            // 填入資料
            foreach (ClassInfo ci in ClassInfoList)
            {
                DataRow row = dtTable.NewRow();
                row["學校名稱"] = SchoolName;
                row["學年度"] = SelSchoolYear;
                row["學期"] = SelSemester;
                row["試別名稱"] = SelExamName;
                row["參考試別"] = SelRefExamName;
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
                                        ss++;
                                    }
                                }
                            }

                        }
                    }

                }

                // 成績索引
                Dictionary<string, DomainInfo> tmpDomainInfoDict = new Dictionary<string, DomainInfo>();
                Dictionary<string, SubjectInfo> tmpSubjectInfoDict = new Dictionary<string, SubjectInfo>();

                tmpDomainCreditDict.Clear();
                tmpSubjectCreditDict.Clear();

                // 自己班上有的科目名稱
                List<string> hasSubjectNameList = new List<string>();

                foreach (StudentInfo si in ci.Students)
                {
                    foreach (DomainInfo di in si.DomainInfoList)
                    {
                        foreach (SubjectInfo subj in di.SubjectInfoList)
                        {
                            if (!hasSubjectNameList.Contains(subj.Name))
                                hasSubjectNameList.Add(subj.Name);
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

                    if (si.SumScoreF.HasValue)
                        row["總分_定期" + studCot] = Math.Round(si.SumScoreF.Value, parseNumber, MidpointRounding.AwayFromZero);

                    if (si.SumScoreAF.HasValue)
                        row["加權總分_定期" + studCot] = Math.Round(si.SumScoreAF.Value, parseNumber, MidpointRounding.AwayFromZero);

                    if (si.AvgScoreF.HasValue)
                        row["平均_定期" + studCot] = Math.Round(si.AvgScoreF.Value, parseNumber, MidpointRounding.AwayFromZero);

                    if (si.AvgScoreAF.HasValue)
                        row["加權平均_定期" + studCot] = Math.Round(si.AvgScoreAF.Value, parseNumber, MidpointRounding.AwayFromZero);


                    #region 評量成績排名
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

                    if (si.ClassSumRankF.HasValue)
                        row["總分_定期班排名" + studCot] = si.ClassSumRankF.Value;

                    if (si.ClassSumRankAF.HasValue)
                        row["加權總分_定期班排名" + studCot] = si.ClassSumRankAF.Value;

                    if (si.ClassAvgRankF.HasValue)
                        row["平均_定期班排名" + studCot] = si.ClassAvgRankF.Value;

                    if (si.ClassAvgRankAF.HasValue)
                        row["加權平均_定期班排名" + studCot] = si.ClassAvgRankAF.Value;

                    if (si.YearSumRankF.HasValue)
                        row["總分_定期年排名" + studCot] = si.YearSumRankF.Value;

                    if (si.YearSumRankAF.HasValue)
                        row["加權總分_定期年排名" + studCot] = si.YearSumRankAF.Value;

                    if (si.YearAvgRankF.HasValue)
                        row["平均_定期年排名" + studCot] = si.YearAvgRankF.Value;

                    if (si.YearAvgRankAF.HasValue)
                        row["加權平均_定期年排名" + studCot] = si.YearAvgRankAF.Value;


                    if (si.ClassType1SumRank.HasValue)
                        row["總分類別1排名" + studCot] = si.ClassType1SumRank.Value;

                    if (si.ClassType1SumRankA.HasValue)
                        row["加權總分類別1排名" + studCot] = si.ClassType1SumRankA.Value;

                    if (si.ClassType1AvgRank.HasValue)
                        row["平均類別1排名" + studCot] = si.ClassType1AvgRank.Value;

                    if (si.ClassType1AvgRankA.HasValue)
                        row["加權平均類別1排名" + studCot] = si.ClassType1AvgRankA.Value;

                    if (si.ClassType1SumRankF.HasValue)
                        row["總分_定期類別1排名" + studCot] = si.ClassType1SumRankF.Value;

                    if (si.ClassType1SumRankAF.HasValue)
                        row["加權總分_定期類別1排名" + studCot] = si.ClassType1SumRankAF.Value;

                    if (si.ClassType1AvgRankF.HasValue)
                        row["平均_定期類別1排名" + studCot] = si.ClassType1AvgRankF.Value;

                    if (si.ClassType1AvgRankAF.HasValue)
                        row["加權平均_定期類別1排名" + studCot] = si.ClassType1AvgRankAF.Value;

                    if (si.ClassType2SumRank.HasValue)
                        row["總分類別2排名" + studCot] = si.ClassType2SumRank.Value;

                    if (si.ClassType2SumRankA.HasValue)
                        row["加權總分類別2排名" + studCot] = si.ClassType2SumRankA.Value;

                    if (si.ClassType2AvgRank.HasValue)
                        row["平均類別2排名" + studCot] = si.ClassType2AvgRank.Value;

                    if (si.ClassType2AvgRankA.HasValue)
                        row["加權平均類別2排名" + studCot] = si.ClassType2AvgRankA.Value;

                    if (si.ClassType2SumRankF.HasValue)
                        row["總分_定期類別2排名" + studCot] = si.ClassType2SumRankF.Value;

                    if (si.ClassType2SumRankAF.HasValue)
                        row["加權總分_定期類別2排名" + studCot] = si.ClassType2SumRankAF.Value;

                    if (si.ClassType2AvgRankF.HasValue)
                        row["平均_定期類別2排名" + studCot] = si.ClassType2AvgRankF.Value;

                    if (si.ClassType2AvgRankAF.HasValue)
                        row["加權平均_定期類別2排名" + studCot] = si.ClassType2AvgRankAF.Value;
                    #endregion

                    #region 評量成績排名_參考試別
                    if (si.ClassSumRefRank.HasValue)
                        row["總分班排名_參考試別" + studCot] = si.ClassSumRefRank.Value;

                    if (si.ClassSumRefRankA.HasValue)
                        row["加權總分班排名_參考試別" + studCot] = si.ClassSumRefRankA.Value;

                    if (si.ClassAvgRefRank.HasValue)
                        row["平均班排名_參考試別" + studCot] = si.ClassAvgRefRank.Value;

                    if (si.ClassAvgRefRankA.HasValue)
                        row["加權平均班排名_參考試別" + studCot] = si.ClassAvgRefRankA.Value;

                    if (si.YearSumRefRank.HasValue)
                        row["總分年排名_參考試別" + studCot] = si.YearSumRefRank.Value;

                    if (si.YearSumRefRankA.HasValue)
                        row["加權總分年排名_參考試別" + studCot] = si.YearSumRefRankA.Value;

                    if (si.YearAvgRefRank.HasValue)
                        row["平均年排名_參考試別" + studCot] = si.YearAvgRefRank.Value;

                    if (si.YearAvgRefRankA.HasValue)
                        row["加權平均年排名_參考試別" + studCot] = si.YearAvgRefRankA.Value;

                    if (si.ClassSumRefRankF.HasValue)
                        row["總分_定期班排名_參考試別" + studCot] = si.ClassSumRefRankF.Value;

                    if (si.ClassSumRefRankAF.HasValue)
                        row["加權總分_定期班排名_參考試別" + studCot] = si.ClassSumRefRankAF.Value;

                    if (si.ClassAvgRefRankF.HasValue)
                        row["平均_定期班排名_參考試別" + studCot] = si.ClassAvgRefRankF.Value;

                    if (si.ClassAvgRefRankAF.HasValue)
                        row["加權平均_定期班排名_參考試別" + studCot] = si.ClassAvgRefRankAF.Value;

                    if (si.YearSumRefRankF.HasValue)
                        row["總分_定期年排名_參考試別" + studCot] = si.YearSumRefRankF.Value;

                    if (si.YearSumRefRankAF.HasValue)
                        row["加權總分_定期年排名_參考試別" + studCot] = si.YearSumRefRankAF.Value;

                    if (si.YearAvgRefRankF.HasValue)
                        row["平均_定期年排名_參考試別" + studCot] = si.YearAvgRefRankF.Value;

                    if (si.YearAvgRefRankAF.HasValue)
                        row["加權平均_定期年排名_參考試別" + studCot] = si.YearAvgRefRankAF.Value;


                    if (si.ClassType1SumRefRank.HasValue)
                        row["總分類別1排名_參考試別" + studCot] = si.ClassType1SumRefRank.Value;

                    if (si.ClassType1SumRefRankA.HasValue)
                        row["加權總分類別1排名_參考試別" + studCot] = si.ClassType1SumRefRankA.Value;

                    if (si.ClassType1AvgRefRank.HasValue)
                        row["平均類別1排名_參考試別" + studCot] = si.ClassType1AvgRefRank.Value;

                    if (si.ClassType1AvgRefRankA.HasValue)
                        row["加權平均類別1排名_參考試別" + studCot] = si.ClassType1AvgRefRankA.Value;

                    if (si.ClassType1SumRefRankF.HasValue)
                        row["總分_定期類別1排名_參考試別" + studCot] = si.ClassType1SumRefRankF.Value;

                    if (si.ClassType1SumRefRankAF.HasValue)
                        row["加權總分_定期類別1排名_參考試別" + studCot] = si.ClassType1SumRefRankAF.Value;

                    if (si.ClassType1AvgRefRankF.HasValue)
                        row["平均_定期類別1排名_參考試別" + studCot] = si.ClassType1AvgRefRankF.Value;

                    if (si.ClassType1AvgRefRankAF.HasValue)
                        row["加權平均_定期類別1排名_參考試別" + studCot] = si.ClassType1AvgRefRankAF.Value;

                    if (si.ClassType2SumRefRank.HasValue)
                        row["總分類別2排名_參考試別" + studCot] = si.ClassType2SumRefRank.Value;

                    if (si.ClassType2SumRefRankA.HasValue)
                        row["加權總分類別2排名_參考試別" + studCot] = si.ClassType2SumRefRankA.Value;

                    if (si.ClassType2AvgRefRank.HasValue)
                        row["平均類別2排名_參考試別" + studCot] = si.ClassType2AvgRefRank.Value;

                    if (si.ClassType2AvgRefRankA.HasValue)
                        row["加權平均類別2排名_參考試別" + studCot] = si.ClassType2AvgRefRankA.Value;

                    if (si.ClassType2SumRefRankF.HasValue)
                        row["總分_定期類別2排名_參考試別" + studCot] = si.ClassType2SumRefRankF.Value;

                    if (si.ClassType2SumRefRankAF.HasValue)
                        row["加權總分_定期類別2排名_參考試別" + studCot] = si.ClassType2SumRefRankAF.Value;

                    if (si.ClassType2AvgRefRankF.HasValue)
                        row["平均_定期類別2排名_參考試別" + studCot] = si.ClassType2AvgRefRankF.Value;

                    if (si.ClassType2AvgRefRankAF.HasValue)
                        row["加權平均_定期類別2排名_參考試別" + studCot] = si.ClassType2AvgRefRankAF.Value;
                    #endregion

                    #region 評量成績排名_進退步
                    if (si.ClassSumRefRank.HasValue && si.ClassSumRank.HasValue)
                        row["總分班排名_進退步" + studCot] = si.ClassSumRefRank.Value - si.ClassSumRank.Value;

                    if (si.ClassSumRefRankA.HasValue && si.ClassSumRankA.HasValue)
                        row["加權總分班排名_進退步" + studCot] = si.ClassSumRefRankA.Value - si.ClassSumRankA.Value;

                    if (si.ClassAvgRefRank.HasValue && si.ClassAvgRank.HasValue)
                        row["平均班排名_進退步" + studCot] = si.ClassAvgRefRank.Value - si.ClassAvgRank.Value;

                    if (si.ClassAvgRefRankA.HasValue && si.ClassAvgRankA.HasValue)
                        row["加權平均班排名_進退步" + studCot] = si.ClassAvgRefRankA.Value - si.ClassAvgRankA.Value;

                    if (si.YearSumRefRank.HasValue && si.YearSumRank.HasValue)
                        row["總分年排名_進退步" + studCot] = si.YearSumRefRank.Value - si.YearSumRank.Value;

                    if (si.YearSumRefRankA.HasValue && si.YearSumRankA.HasValue)
                        row["加權總分年排名_進退步" + studCot] = si.YearSumRefRankA.Value - si.YearSumRankA.Value;

                    if (si.YearAvgRefRank.HasValue && si.YearAvgRank.HasValue)
                        row["平均年排名_進退步" + studCot] = si.YearAvgRefRank.Value - si.YearAvgRank.Value;

                    if (si.YearAvgRefRankA.HasValue && si.YearAvgRankA.HasValue)
                        row["加權平均年排名_進退步" + studCot] = si.YearAvgRefRankA.Value - si.YearAvgRankA.Value;

                    if (si.ClassSumRefRankF.HasValue && si.ClassSumRankF.HasValue)
                        row["總分_定期班排名_進退步" + studCot] = si.ClassSumRefRankF.Value - si.ClassSumRankF.Value;

                    if (si.ClassSumRefRankAF.HasValue && si.ClassSumRankAF.HasValue)
                        row["加權總分_定期班排名_進退步" + studCot] = si.ClassSumRefRankAF.Value - si.ClassSumRankAF.Value;

                    if (si.ClassAvgRefRankF.HasValue && si.ClassAvgRankF.HasValue)
                        row["平均_定期班排名_進退步" + studCot] = si.ClassAvgRefRankF.Value - si.ClassAvgRankF.Value;

                    if (si.ClassAvgRefRankAF.HasValue && si.ClassAvgRankAF.HasValue)
                        row["加權平均_定期班排名_進退步" + studCot] = si.ClassAvgRefRankAF.Value - si.ClassAvgRankAF.Value;

                    if (si.YearSumRefRankF.HasValue && si.YearSumRankF.HasValue)
                        row["總分_定期年排名_進退步" + studCot] = si.YearSumRefRankF.Value - si.YearSumRankF.Value;

                    if (si.YearSumRefRankAF.HasValue && si.YearSumRankAF.HasValue)
                        row["加權總分_定期年排名_進退步" + studCot] = si.YearSumRefRankAF.Value - si.YearSumRankAF.Value;

                    if (si.YearAvgRefRankF.HasValue && si.YearAvgRankF.HasValue)
                        row["平均_定期年排名_進退步" + studCot] = si.YearAvgRefRankF.Value - si.YearAvgRankF.Value;

                    if (si.YearAvgRefRankAF.HasValue && si.YearAvgRankAF.HasValue)
                        row["加權平均_定期年排名_進退步" + studCot] = si.YearAvgRefRankAF.Value - si.YearAvgRankAF.Value;


                    if (si.ClassType1SumRefRank.HasValue && si.ClassType1SumRank.HasValue)
                        row["總分類別1排名_進退步" + studCot] = si.ClassType1SumRefRank.Value - si.ClassType1SumRank.Value;

                    if (si.ClassType1SumRefRankA.HasValue && si.ClassType1SumRankA.HasValue)
                        row["加權總分類別1排名_進退步" + studCot] = si.ClassType1SumRefRankA.Value - si.ClassType1SumRankA.Value;

                    if (si.ClassType1AvgRefRank.HasValue && si.ClassType1AvgRank.HasValue)
                        row["平均類別1排名_進退步" + studCot] = si.ClassType1AvgRefRank.Value - si.ClassType1AvgRank.Value;

                    if (si.ClassType1AvgRefRankA.HasValue && si.ClassType1AvgRankA.HasValue)
                        row["加權平均類別1排名_進退步" + studCot] = si.ClassType1AvgRefRankA.Value - si.ClassType1AvgRankA.Value;

                    if (si.ClassType1SumRefRankF.HasValue && si.ClassType1SumRankF.HasValue)
                        row["總分_定期類別1排名_進退步" + studCot] = si.ClassType1SumRefRankF.Value - si.ClassType1SumRankF.Value;

                    if (si.ClassType1SumRefRankAF.HasValue && si.ClassType1SumRankAF.HasValue)
                        row["加權總分_定期類別1排名_進退步" + studCot] = si.ClassType1SumRefRankAF.Value - si.ClassType1SumRankAF.Value;

                    if (si.ClassType1AvgRefRankF.HasValue && si.ClassType1AvgRankF.HasValue)
                        row["平均_定期類別1排名_進退步" + studCot] = si.ClassType1AvgRefRankF.Value - si.ClassType1AvgRankF.Value;

                    if (si.ClassType1AvgRefRankAF.HasValue && si.ClassType1AvgRankAF.HasValue)
                        row["加權平均_定期類別1排名_進退步" + studCot] = si.ClassType1AvgRefRankAF.Value - si.ClassType1AvgRankAF.Value;

                    if (si.ClassType2SumRefRank.HasValue && si.ClassType2SumRank.HasValue)
                        row["總分類別2排名_進退步" + studCot] = si.ClassType2SumRefRank.Value - si.ClassType2SumRank.Value;

                    if (si.ClassType2SumRefRankA.HasValue && si.ClassType2SumRankA.HasValue)
                        row["加權總分類別2排名_進退步" + studCot] = si.ClassType2SumRefRankA.Value - si.ClassType2SumRankA.Value;

                    if (si.ClassType2AvgRefRank.HasValue && si.ClassType2AvgRank.HasValue)
                        row["平均類別2排名_進退步" + studCot] = si.ClassType2AvgRefRank.Value - si.ClassType2AvgRank.Value;

                    if (si.ClassType2AvgRefRankA.HasValue && si.ClassType2AvgRankA.HasValue)
                        row["加權平均類別2排名_進退步" + studCot] = si.ClassType2AvgRefRankA.Value - si.ClassType2AvgRankA.Value;

                    if (si.ClassType2SumRefRankF.HasValue && si.ClassType2SumRankF.HasValue)
                        row["總分_定期類別2排名_進退步" + studCot] = si.ClassType2SumRefRankF.Value - si.ClassType2SumRankF.Value;

                    if (si.ClassType2SumRefRankAF.HasValue && si.ClassType2SumRankAF.HasValue)
                        row["加權總分_定期類別2排名_進退步" + studCot] = si.ClassType2SumRefRankAF.Value - si.ClassType2SumRankAF.Value;

                    if (si.ClassType2AvgRefRankF.HasValue && si.ClassType2AvgRankF.HasValue)
                        row["平均_定期類別2排名_進退步" + studCot] = si.ClassType2AvgRefRankF.Value - si.ClassType2AvgRankF.Value;

                    if (si.ClassType2AvgRefRankAF.HasValue && si.ClassType2AvgRankAF.HasValue)
                        row["加權平均_定期類別2排名_進退步" + studCot] = si.ClassType2AvgRefRankAF.Value - si.ClassType2AvgRankAF.Value;
                    #endregion

                    tmpDomainInfoDict.Clear();
                    tmpSubjectInfoDict.Clear();
                    // 領域科目成績
                    foreach (DomainInfo di in si.DomainInfoList)
                    {
                        if (!tmpDomainInfoDict.ContainsKey(di.Name))
                            tmpDomainInfoDict.Add(di.Name, di);

                        // 領域的科目
                        if (SelDomainSubjectDict.ContainsKey(di.Name))
                        {
                            string d1 = di.Name + "領域成績" + studCot;
                            string d1_1 = di.Name + "領域_定期成績" + studCot;
                            if (dtTable.Columns.Contains(d1))
                            {
                                if (di.Score.HasValue)
                                    row[d1] = Math.Round(di.Score.Value, parseNumber, MidpointRounding.AwayFromZero);
                            }

                            if (dtTable.Columns.Contains(d1_1))
                            {
                                // 定期
                                if (di.ScoreF.HasValue)
                                {
                                    if (Program.ScoreValueMap.ContainsKey(di.ScoreF.Value))
                                    {
                                        row[d1_1] = Program.ScoreValueMap[di.ScoreF.Value].UseText;
                                    } 
                                    else
                                    {
                                        row[d1_1] = Math.Round(di.ScoreF.Value, parseNumber, MidpointRounding.AwayFromZero);
                                    }
                                }                                    
                            }

                            string d2 = di.Name + "領域學分" + studCot;
                            if (dtTable.Columns.Contains(d2))
                            {
                                if (di.Credit.HasValue)
                                {
                                    if (!tmpDomainCreditDict.ContainsKey(di.Name))
                                        tmpDomainCreditDict.Add(di.Name, new List<decimal>());

                                    if (!tmpDomainCreditDict[di.Name].Contains(di.Credit.Value))
                                        tmpDomainCreditDict[di.Name].Add(di.Credit.Value);

                                    row[d2] = di.Credit.Value;
                                }
                            }

                            int subjCot = 1;

                            foreach (string ssName in SelDomainSubjectDict[di.Name])
                            {
                                // 班上學生有這再處理
                                if (hasSubjectNameList.Contains(ssName))
                                {
                                    foreach (SubjectInfo subj in di.SubjectInfoList)
                                    {
                                        if (!tmpSubjectInfoDict.ContainsKey(subj.Name))
                                            tmpSubjectInfoDict.Add(subj.Name, subj);

                                        if (subj.Name == ssName)
                                        {
                                            string s1Key = di.Name + "領域_科目名稱" + studCot + "_" + subjCot;
                                            string s2Key = di.Name + "領域_科目成績" + studCot + "_" + subjCot;
                                            string s2_1Key = di.Name + "領域_科目_定期成績" + studCot + "_" + subjCot;
                                            string s3Key = di.Name + "領域_科目學分" + studCot + "_" + subjCot;
                                            string s4Key = di.Name + "領域_科目_平時成績" + studCot + "_" + subjCot;


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

                                            // 定期成績
                                            if (dtTable.Columns.Contains(s2_1Key))
                                            {
                                                if (dtTable.Columns.Contains(s2_1Key))
                                                    if (subj.ScoreF.HasValue)
                                                    {
                                                        if (Program.ScoreValueMap.ContainsKey(subj.ScoreF.Value))
                                                        {
                                                            row[s2_1Key] = Program.ScoreValueMap[subj.ScoreF.Value].UseText;
                                                        }
                                                        else
                                                        {
                                                            row[s2_1Key] = subj.ScoreF.Value;
                                                        }
                                                    }
                                                        
                                            }

                                            if (dtTable.Columns.Contains(s3Key))
                                            {
                                                if (subj.Credit.HasValue)
                                                {
                                                    row[s3Key] = subj.Credit.Value;

                                                    if (!tmpSubjectCreditDict.ContainsKey(subj.Name))
                                                        tmpSubjectCreditDict.Add(subj.Name, new List<decimal>());

                                                    if (!tmpSubjectCreditDict[subj.Name].Contains(subj.Credit.Value))
                                                        tmpSubjectCreditDict[subj.Name].Add(subj.Credit.Value);
                                                }

                                            }

                                            // 平時成績
                                            if (dtTable.Columns.Contains(s4Key))
                                            {
                                                if (dtTable.Columns.Contains(s4Key))
                                                    if (subj.ScoreA.HasValue)
                                                    {
                                                        if (Program.ScoreValueMap.ContainsKey(subj.ScoreA.Value))
                                                        {
                                                            row[s4Key] = Program.ScoreValueMap[subj.ScoreA.Value].UseText;
                                                        }
                                                        else
                                                        {
                                                            row[s4Key] = subj.ScoreA.Value;
                                                        }
                                                    }
                                            }
                                        }
                                    }
                                    subjCot++;
                                }

                            }
                        }

                    }

                    int daCount = 1;
                    // 單領域
                    foreach (string dName in SelDomainNameList)
                    {
                        if (tmpDomainInfoDict.ContainsKey(dName))
                        {
                            DomainInfo di = tmpDomainInfoDict[dName];
                            string da1 = "學生_領域名稱" + studCot + "_" + daCount;
                            string da2 = "學生_領域成績" + studCot + "_" + daCount;
                            string da3 = "學生_領域_定期成績" + studCot + "_" + daCount;
                            string da4 = "學生_領域學分" + studCot + "_" + daCount;

                            if (dtTable.Columns.Contains(da1))
                            {
                                row[da1] = di.Name;
                            }
                            if (dtTable.Columns.Contains(da2))
                            {
                                if (di.Score.HasValue)
                                    row[da2] = Math.Round(di.Score.Value, parseNumber, MidpointRounding.AwayFromZero);
                            }

                            if (dtTable.Columns.Contains(da3))
                            {
                                // 定期
                                if (di.ScoreF.HasValue)
                                {
                                    if (Program.ScoreValueMap.ContainsKey(di.ScoreF.Value))
                                    {
                                        row[da3] = Program.ScoreValueMap[di.ScoreF.Value].UseText; 
                                    }
                                    else
                                    {
                                        row[da3] = Math.Round(di.ScoreF.Value, parseNumber, MidpointRounding.AwayFromZero);
                                    }
                                }                                    
                            }

                            if (dtTable.Columns.Contains(da4))
                            {
                                if (di.Credit.HasValue)
                                {
                                    row[da4] = di.Credit.Value;
                                }
                            }
                        }
                        daCount++;
                    }

                    int saCount = 1;
                    // 單科目
                    foreach (string sName in SelSubjectNameList)
                    {
                        if (tmpSubjectInfoDict.ContainsKey(sName))
                        {


                            SubjectInfo sia = tmpSubjectInfoDict[sName];
                            string sa1 = "學生_科目名稱" + studCot + "_" + saCount;
                            string sa2 = "學生_科目成績" + studCot + "_" + saCount;
                            string sa3 = "學生_科目_定期成績" + studCot + "_" + saCount;
                            string sa4 = "學生_科目學分" + studCot + "_" + saCount;
                            string sa5 = "學生_科目_平時成績" + studCot + "_" + saCount;

                            if (dtTable.Columns.Contains(sa1))
                            {
                                row[sa1] = sia.Name;
                            }
                            if (dtTable.Columns.Contains(sa2))
                            {
                                if (sia.Score.HasValue)
                                    row[sa2] = Math.Round(sia.Score.Value, parseNumber, MidpointRounding.AwayFromZero);
                            }

                            if (dtTable.Columns.Contains(sa3))
                            {
                                // 定期
                                if (sia.ScoreF.HasValue)
                                {
                                    if (Program.ScoreValueMap.ContainsKey(sia.ScoreF.Value))
                                    {
                                        row[sa3] = Program.ScoreValueMap[sia.ScoreF.Value].UseText;
                                    }
                                    else
                                    {
                                        row[sa3] = Math.Round(sia.ScoreF.Value, parseNumber, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }

                            if (dtTable.Columns.Contains(sa4))
                            {
                                if (sia.Credit.HasValue)
                                {
                                    row[sa4] = sia.Credit.Value;
                                }
                            }

                            if (dtTable.Columns.Contains(sa5))
                            {
                                // 平時成績
                                if (sia.ScoreA.HasValue)
                                {
                                    if (Program.ScoreValueMap.ContainsKey(sia.ScoreA.Value))
                                    {
                                        row[sa5] = Program.ScoreValueMap[sia.ScoreA.Value].UseText;
                                    }
                                    else
                                    {
                                        row[sa5] = Math.Round(sia.ScoreA.Value, parseNumber, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }

                        }
                        saCount++;
                    }
                    studCot += 1;
                }

                // 填入學生領域與科目名稱
                int t_d = 1;
                foreach (string dName in SelDomainNameList)
                {
                    row["學生_領域名稱" + t_d] = dName;

                    if (tmpDomainCreditDict.ContainsKey(dName))
                    {
                        row["學生_領域學分" + t_d] = string.Join(",", tmpDomainCreditDict[dName].ToArray());
                    }

                    t_d++;
                }
                t_d = 1;
                foreach (string sName in SelSubjectNameList)
                {
                    row["學生_科目名稱" + t_d] = sName;
                    if (tmpSubjectCreditDict.ContainsKey(sName))
                    {
                        row["學生_科目學分" + t_d] = string.Join(",", tmpSubjectCreditDict[sName].ToArray());
                    }
                    t_d++;
                }

                List<string> tmpNameList = new List<string>();
                // 動態建立 table colmne
                for (int dd = 1; dd <= SelDomainNameList.Count; dd++)
                {
                    tmpNameList.Add("班級_領域成績_名稱" + dd);
                    tmpNameList.Add("年級_領域成績_名稱" + dd);
                    foreach (string rt in rankTypeList)
                    {
                        tmpNameList.Add("班級_領域成績_" + dd + "_" + rt);
                        tmpNameList.Add("班級_領域定期成績_" + dd + "_" + rt);
                        tmpNameList.Add("年級_領域成績_" + dd + "_" + rt);
                        tmpNameList.Add("年級_領域定期成績_" + dd + "_" + rt);
                    }
                }

                for (int ss = 1; ss <= SelSubjectNameList.Count; ss++)
                {
                    tmpNameList.Add("班級_科目成績_名稱" + ss);
                    tmpNameList.Add("年級_科目成績_名稱" + ss);
                    foreach (string rt in rankTypeList)
                    {
                        tmpNameList.Add("班級_科目成績_" + ss + "_" + rt);
                        tmpNameList.Add("班級_科目定期成績_" + ss + "_" + rt);
                        tmpNameList.Add("年級_科目成績_" + ss + "_" + rt);
                        tmpNameList.Add("年級_科目定期成績_" + ss + "_" + rt);
                    }
                }

                foreach (string name in tmpNameList)
                {
                    if (!dtTable.Columns.Contains(name))
                        dtTable.Columns.Add(name);
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
                        {
                            row[dName + "班級領域_科目及格人數" + subjCot] = ci.ClassScorePassCountDict[key];
                            subjCot++;
                        }
                    }
                }

                // 填入班組距
                if (ClassExamRankMatrixDict.ContainsKey(ci.ClassName))
                {
                    int dd = 1;
                    // 領域
                    foreach (string dName in SelDomainNameList)
                    {
                        string vKey = "定期評量/領域成績" + dName + "班排名";
                        if (ClassExamRankMatrixDict[ci.ClassName].ContainsKey(vKey))
                        {
                            if (dtTable.Columns.Contains("班級_領域成績_名稱" + dd))
                            {
                                row["班級_領域成績_名稱" + dd] = dName;
                            }
                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[ci.ClassName][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[ci.ClassName][vKey][rt].ToString();

                                string cKey = "班級_領域成績_" + dd + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }

                        vKey = "定期評量_定期/領域成績" + dName + "班排名";
                        if (ClassExamRankMatrixDict[ci.ClassName].ContainsKey(vKey))
                        {
                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[ci.ClassName][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[ci.ClassName][vKey][rt].ToString();

                                string cKey = "班級_領域定期成績_" + dd + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }
                        dd++;
                    }

                    // 科目成績
                    int ss = 1;

                    foreach (string subjName in SelSubjectNameList)
                    {
                        string vKey = "定期評量/科目成績" + subjName + "班排名";
                        if (ClassExamRankMatrixDict[ci.ClassName].ContainsKey(vKey))
                        {
                            if (dtTable.Columns.Contains("班級_科目成績_名稱" + ss))
                            {
                                row["班級_科目成績_名稱" + ss] = subjName;
                            }
                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[ci.ClassName][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[ci.ClassName][vKey][rt].ToString();

                                string cKey = "班級_科目成績_" + ss + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }

                        vKey = "定期評量_定期/科目成績" + subjName + "班排名";
                        if (ClassExamRankMatrixDict[ci.ClassName].ContainsKey(vKey))
                        {
                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[ci.ClassName][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[ci.ClassName][vKey][rt].ToString();

                                string cKey = "班級_科目定期成績_" + ss + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }
                        ss++;
                    }
                }


                // 填入年組距
                string grc = ci.GradeYear + "年級";
                if (ClassExamRankMatrixDict.ContainsKey(grc))
                {
                    int dd = 1;
                    // 領域
                    foreach (string dName in SelDomainNameList)
                    {
                        string vKey = "定期評量/領域成績" + dName + "年排名";
                        if (ClassExamRankMatrixDict[grc].ContainsKey(vKey))
                        {
                            if (dtTable.Columns.Contains("年級_領域成績_名稱" + dd))
                            {
                                row["年級_領域成績_名稱" + dd] = dName;
                            }

                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[grc][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[grc][vKey][rt].ToString();

                                string cKey = "年級_領域成績_" + dd + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }

                        vKey = "定期評量_定期/領域成績" + dName + "年排名";
                        if (ClassExamRankMatrixDict[grc].ContainsKey(vKey))
                        {
                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[grc][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[grc][vKey][rt].ToString();

                                string cKey = "年級_領域定期成績_" + dd + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }
                        dd++;
                    }

                    // 科目成績
                    int ss = 1;
                    foreach (string subjName in SelSubjectNameList)
                    {
                        string vKey = "定期評量/科目成績" + subjName + "年排名";
                        if (ClassExamRankMatrixDict[grc].ContainsKey(vKey))
                        {
                            if (dtTable.Columns.Contains("年級_科目成績_名稱" + ss))
                            {
                                row["年級_科目成績_名稱" + ss] = subjName;
                            }

                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[grc][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[grc][vKey][rt].ToString();

                                string cKey = "年級_科目成績_" + ss + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }

                        vKey = "定期評量_定期/科目成績" + subjName + "年排名";
                        if (ClassExamRankMatrixDict[grc].ContainsKey(vKey))
                        {
                            foreach (string rt in rankTypeList)
                            {
                                string value = "";
                                if (ClassExamRankMatrixDict[grc][vKey][rt] != null)
                                    value = ClassExamRankMatrixDict[grc][vKey][rt].ToString();

                                string cKey = "年級_科目定期成績_" + ss + "_" + rt;
                                if (dtTable.Columns.Contains(cKey))
                                {
                                    if (cKey.Contains("avg"))
                                        row[cKey] = parseScore1(value);
                                    else
                                        row[cKey] = value;
                                }
                            }
                        }
                        ss++;
                    }

                }


                //int rowClsCount = 1;
                //// 處理班級PR、百分比、五標、組距 領域
                //foreach (int gr in grClassNameDict.Keys)
                //{
                //    foreach (string className in grClassNameDict[gr])
                //    {
                //        row["班級" + rowClsCount + "_名稱"] = className;
                //        if (ClassExamRankMatrixDict.ContainsKey(className))
                //        {
                //            // 領域
                //            foreach (string dName in tmpDomainNameList)
                //            {
                //                string vKey = "定期評量/領域成績" + dName + "班排名";
                //                if (ClassExamRankMatrixDict[className].ContainsKey(vKey))
                //                {
                //                    foreach (string rt in rankTypeList)
                //                    {
                //                        string value = "";
                //                        if (ClassExamRankMatrixDict[className][vKey][rt] != null)
                //                            value = ClassExamRankMatrixDict[className][vKey][rt].ToString();

                //                        string cKey = "班級" + rowClsCount + "_" + dName + "_領域成績_" + rt;
                //                        if (dtTable.Columns.Contains(cKey))
                //                        {
                //                            if (cKey.Contains("avg"))
                //                                row[cKey] = parseScore1(value);
                //                            else
                //                                row[cKey] = value;
                //                        }
                //                    }
                //                }

                //                vKey = "定期評量_定期/領域成績" + dName + "班排名";
                //                if (ClassExamRankMatrixDict[className].ContainsKey(vKey))
                //                {
                //                    foreach (string rt in rankTypeList)
                //                    {
                //                        string value = "";
                //                        if (ClassExamRankMatrixDict[className][vKey][rt] != null)
                //                            value = ClassExamRankMatrixDict[className][vKey][rt].ToString();

                //                        string cKey = "班級" + rowClsCount + "_" + dName + "_領域定期成績_" + rt;
                //                        if (dtTable.Columns.Contains(cKey))
                //                        {
                //                            if (cKey.Contains("avg"))
                //                                row[cKey] = parseScore1(value);
                //                            else
                //                                row[cKey] = value;
                //                        }
                //                    }
                //                }
                //            }

                //            // 科目成績
                //            int subjIdex = 1;

                //            foreach (string subjName in tmpSubjectNameList)
                //            {
                //                row["班級" + rowClsCount + "_科目名稱_" + subjIdex] = subjName;
                //                string vKey = "定期評量/科目成績" + subjName + "班排名";
                //                if (ClassExamRankMatrixDict[className].ContainsKey(vKey))
                //                {
                //                    foreach (string rt in rankTypeList)
                //                    {
                //                        string value = "";
                //                        if (ClassExamRankMatrixDict[className][vKey][rt] != null)
                //                            value = ClassExamRankMatrixDict[className][vKey][rt].ToString();

                //                        string cKey = "班級" + rowClsCount + "_科目成績_" + rt + subjIdex;
                //                        if (dtTable.Columns.Contains(cKey))
                //                        {
                //                            if (cKey.Contains("avg"))
                //                                row[cKey] = parseScore1(value);
                //                            else
                //                                row[cKey] = value;
                //                        }
                //                    }
                //                }

                //                vKey = "定期評量_定期/科目成績" + subjName + "班排名";
                //                if (ClassExamRankMatrixDict[className].ContainsKey(vKey))
                //                {
                //                    foreach (string rt in rankTypeList)
                //                    {
                //                        string value = "";
                //                        if (ClassExamRankMatrixDict[className][vKey][rt] != null)
                //                            value = ClassExamRankMatrixDict[className][vKey][rt].ToString();

                //                        string cKey = "班級" + rowClsCount + "_科目定期成績_" + rt + subjIdex;
                //                        if (dtTable.Columns.Contains(cKey))
                //                        {
                //                            if (cKey.Contains("avg"))
                //                                row[cKey] = parseScore1(value);
                //                            else
                //                                row[cKey] = value;
                //                        }
                //                    }
                //                }
                //                subjIdex++;
                //            }



                //            rowClsCount++;
                //        }
                //    }
                //}


                //int rowGrCount = 1;
                //// 處理年級
                //foreach (int year in grClassNameDict.Keys)
                //{
                //    string gr = year + "年級";
                //    if (ClassExamRankMatrixDict.ContainsKey(gr))
                //    {
                //        row["年級" + rowGrCount + "_名稱"] = gr;
                //        // 領域
                //        foreach (string dName in tmpDomainNameList)
                //        {
                //            string vKey = "定期評量/領域成績" + dName + "年排名";
                //            if (ClassExamRankMatrixDict[gr].ContainsKey(vKey))
                //            {
                //                foreach (string rt in rankTypeList)
                //                {
                //                    string value = "";
                //                    if (ClassExamRankMatrixDict[gr][vKey][rt] != null)
                //                        value = ClassExamRankMatrixDict[gr][vKey][rt].ToString();

                //                    string cKey = "年級" + rowGrCount + "_" + dName + "_領域成績_" + rt;
                //                    if (dtTable.Columns.Contains(cKey))
                //                    {
                //                        if (cKey.Contains("avg"))
                //                            row[cKey] = parseScore1(value);
                //                        else
                //                            row[cKey] = value;
                //                    }
                //                }
                //            }

                //            vKey = "定期評量_定期/領域成績" + dName + "年排名";
                //            if (ClassExamRankMatrixDict[gr].ContainsKey(vKey))
                //            {
                //                foreach (string rt in rankTypeList)
                //                {
                //                    string value = "";
                //                    if (ClassExamRankMatrixDict[gr][vKey][rt] != null)
                //                        value = ClassExamRankMatrixDict[gr][vKey][rt].ToString();

                //                    string cKey = "年級" + rowGrCount + "_" + dName + "_領域定期成績_" + rt;
                //                    if (dtTable.Columns.Contains(cKey))
                //                    {
                //                        if (cKey.Contains("avg"))
                //                            row[cKey] = parseScore1(value);
                //                        else
                //                            row[cKey] = value;
                //                    }
                //                }
                //            }
                //        }

                //        // 科目成績
                //        int subjIdex = 1;

                //        foreach (string subjName in tmpSubjectNameList)
                //        {
                //            row["年級" + rowGrCount + "_科目名稱_" + subjIdex] = subjName;
                //            string vKey = "定期評量/科目成績" + subjName + "年排名";
                //            if (ClassExamRankMatrixDict[gr].ContainsKey(vKey))
                //            {
                //                foreach (string rt in rankTypeList)
                //                {
                //                    string value = "";
                //                    if (ClassExamRankMatrixDict[gr][vKey][rt] != null)
                //                        value = ClassExamRankMatrixDict[gr][vKey][rt].ToString();

                //                    string cKey = "年級" + rowGrCount + "_科目成績_" + rt + subjIdex;
                //                    if (dtTable.Columns.Contains(cKey))
                //                    {
                //                        if (cKey.Contains("avg"))
                //                            row[cKey] = parseScore1(value);
                //                        else
                //                            row[cKey] = value;
                //                    }
                //                }
                //            }

                //            vKey = "定期評量_定期/科目成績" + subjName + "年排名";
                //            if (ClassExamRankMatrixDict[gr].ContainsKey(vKey))
                //            {
                //                foreach (string rt in rankTypeList)
                //                {
                //                    string value = "";
                //                    if (ClassExamRankMatrixDict[gr][vKey][rt] != null)
                //                        value = ClassExamRankMatrixDict[gr][vKey][rt].ToString();

                //                    string cKey = "年級" + rowGrCount + "_科目定期成績_" + rt + subjIdex;
                //                    if (dtTable.Columns.Contains(cKey))
                //                    {
                //                        if (cKey.Contains("avg"))
                //                            row[cKey] = parseScore1(value);
                //                        else
                //                            row[cKey] = value;
                //                    }
                //                }
                //            }
                //            subjIdex++;
                //        }



                //        rowGrCount++;
                //    }

                //}

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





            bgWorkerReport.ReportProgress(100);
        }

        public string parseScore1(string str)
        {
            string value = "";
            decimal dc;
            if (decimal.TryParse(str, out dc))
            {
                value = Math.Round(dc, parseNumber).ToString();
            }

            return value;
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
            cboRefExam.Items.Clear();
            cboRefExam.Items.Add("");
            foreach (ExamRecord exName in _exams)
            {
                cboExam.Items.Add(exName.Name);
                cboRefExam.Items.Add(exName.Name);
            }



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

            if (string.IsNullOrEmpty(_Configure.RefExamID))
                _Configure.RefExamID = "";

            int idx = 0;
            foreach (ExamRecord exName in _exams)
            {
                if (exName.ID == _Configure.ExamRecordID)
                    cboExam.SelectedIndex = idx;

                if (exName.ID == _Configure.RefExamID)
                    cboRefExam.SelectedIndex = idx;

                idx++;
            }

            if (!string.IsNullOrEmpty(_Configure.SelSetConfigName))
                cboConfigure.Text = userSelectConfigName;

            ReloadCanSelectDomainSubject(cboSchoolYear.Text, cboSemester.Text, _ClassIDList, _Configure.ExamRecordID);

            btnSaveConfig.Enabled = btnPrint.Enabled = true;

        }

        /// <summary>
        /// 重新讀取可選科目領域
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="_ClassIDList"></param>
        /// <param name="ExamID"></param>
        private void ReloadCanSelectDomainSubject(string SchoolYear, string Semester, List<string> _ClassIDList, string ExamID)
        {
            CanSelectExamDomainSubjectDict = DAO.DataAccess.GetExamDomainSubjectDictByClass(SchoolYear, Semester, _ClassIDList, ExamID);
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
            _Configure.RefSchoolYear = cboRefSchoolYear.Text;
            _Configure.RefSemester = cboRefSemester.Text;
            _Configure.SelSetConfigName = cboConfigure.Text;

            // 四捨五入進位
            if (!int.TryParse(cboParseNumber.Text, out parseNumber))
            {
                MsgBox.Show("平均計算至小數點後需要輸入整數!");
                return;
            }

            _Configure.ParseNumber = parseNumber;
            _Configure.RefExamID = "";
            foreach (ExamRecord exm in _exams)
            {
                if (exm.Name == cboExam.Text)
                {
                    _Configure.ExamRecord = exm;
                }

                // 儲存參考試別編號
                if (exm.Name == cboRefExam.Text)
                {
                    _Configure.RefExamID = exm.ID;
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
            Program.ScoreTextMap.Clear();
            Program.ScoreValueMap.Clear();
            #region 取得評量成績缺考暨免試設定
            Framework.ConfigData cd = JHSchool.School.Configuration["評量成績缺考暨免試設定"];
            if (!string.IsNullOrEmpty(cd["評量成績缺考暨免試設定"]))
            {
                XmlElement element = Framework.XmlHelper.LoadXml(cd["評量成績缺考暨免試設定"]);

                foreach (XmlElement each in element.SelectNodes("Setting"))
                {
                    var UseText = each.SelectSingleNode("UseText").InnerText;
                    var AllowCalculation = bool.Parse(each.SelectSingleNode("AllowCalculation").InnerText);
                    decimal Score;
                    decimal.TryParse(each.SelectSingleNode("Score").InnerText, out Score);
                    var Active = bool.Parse(each.SelectSingleNode("Active").InnerText);
                    var UseValue = decimal.Parse(each.SelectSingleNode("UseValue").InnerText);

                    if (Active)
                    {
                        if (!Program.ScoreTextMap.ContainsKey(UseText))
                        {
                            Program.ScoreTextMap.Add(UseText, new DAO.ScoreMap
                            {
                                UseText = UseText,  //「缺」或「免」
                                AllowCalculation = AllowCalculation,  //是否計算成績
                                Score = Score,  //計算成績時，應以多少分來計算
                                Active = Active, //此設定是否啟用
                                UseValue = UseValue, //代表「缺」或「免」的負數
                            });
                        }
                        if (!Program.ScoreValueMap.ContainsKey(UseValue))
                        {
                            Program.ScoreValueMap.Add(UseValue, new DAO.ScoreMap
                            {
                                UseText = UseText,
                                AllowCalculation = AllowCalculation,
                                Score = Score,
                                Active = Active,
                                UseValue = UseValue,
                            });
                        }
                    }
                }
            }

            #endregion
            // 
            //List<string> ddList = DataAccess.GetDomainConfigSortName();

            int sy;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sy))
            {
                for (int i = (sy - 2); i <= (sy + 3); i++)
                {
                    cboSchoolYear.Items.Add(i);
                    cboRefSchoolYear.Items.Add(i);
                }
            }

            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");
            cboRefSemester.Items.Add("1");
            cboRefSemester.Items.Add("2");
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            cboSemester.Text = K12.Data.School.DefaultSemester;
            cboRefSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            cboRefSemester.Text = K12.Data.School.DefaultSemester;
            for (int i = 0; i <= 3; i++)
                cboParseNumber.Items.Add(i);

            //cboSchoolYear.Enabled = false;
            //cboSemester.Enabled = false;
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

        // 解析檔名
        public string ParseFileName(string fileName)
        {
            string name = fileName;

            if (fileName == null)
                throw new ArgumentNullException();

            if (name.Length == 0)
                throw new ArgumentException();

            if (name.Length > 245)
                throw new PathTooLongException();

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name;
        }


        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = ParseFileName("評量成績單樣板(" + _Configure.Name + ")");

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

                // 參考試別
                if (er.Name == cboRefExam.Text)
                {
                    SelRefExamID = er.ID;
                    SelRefExamName = er.Name;
                }
            }

            SelSchoolYear = cboSchoolYear.Text;
            SelSemester = cboSemester.Text;
            RefSelSchoolYear = cboRefSchoolYear.Text;
            RefSelSemester = cboRefSemester.Text;

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
            Global._SelRefExamID = SelRefExamID;

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

            ReloadCanSelectDomainSubject(cboSchoolYear.Text, cboSemester.Text, _ClassIDList, SelExamID);
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
                    cboRefSchoolYear.Text = _Configure.RefSchoolYear;
                    cboRefSemester.Text = _Configure.RefSemester;
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

                    if (_Configure.RefExamID != "")
                    {
                        int idx = 1;
                        foreach (ExamRecord rec in _exams)
                        {
                            if (rec.ID == _Configure.RefExamID)
                            {
                                cboRefExam.SelectedIndex = idx;
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

        private void cboRefExam_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (ExamRecord er in _exams)
            {
                if (er.Name == cboRefExam.Text)
                {
                    SelRefExamID = er.ID;
                    SelRefExamName = er.Name;
                }
            }
        }

    }
}
