﻿using System;
using System.Collections.Generic;
using System.Text;
using JHSchool.Data;
using System.ComponentModel;
using JHSchool.Evaluation.Ranking;
using Aspose.Cells;
using System.IO;
using K12.Data;
using JHSchool.Evaluation;
using HsinChu.ClassExamScoreAvgComparison.Model;
using FISCA.Presentation.Controls;
using FISCA.Presentation;
using System.Windows.Forms;
using FISCA.Data;
using System.Data;

namespace HsinChu.ClassExamScoreAvgComparison
{
    internal class Report
    {
        //private List<ComputeMethod> _methods;
        private List<ClassExamScoreData> _data;
        private Dictionary<string, JHCourseRecord> _courseDict;
        private JHExamRecord _exam;
        private List<string> _domains;
        private JHSchool.Evaluation.Calculation.ScoreCalculator _calc;
        public static int _UserSelectCount = 0;

        private BackgroundWorker _worker;

        public event EventHandler GenerateCompleted;
        public event EventHandler GenerateError;

        private string SchoolName;
        private string Semester;

        //科目資料管理
        List<string> subjectNameList = new List<string>();

        public Dictionary<decimal, DAL.ScoreMap> _ScoreValueMap = new Dictionary<decimal, DAL.ScoreMap>();
        public string _ScoreSource;

        //public Report(List<ClassExamScoreData> data, Dictionary<string, JHCourseRecord> courseDict, JHExamRecord exam, List<ComputeMethod> methods)
        //{
        public Report(List<ClassExamScoreData> data, Dictionary<string, JHCourseRecord> courseDict, JHExamRecord exam, List<string> domains, Dictionary<decimal, DAL.ScoreMap> scoreMapDic, string scoreSource)
        {
            _data = data;
            _courseDict = courseDict;
            _exam = exam;
            _domains = domains;
            _calc = new JHSchool.Evaluation.Calculation.ScoreCalculator(null);
            _ScoreValueMap = scoreMapDic;
            _ScoreSource = scoreSource;

            //_methods = methods;

            InitializeWorker();
            //GetSubjectList();
            SchoolName = School.ChineseName;
            // Semester = string.Format("{0}學年度 第{1}學期", School.DefaultSchoolYear, School.DefaultSemester);
            Semester = string.Format("{0}學年度 第{1}學期", Global.UserSelectSchoolYear, Global.UserSelectSemester);

        }

        /// <summary>
        /// 初始化 BackgroundWorker
        /// </summary>
        private void InitializeWorker()
        {
            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;
        }

        #region BackgroundWorker Events
        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MsgBox.Show("產生報表時發生錯誤" + e.Error.Message);
                if (GenerateError != null)
                    GenerateError.Invoke(this, new EventArgs());
                return;
            }

            #region 儲存
            Workbook book = e.Result as Workbook;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists) dir.Create();

            path = Path.Combine(path, Global.ReportName + ".xls");

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
                book.Save(path, FileFormatType.Excel2003);
                FISCA.LogAgent.ApplicationLog.Log("成績系統.報表", "列印" + Global.ReportName, string.Format("產生{0}份" + Global.ReportName, _data.Count));

                MotherForm.SetStatusBarMessage(Global.ReportName + "產生完成");
                if (GenerateCompleted != null)
                    GenerateCompleted.Invoke(this, new EventArgs());

                if (MsgBox.Show(Global.ReportName + "產生完成，是否立刻開啟？", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(path);
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存失敗");
            }
            #endregion

            #region 大掃除
            foreach (var ced in _data)
                ced.Clear();
            #endregion
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // 取得班級學生人數
            Dictionary<string, int> ClassStudCount = new Dictionary<string, int>();
            foreach (ClassExamScoreData cee in _data)
                if (!ClassStudCount.ContainsKey(cee.Class.ID))
                    ClassStudCount.Add(cee.Class.ID, cee.Students.Count);

            Workbook book = new Workbook();
            Worksheet ws;
            int rowIndex = 0;
            Workbook template = new Workbook();

            if (_data.Count > 30 || _UserSelectCount > 14)
            {
                book.Open(new MemoryStream(Resource1.班級評量成績平均比較表60));
                ws = book.Worksheets[0];
                template.Open(new MemoryStream(Resource1.班級評量成績平均比較表60));
            }
            else
            {
                book.Open(new MemoryStream(Resource1.班級評量成績平均比較表));
                ws = book.Worksheets[0];
                template.Open(new MemoryStream(Resource1.班級評量成績平均比較表));
            }
            Range all = template.Worksheets.GetRangeByName("All");

            //Range printDate = template.Worksheets.GetRangeByName("PrintDate");
            //printDate[0, 0].PutValue(DateTime.Today.ToString("yyyy/MM/dd"));

            Range title = template.Worksheets.GetRangeByName("Title");
            //Range feedback = template.Worksheets.GetRangeByName("Feedback");
            Range rowHeaders = template.Worksheets.GetRangeByName("RowHeaders");
            Range columnHeaders = template.Worksheets.GetRangeByName("ColumnHeaders");
            //Range rankColumnHeader = template.Worksheets.GetRangeByName("RankColumnHeader");

            int RowNumber = all.RowCount;
            int DataRowNumber = rowHeaders.RowCount;

            //_data.Sort(delegate(ClassExamScoreData x, ClassExamScoreData y)
            //{
            //    return x.Class.Name.CompareTo(y.Class.Name);
            //});

            int dataRowIndex = rowHeaders.FirstRow;

            List<string> headers = new List<string>();
            Dictionary<string, int> colMapping = new Dictionary<string, int>();
            foreach (ClassExamScoreData ced in _data)
            {
                //ws.Cells.CreateRange(rowIndex, RowNumber, false).Copy(all);

                //int classColIndex = 0;

                foreach (string courseID in ced.ValidCourseIDs)
                {
                    JHCourseRecord course = _courseDict[courseID];
                    if (!headers.Contains(course.Subject))
                        headers.Add(course.Subject);

                    #region comment
                    //if (!colMapping.ContainsKey(GetTaggedDomain(course.Domain)) && _domains.Contains(course.Domain))
                    //{
                    //    colMapping.Add(GetTaggedDomain(course.Domain), classColIndex);
                    //    ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue(course.Domain);
                    //    ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].Style.Font.IsBold = true;

                    //    classColIndex++;
                    //    ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue("排序");
                    //    classColIndex++;
                    //}
                    //if (!colMapping.ContainsKey(course.Subject))
                    //{
                    //    colMapping.Add(course.Subject, classColIndex);
                    //    ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue(course.Subject);

                    //    classColIndex++;
                    //    ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue("排序");
                    //    classColIndex++;
                    //}
                    #endregion
                }

                foreach (StudentRow row in ced.Rows.Values)
                {
                    foreach (var sce in row.RawScoreList)
                    {
                        if (sce.RefExamID != _exam.ID) continue;
                        JHCourseRecord course = _courseDict[sce.RefCourseID];

                        if (!headers.Contains(GetTaggedDomain(course.Domain)))
                            headers.Add(GetTaggedDomain(course.Domain));
                    }
                }
            }

            #region 產生 Headers
            headers.Sort(Sort);
            int classColIndex = 0;
            foreach (string each in headers)
            {
                if (IsDomain(each))
                {
                    if (_domains.Contains(GetOriginalDomain(each)))
                    {
                        colMapping.Add(each, classColIndex);
                        ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue(GetOriginalDomain(each));
                        ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].Style.Font.IsBold = true;
                        classColIndex++;
                        ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue("排序");
                        classColIndex++;
                    }
                }
                else
                {
                    colMapping.Add(each, classColIndex);
                    ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue(each);
                    classColIndex++;
                    ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue("排序");
                    classColIndex++;
                }
            }
            #endregion

            foreach (ClassExamScoreData ced in _data)
            {
                Dictionary<string, decimal?> subjectScores = new Dictionary<string, decimal?>();
                Dictionary<string, int> subjectCount = new Dictionary<string, int>();

                Dictionary<string, decimal?> subjectCredits = new Dictionary<string, decimal?>();


                foreach (StudentRow row in ced.Rows.Values)
                {
                    foreach (var sce in row.RawScoreList)
                    {
                        List<string> asIDs = new List<string>();
                        Dictionary<string, JHAEIncludeRecord> aeDict = new Dictionary<string, JHAEIncludeRecord>();
                        if (sce.RefExamID != _exam.ID) continue;
                        JHCourseRecord course = _courseDict[sce.RefCourseID];


                        if (!subjectCredits.ContainsKey(GetDomainSubjectKey(course.Domain, course.Subject)))
                            subjectCredits.Add(GetDomainSubjectKey(course.Domain, course.Subject), course.Credit);

                        string asID = _courseDict[sce.RefCourseID].RefAssessmentSetupID;
                        if (_ScoreSource != "定期")
                        {
                            asIDs.Add(asID);

                            foreach (JHAEIncludeRecord record in JHAEInclude.SelectByAssessmentSetupIDs(asIDs))
                            {
                                if (record.RefExamID != _exam.ID)
                                    continue;
                                aeDict.Add(record.RefAssessmentSetupID, record);
                            }

                        }


                        //if (!headers.Contains(GetTaggedDomain(course.Domain)))
                        //{
                        //    headers.Add(GetTaggedDomain(course.Domain));
                        //}
                        //if (!colMapping.ContainsKey(GetTaggedDomain(course.Domain)) && _domains.Contains(course.Domain))
                        //{
                        //    colMapping.Add(GetTaggedDomain(course.Domain), classColIndex);
                        //    //ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue(course.Domain);
                        //    //ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].Style.Font.IsBold = true;

                        //    classColIndex += 2;
                        //    //ws.Cells[columnHeaders.FirstRow + rowIndex, columnHeaders.FirstColumn + classColIndex].PutValue("排序");
                        //    //classColIndex++;
                        //}

                        string ds = GetDomainSubjectKey(course.Domain, course.Subject);
                        if (!subjectScores.ContainsKey(ds))
                        {
                            subjectScores.Add(ds, 0);
                            subjectCount.Add(ds, 0);
                        }
                        if (_ScoreSource == "定期")
                        {
                            if (sce.Score.HasValue)
                            {
                                decimal? score = null;
                                if (_ScoreValueMap.ContainsKey(sce.Score.Value))
                                {
                                    if (_ScoreValueMap[sce.Score.Value].AllowCalculation) //可以被計算 ex缺考0分
                                    {
                                        score = _ScoreValueMap[sce.Score.Value].Score.Value;
                                    }
                                    else  //不能被計算 ex 免試
                                    {

                                    }
                                }
                                else
                                {
                                    score = sce.Score.Value;
                                }

                                if (score != null)
                                {
                                    subjectCount[ds]++;
                                    //decimal value = decimal.Zero;
                                    //if (row.StudentScore.CalculationRule != null)
                                    //    value = row.StudentScore.CalculationRule.ParseSubjectScore(score.Value);
                                    //else
                                    //    value = _calc.ParseSubjectScore(score.Value);

                                    //subjectScores[ds] += value;

                                    //2023-01-30 Cynthia 以原始評量分數去計算
                                    subjectScores[ds] += score;
                                }
                            }
                        }
                        else
                        {
                            decimal? score = null;

                            //對照後的 定期評量分數
                            decimal? scoreA = null;
                            //對照後的 平時評量分數
                            decimal? assignmentScore = null;

                            //找對照
                            if (sce.Score.HasValue)
                            {
                                if (_ScoreValueMap.ContainsKey(sce.Score.Value))
                                {
                                    if (_ScoreValueMap[sce.Score.Value].AllowCalculation) //可以被計算 ex缺考0分
                                        scoreA = _ScoreValueMap[sce.Score.Value].Score.Value;
                                    else  //不能被計算 ex 免試
                                    {
                                        scoreA = null;
                                    }
                                }
                                else
                                {
                                    scoreA = sce.Score.Value;
                                }
                            }
                            if (sce.AssignmentScore.HasValue)
                            {
                                if (_ScoreValueMap.ContainsKey(sce.AssignmentScore.Value))
                                {
                                    if (_ScoreValueMap[sce.AssignmentScore.Value].AllowCalculation) //可以被計算 ex缺考0分
                                        assignmentScore = _ScoreValueMap[sce.AssignmentScore.Value].Score.Value;
                                    else  //不能被計算 ex 免試
                                    {
                                        assignmentScore = null;
                                    }
                                }
                                else
                                {
                                    assignmentScore = sce.AssignmentScore.Value;
                                }
                            }

                            if (scoreA != null && assignmentScore != null)
                            {
                                if (Global.ScorePercentageHSDict.ContainsKey(aeDict[asID].RefAssessmentSetupID))
                                {
                                    decimal ff = Global.ScorePercentageHSDict[aeDict[asID].RefAssessmentSetupID];
                                    decimal f = scoreA.Value * ff * 0.01M;
                                    decimal a = assignmentScore.Value * (100 - ff) * 0.01M;

                                    score = f + a;
                                }
                                else
                                    score = scoreA.Value * 0.5M + assignmentScore.Value * 0.5M;


                            }
                            else if (scoreA != null)
                                score = scoreA;
                            else if (assignmentScore != null)
                                score = assignmentScore;


                            if (score != null)
                            {
                                subjectCount[ds]++;
                                //decimal value = decimal.Zero;
                                //if (row.StudentScore.CalculationRule != null)
                                //    value = row.StudentScore.CalculationRule.ParseSubjectScore(score.Value);
                                //else
                                //    value = _calc.ParseSubjectScore(score.Value);
                                //subjectScores[ds] += value;
                                subjectScores[ds] += score;
                            }
                        }
                    }
                }

                foreach (string ds in new List<string>(subjectScores.Keys))
                {
                    if (subjectScores[ds].HasValue && subjectCount[ds] > 0)
                        subjectScores[ds] = subjectScores[ds].Value / (decimal)subjectCount[ds];
                }

                Dictionary<string, decimal?> domainScores = new Dictionary<string, decimal?>();
                Dictionary<string, int> domainCount = new Dictionary<string, int>();

                Dictionary<string, decimal?> domainCredits = new Dictionary<string, decimal?>();

                //foreach (string ds in subjectScores.Keys)
                //{
                //    string domain = GetOnlyDomain(ds);
                //    if (!domainScores.ContainsKey(domain))
                //    {
                //        domainScores.Add(domain, 0);
                //        domainCount.Add(domain, 0);
                //    }
                //    domainCount[domain]++;
                //    domainScores[domain] += subjectScores[ds];
                //}

                foreach (string ds in subjectScores.Keys)
                {
                    string domain = GetOnlyDomain(ds);

                    if (!domainScores.ContainsKey(domain))
                        domainScores.Add(domain, 0);

                    if (subjectCredits.ContainsKey(ds))
                    {
                        domainScores[domain] += subjectScores[ds] * subjectCredits[ds];
                    }
                }


                foreach (string d in subjectCredits.Keys)
                {
                    string domain = GetOnlyDomain(d);
                    if (!domainCredits.ContainsKey(domain))
                    {
                        domainCredits.Add(domain, 0);
                    }
                    domainCredits[domain] += subjectCredits[d];
                }

                foreach (string domain in new List<string>(domainScores.Keys))
                {
                    //if (domainScores[domain].HasValue && domainCount[domain] > 0)
                    //    domainScores[domain] = domainScores[domain].Value / (decimal)domainCount[domain];
                    if (domainScores[domain].HasValue && domainCredits.ContainsKey(domain) && domainCredits[domain].HasValue)
                    {
                        // 判斷當權數不為0才計算
                        if ((decimal)domainCredits[domain] > 0)
                            domainScores[domain] = domainScores[domain].Value / (decimal)domainCredits[domain];
                    }

                }

                #region 填入班級平均
                //int rankIndex = rankColumnHeader.FirstColumn;

                // 班級平均
                ClassCourseAvg cca = new ClassCourseAvg();

                foreach (var stud in ced.Students)
                {
                    if (ced.Rows.ContainsKey(stud.ID))
                    {
                        //阿寶...這也許是地雷....

                        // 當學生不屬於任何班級跳過
                        if (stud.Class == null)
                            continue;

                        cca.ClassID = stud.Class.ID;
                        cca.ClassName = stud.Class.Name;
                        StudentRow srow = ced.Rows[stud.ID];
                        foreach (CourseScore cs in srow.CourseScoreList)
                        {
                            if (cs.Score.HasValue)
                            {
                                cca.ClassID = stud.Class.ID;
                                cca.ClassName = stud.Class.Name;
                                //decimal value = decimal.Zero;
                                //if (srow.StudentScore.CalculationRule != null)
                                //    value = srow.StudentScore.CalculationRule.ParseSubjectScore(cs.Score.Value);
                                //else
                                //    value = _calc.ParseSubjectScore(cs.Score.Value);
                                //cca.AddSubjectScore(_courseDict[cs.CourseID].Subject, value);
                                cca.AddSubjectScore(_courseDict[cs.CourseID].Subject, cs.Score.Value);
                            }
                            //decimal? s = GetRoundScore(cs.Score);
                            //if (s.HasValue)
                            //{
                            //    cca.ClassID = stud.Class.ID;
                            //    cca.ClassName = stud.Class.Name;
                            //    cca.AddCourseScore(cs.CourseID, s.Value);
                            //}
                        }
                    }
                }

                // 班級名稱
                ws.Cells[dataRowIndex, 0].PutValue(cca.ClassName);
                // 放班級人數
                int peoColIdx = 32;
                // 找到人數放位置
                for (int i = 32; i <= 60; i++)
                    if (ws.Cells[2, i].StringValue.Trim() == "人數")
                    {
                        peoColIdx = i;
                        break;
                    }
                if (cca.ClassID != null && cca.ClassName != "")
                    if (ClassStudCount.ContainsKey(cca.ClassID))
                        ws.Cells[dataRowIndex, peoColIdx].PutValue(ClassStudCount[cca.ClassID].ToString());


                foreach (KeyValuePair<string, decimal?> val in cca.GetSubjectStudScoreAvg())
                {
                    if (val.Value.HasValue)
                    {
                        decimal value = Math.Round((decimal)val.Value, 2, MidpointRounding.AwayFromZero);
                        ws.Cells[dataRowIndex, colMapping[val.Key] + 1].PutValue(value);
                    }
                }
                //填入領域
                bool domainHasData = false;
                foreach (string domain in _domains)
                {
                    if (!domainScores.ContainsKey(domain)) continue;
                    if (!domainScores[domain].HasValue) continue;

                    string taggedDomain = GetTaggedDomain(domain);
                    if (colMapping.ContainsKey(taggedDomain))
                    {
                        decimal value = Math.Round((decimal)domainScores[domain].Value, 2, MidpointRounding.AwayFromZero);
                        ws.Cells[dataRowIndex, colMapping[taggedDomain] + 1].PutValue(value);
                        domainHasData = true;
                    }
                }
                if (cca.GetSubjectStudScoreAvg().Count > 0 || domainHasData)
                    dataRowIndex++;

                #endregion

                #region 填入標題及回條
                //Ex. 新竹市立光華國民中學 97 學年度第 1 學期    第1次平時評量成績單
                ws.Cells[title.FirstRow + rowIndex, title.FirstColumn].PutValue(string.Format("{0}  {1}  {2} 各班各科平均成績比較表", SchoolName, Semester, _exam.Name));
                //Ex. 101 第1次平時評量回條 (家長意見欄)
                //ws.Cells[feedback.FirstRow + rowIndex, feedback.FirstColumn].PutValue(string.Format("{0}  {1}回條  (家長意見欄)", ced.Class.Name, _exam.Name));
                #endregion

                //rowIndex += RowNumber;
                //ws.HPageBreaks.Add(rowIndex, 0);
            }

            e.Result = book;
        }

        private string GetDomainSubjectKey(string p, string p_2)
        {
            return p + "_" + p_2;
        }

        private string GetOnlyDomain(string p)
        {
            if (p.Contains("_")) return p.Split(new string[] { "_" }, StringSplitOptions.None)[0];
            else return p;
        }

        private string GetTaggedDomain(string p)
        {
            return "領域" + p;
        }

        private string GetOriginalDomain(string p)
        {
            if (p.StartsWith("領域")) return p.Replace("領域", "");
            else return p;
        }

        private bool IsDomain(string p)
        {
            if (p.StartsWith("領域")) return true;
            else return false;
        }

        /// <summary>
        /// 取得科目資料管理
        /// </summary>
        /// <returns></returns>
        private List<string> GetSubjectList()
        {

            QueryHelper queryHelper = new QueryHelper();
            DataTable dataTable;
            try
            {
                dataTable = queryHelper.Select(@"WITH    subject_mapping AS 
(
SELECT
    unnest(xpath('//Subjects/Subject/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS subject_name
FROM  
    list 
WHERE name  ='JHEvaluation_Subject_Ordinal'
)SELECT
		replace (subject_name ,'&amp;amp;','&') AS subject_name
	FROM  subject_mapping");
            }
            catch
            {
                throw new Exception("查詢科目對照失敗！");
            }

            //List<string> subjectNameList=new List<string>();

            foreach (DataRow dtRow in dataTable.Rows)
            {

                subjectNameList.Add("領域語文");
                subjectNameList.Add("領域數學");
                subjectNameList.Add("領域社會");
                subjectNameList.Add("領域自然科學");
                subjectNameList.Add("領域自然與生活科技");
                subjectNameList.Add("領域藝術");
                subjectNameList.Add("領域藝術與人文");
                subjectNameList.Add("領域健康與體育");
                subjectNameList.Add("領域綜合活動");
                subjectNameList.Add("領域科技");
                subjectNameList.Add("領域特殊需求");
                subjectNameList.Add(dtRow["subject_name"].ToString());
            }
            return subjectNameList;
        }

        private int Sort(string x, string y)
        {

            List<string> list = new List<string>(new string[] { "國語文", "國文", "英語文", "英文", "英語", "領域語文", "數學", "領域數學", "歷史", "公民", "地理", "領域社會", "理化", "生物", "地球科學", "領域自然與生活科技", "領域自然科學", "音樂", "視覺藝術", "表演藝術", "領域藝術與人文", "領域藝術", "健康教育", "體育", "領域健康與體育", "家政", "童軍", "輔導", "領域綜合活動", "資訊科技", "生活科技", "領域科技" });

            int ix = list.IndexOf(x);
            int iy = list.IndexOf(y);


            //科目資料管理排序
            //int ix = subjectNameList.IndexOf(x);
            //int iy = subjectNameList.IndexOf(y);


            if (ix >= 0 && iy >= 0)
                return ix.CompareTo(iy);
            else if (ix >= 0)
                return -1;
            else if (iy >= 0)
                return 1;
            else
                return x.CompareTo(y);
        }
        #endregion

        /// <summary>
        /// 產生報表
        /// </summary>
        internal void Generate()
        {
            if (!_worker.IsBusy)
                _worker.RunWorkerAsync();
        }

        //private decimal? GetRoundScore(decimal? score)
        //{
        //    if (!score.HasValue) return null;
        //    decimal seed = Convert.ToDecimal(Math.Pow(0.1, Convert.ToDouble(2)));
        //    decimal s = score.Value / seed;
        //    s = decimal.Floor(s);
        //    s *= seed;
        //    return s;
        //}
    }
}
