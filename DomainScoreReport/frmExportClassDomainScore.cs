using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using FISCA.Data;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Data.Configuration;
using Aspose.Words;
using Aspose.Words.Reporting;
using JHSchool.Data;
using JHSchool.Evaluation.Calculation;
using System.Linq;

namespace DomainScoreReport
{
    public partial class frmExportClassDomainScore : BaseForm
    {
        private const string ConfigName = "JHEvaluation_Subject_Ordinal";
        private const string ColumnKey = "DomainOrdinal";
        private List<string> listDomainName = new List<string>();
        private Document docTemplate;
        private List<string> listClassID = new List<string>();
        private QueryHelper qh = new QueryHelper();
        private Dictionary<string, Dictionary<string, Dictionary<string, ScoreRec>>> dicClassStuDomainScore = new Dictionary<string, Dictionary<string, Dictionary<string, ScoreRec>>>();
        private Dictionary<string, string> dicClassNameByID = new Dictionary<string, string>();
        private Dictionary<string, StudentRec> dicStuRecByID = new Dictionary<string, StudentRec>();
        private Dictionary<string, List<string>> dicDomainNameByClassID = new Dictionary<string, List<string>>();
        private BackgroundWorker bgWorker = new BackgroundWorker();
        Dictionary<string, ScoreCalculator> calcCache = new Dictionary<string, ScoreCalculator>();
        Dictionary<string, string> calcIDCache = new Dictionary<string, string>();


        // 各學期領域總分
        Dictionary<string, Dictionary<string, decimal>> StudentSemsDomainScoreSumDict = new Dictionary<string, Dictionary<string, decimal>>();

        // 各學期學分加總
        Dictionary<string, Dictionary<string, decimal>> StudentSemsCreditSumDict = new Dictionary<string, Dictionary<string, decimal>>();

        public frmExportClassDomainScore(List<string> listIDs)
        {
            InitializeComponent();
            InitializeBackgroundWorker();
            listClassID = listIDs;
        }

        private void InitializeBackgroundWorker()
        {
            bgWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);
            bgWorker.WorkerReportsProgress = true;
        }

        private void frmExportClassDomainScore_Load(object sender, EventArgs e)
        {


            // 讀取樣板
            Stream stream = new MemoryStream(Properties.Resources.ClassDomainScore_template);
            docTemplate = new Document(stream);

            // 取得領域對照表
            ConfigData cd = School.Configuration[ConfigName];
            {
                if (cd.Contains(ColumnKey))
                {
                    XmlElement element = cd.GetXml(ColumnKey, XmlHelper.LoadXml("<Domains/>"));
                    foreach (XmlElement domainElement in element.SelectNodes("Domain"))
                    {
                        string name = domainElement.GetAttribute("Name");
                        listDomainName.Add(name);
                    }
                }
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            btnExport.Enabled = false;
            if (!bgWorker.IsBusy)
            {
                ParameterRec data = new ParameterRec();

                data.ClassIDs = string.Join(",", listClassID);

                bgWorker.RunWorkerAsync(data);
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            ParameterRec data = (ParameterRec)e.Argument;

            Document doc = new Document();
            doc.RemoveAllChildren();
            int progress = 0;
            // 取得學年度學期班級學生領域成績
            DataTable dt = GetClassStudentDomainScore(data);
            bgWorker.ReportProgress(progress += 10);
            // 資料解析
            ParseData(dt);
            bgWorker.ReportProgress(progress += 10);
            // 資料填寫
            DataTable table = FillMergeFiledData(data);
            bgWorker.ReportProgress(progress += 10);

            int p = 0;
            if (table.Rows.Count > 0)
            {
                p = 70 / table.Rows.Count;
            }
            foreach (DataRow row in table.Rows)
            {
                Document eachDoc = new Document();
                eachDoc.Sections.Clear();
                Section section = (Section)eachDoc.ImportNode(docTemplate.FirstSection, true);
                eachDoc.AppendChild(section);
                eachDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
                eachDoc.MailMerge.Execute(row);

                doc.Sections.Add(doc.ImportNode(eachDoc.Sections[0], true));
                bgWorker.ReportProgress(progress += p);
            }

            string path = $"{Application.StartupPath}\\Reports\\班級成績預警通知單.docx";
            int i = 1;
            while (File.Exists(path))
            {
                string docName = Path.GetFileNameWithoutExtension(path);
                string newPath = $"{Path.GetDirectoryName(path)}\\成績預警通知單{i++}{Path.GetExtension(path)}";

                path = newPath;
            }

            doc.Save(path, SaveFormat.Docx);
            e.Result = path;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnExport.Enabled = true;
            string path = (string)e.Result;

            MotherForm.SetStatusBarMessage("班級成績預警通知單 產生完成");
            DialogResult result = MsgBox.Show($"{path}\n班級成績預警通知單產生完成，是否立刻開啟？", "訊息", MessageBoxButtons.YesNo);

            if (DialogResult.Yes == result)
            {
                System.Diagnostics.Process.Start(path);
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("班級成績預警通知單 產生中:", e.ProgressPercentage);
        }

        private DataTable GetClassStudentDomainScore(ParameterRec data)
        {
            //            string sql = string.Format(@"
            //WITH data_row AS (
            //    SELECT
            //        {0}::INT AS school_year
            //        , {1}::INT AS semester
            //), target_student AS(
            //	SELECT
            //		*
            //	FROM
            //		student
            //	WHERE
            //		ref_class_id IN ({2})
            //        AND status IN(1, 2)
            //)
            //SELECT
            //    student.id
            //    , student.name
            //    , student.seat_no
            //    , class.id AS class_id
            //    , class.class_name
            //	, sems_subj_score_ext.semester
            //	, sems_subj_score_ext.school_year
            //	, array_to_string(xpath('/Domain/@原始成績', subj_score_ele), '')::text AS 原始成績
            //	, array_to_string(xpath('/Domain/@成績', subj_score_ele), '')::text AS 成績
            //	, array_to_string(xpath('/Domain/@領域', subj_score_ele), '')::text AS 領域
            //    , array_to_string(xpath('/Domain/@權數', subj_score_ele), '')::text AS 權數
            //FROM (
            //		SELECT 
            //			sems_subj_score.*
            //			, unnest(xpath('/root/Domains/Domain', xmlparse(content '<root>' || score_info || '</root>'))) as subj_score_ele
            //		FROM 
            //			sems_subj_score 
            //			INNER JOIN target_student
            //				ON target_student.id = sems_subj_score.ref_student_id 
            //	) as sems_subj_score_ext
            //    LEFT OUTER JOIN student
            //        ON student.id = sems_subj_score_ext.ref_student_id
            //    LEFT OUTER JOIN class
            //        ON class.id = student.ref_class_id
            //    INNER JOIN data_row
            //    	ON data_row.school_year = sems_subj_score_ext.school_year
            //    	AND data_row.semester = sems_subj_score_ext.semester
            //ORDER BY
            //	sems_subj_score_ext.grade_year 
            //    , class.display_order
            //    , student.seat_no
            //                ", data.SchoolYear, data.Semester, data.ClassIDs);

            //return qh.Select(sql);

            // 取得班級學生含研修
            QueryHelper qh1 = new QueryHelper();
            string qry = @"SELECT student.id,student.name,student.seat_no,class.id AS class_id,class.class_name FROM student INNER JOIN class on student.ref_class_id = class.id WHERE student.status IN(1,2) AND class.id IN(" + data.ClassIDs + @") ORDER BY class.grade_year,class.display_order,class.class_name,student.seat_no";

            DataTable dtStudent = qh1.Select(qry);

            List<string> StudentIDList = new List<string>();

            foreach (DataRow dr in dtStudent.Rows)
            {
                StudentIDList.Add(dr["id"].ToString());
            }

            // 取得學生學期成績，即時計算使用
            List<JHSemesterScoreRecord> StudentSemesterScoreList = JHSemesterScore.SelectByStudentIDs(StudentIDList);

            // 建立成績資料索引
            Dictionary<string, List<JHSemesterScoreRecord>> StudentSemesterScoreDict = new Dictionary<string, List<JHSemesterScoreRecord>>();

            foreach (JHSemesterScoreRecord rec in StudentSemesterScoreList)
            {
                if (!StudentSemesterScoreDict.ContainsKey(rec.RefStudentID))
                    StudentSemesterScoreDict.Add(rec.RefStudentID, new List<JHSemesterScoreRecord>());

                StudentSemesterScoreDict[rec.RefStudentID].Add(rec);
            }


            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");
            dt.Columns.Add("seat_no");
            dt.Columns.Add("class_id");
            dt.Columns.Add("class_name");
            dt.Columns.Add("semester");
            dt.Columns.Add("school_year");
            dt.Columns.Add("原始成績");
            dt.Columns.Add("成績");
            dt.Columns.Add("領域");
            dt.Columns.Add("權數");

            // 各領域學期成績
            Dictionary<string, List<decimal>> domainnScoreDict = new Dictionary<string, List<decimal>>();



            // 各領域學期成績原始
            Dictionary<string, List<decimal>> domainnScoreOrignDict = new Dictionary<string, List<decimal>>();

            // 各領域成績(算術平均)
            Dictionary<string, decimal> domainnScoreAvgDict = new Dictionary<string, decimal>();
            // 各領域成績(算術平均)(原始)
            Dictionary<string, decimal> domainnScoreAvgOrignDict = new Dictionary<string, decimal>();

            // 各領域學分
            Dictionary<string, List<decimal>> domainCreditDict = new Dictionary<string, List<decimal>>();

            // 各領域學分平均
            Dictionary<string, decimal> domainAvgCreditDict = new Dictionary<string, decimal>();

            #region 取得學生成績計算規則
            ScoreCalculator defaultScoreCalculator = new ScoreCalculator(null);

            //key: ScoreCalcRuleID
            calcCache.Clear();

            //key: StudentID, val: ScoreCalcRuleID
            calcIDCache.Clear();

            List<string> scoreCalcRuleIDList = new List<string>();

            List<StudentRecord> StudentRecList = K12.Data.Student.SelectByIDs(StudentIDList);
            foreach (StudentRecord student in StudentRecList)
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


            #endregion

            // 計算畢業成績使用
            StudentSemsDomainScoreSumDict.Clear();
            StudentSemsCreditSumDict.Clear();

            // 開始填資料
            foreach (DataRow dr in dtStudent.Rows)
            {
                string sid = dr["id"].ToString();

                //DataRow newRow = dt.NewRow();             
                //newRow["id"] = sid;
                //newRow["name"] = dr["name"].ToString();
                //newRow["seat_no"] = dr["seat_no"].ToString();
                //newRow["class_id"] = dr["class_id"].ToString();
                //newRow["class_name"] = dr["class_name"].ToString();
                //newRow["semester"] = "";
                //newRow["school_year"] = "";

                //newRow["原始成績"] = "0";
                //newRow["成績"] = "0";
                //newRow["領域"] = "0";
                //newRow["權數"] = "0";

                // 即時計算領域成績
                if (StudentSemesterScoreDict.ContainsKey(sid))
                {

                    // 成績計算規則
                    ScoreCalculator studentCalculator = defaultScoreCalculator;
                    if (calcIDCache.ContainsKey(sid) && calcCache.ContainsKey(calcIDCache[sid]))
                        studentCalculator = calcCache[calcIDCache[sid]];

                    domainnScoreDict.Clear();

                    domainnScoreOrignDict.Clear();
                    domainnScoreAvgDict.Clear();
                    domainnScoreAvgOrignDict.Clear();
                    domainCreditDict.Clear();

                    if (!StudentSemsDomainScoreSumDict.ContainsKey(sid))
                        StudentSemsDomainScoreSumDict.Add(sid, new Dictionary<string, decimal>());

                    if (!StudentSemsCreditSumDict.ContainsKey(sid))
                        StudentSemsCreditSumDict.Add(sid, new Dictionary<string, decimal>());

                    // 整理各學期領域成績
                    foreach (JHSemesterScoreRecord rec in StudentSemesterScoreDict[sid])
                    {
                        string key = rec.SchoolYear + "_" + rec.Semester;

                        if (!StudentSemsDomainScoreSumDict[sid].ContainsKey(key))
                            StudentSemsDomainScoreSumDict[sid].Add(key, 0);

                        if (!StudentSemsCreditSumDict[sid].ContainsKey(key))
                            StudentSemsCreditSumDict[sid].Add(key, 0);

                        // 讀取成績
                        foreach (string dName in rec.Domains.Keys)
                        {
                            // 加入領域成績
                            if (rec.Domains[dName].Score.HasValue)
                            {
                                if (!domainnScoreDict.ContainsKey(dName))
                                {
                                    domainnScoreDict.Add(dName, new List<decimal>());
                                }
                                domainnScoreDict[dName].Add(rec.Domains[dName].Score.Value);

                                if (rec.Domains[dName].Credit.HasValue)
                                {
                                    StudentSemsDomainScoreSumDict[sid][key] += (rec.Domains[dName].Score.Value * rec.Domains[dName].Credit.Value);
                                }
                            }

                            // 加入領域原始
                            if (rec.Domains[dName].ScoreMakeup.HasValue)
                            {
                                if (!domainnScoreOrignDict.ContainsKey(dName))
                                {
                                    domainnScoreOrignDict.Add(dName, new List<decimal>());
                                }
                                domainnScoreOrignDict[dName].Add(rec.Domains[dName].ScoreMakeup.Value);

                            }

                            // 加入權數
                            if (rec.Domains[dName].Credit.HasValue)
                            {
                                if (!domainCreditDict.ContainsKey(dName))
                                    domainCreditDict.Add(dName, new List<decimal>());

                                domainCreditDict[dName].Add(rec.Domains[dName].Credit.Value);

                                StudentSemsCreditSumDict[sid][key] += rec.Domains[dName].Credit.Value;
                            }
                        }
                    }

                    // 計算成績
                    foreach (string dName in domainnScoreDict.Keys)
                    {
                        // 使用畢業成績計算規則四捨五入方式，即時計算平均
                        decimal avgScore = studentCalculator.ParseGraduateScore(domainnScoreDict[dName].Average());
                        if (!domainnScoreAvgDict.ContainsKey(dName))
                            domainnScoreAvgDict.Add(dName, avgScore);
                    }

                    foreach (string dName in domainnScoreOrignDict.Keys)
                    {
                        // 使用畢業成績計算規則四捨五入方式，即時計算平均(原始)
                        decimal avgScore = studentCalculator.ParseGraduateScore(domainnScoreOrignDict[dName].Average());
                        if (!domainnScoreAvgOrignDict.ContainsKey(dName))
                            domainnScoreAvgOrignDict.Add(dName, avgScore);
                    }

                    // 處理學分

                    foreach (string dName in domainCreditDict.Keys)
                    {
                        // 使用畢業成績計算規則四捨五入方式，即時計算平均(原始)
                        decimal avgScore = studentCalculator.ParseGraduateScore(domainCreditDict[dName].Average());
                        if (!domainAvgCreditDict.ContainsKey(dName))
                            domainAvgCreditDict.Add(dName, avgScore);
                    }

                    foreach (string domainName in domainnScoreAvgDict.Keys)
                    {
                        // 填資料
                        DataRow newRow = dt.NewRow();
                        newRow["id"] = sid;
                        newRow["name"] = dr["name"].ToString();
                        newRow["seat_no"] = dr["seat_no"].ToString();
                        newRow["class_id"] = dr["class_id"].ToString();
                        newRow["class_name"] = dr["class_name"].ToString();
                        newRow["semester"] = "";
                        newRow["school_year"] = "";

                        newRow["原始成績"] = "";
                        if (domainnScoreAvgOrignDict.ContainsKey(domainName))
                        {
                            newRow["原始成績"] = domainnScoreAvgOrignDict[domainName];
                        }

                        newRow["成績"] = domainnScoreAvgDict[domainName];
                        newRow["領域"] = domainName;
                        newRow["權數"] = "";
                        if (domainAvgCreditDict.ContainsKey(domainName))
                        {
                            newRow["權數"] = domainAvgCreditDict[domainName];
                        }

                        dt.Rows.Add(newRow);
                    }


                }

            }

            return dt;



        }

        private void ParseData(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                string classID = "" + row["class_id"];
                string studentID = "" + row["id"];
                string domain = "" + row["領域"];

                if (domain != "")
                {
                    // 成績資料
                    if (!dicClassStuDomainScore.ContainsKey(classID))
                    {
                        dicClassStuDomainScore.Add(classID, new Dictionary<string, Dictionary<string, ScoreRec>>());
                    }
                    if (!dicClassStuDomainScore[classID].ContainsKey(studentID))
                    {
                        dicClassStuDomainScore[classID].Add(studentID, new Dictionary<string, ScoreRec>());
                    }

                    ScoreRec sr = new ScoreRec();
                    sr.Score = "" + row["成績"];
                    sr.OriginScore = "" + row["原始成績"];
                    sr.Power = "" + row["權數"];

                    dicClassStuDomainScore[classID][studentID].Add(domain, sr);
                    // 班級資料
                    if (!dicClassNameByID.ContainsKey(classID))
                    {
                        dicClassNameByID.Add(classID, "" + row["class_name"]);
                    }
                    // 學生資料
                    if (!dicStuRecByID.ContainsKey(studentID))
                    {
                        StudentRec stuRec = new StudentRec();
                        stuRec.SeatNo = "" + row["seat_no"];
                        stuRec.Name = "" + row["name"];

                        dicStuRecByID.Add(studentID, stuRec);
                    }
                    // 班級領域資料
                    if (!dicDomainNameByClassID.ContainsKey(classID))
                    {
                        dicDomainNameByClassID.Add(classID, new List<string>());
                    }
                    if (!dicDomainNameByClassID[classID].Contains(domain))
                    {
                        dicDomainNameByClassID[classID].Add(domain);
                    }
                }
            }
        }

        private DataTable CreateMergeFiledTable()
        {
            DataTable table = new DataTable();
            DataColumn col;

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "school_name";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "school_year";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "semester";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "class_name";
            table.Columns.Add(col);

            for (int i = 1; i <= 50; i++)    ///cyn 2021/4/27 由40改到50
            {
                // 座號
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"seat_no_{i}";
                table.Columns.Add(col);

                // 姓名
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"name_{i}";
                table.Columns.Add(col);

                // 領域成績
                for (int d = 1; d <= 8; d++)
                {
                    col = new DataColumn();
                    col.DataType = Type.GetType("System.String");
                    col.ColumnName = $"stu_{i}_domain_{d}";
                    table.Columns.Add(col);
                }

                // 平均成績
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"stu_{i}_avg_score";
                table.Columns.Add(col);

                // 七大領域及格數
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"stu_{i}_pass_count";
                table.Columns.Add(col);

            }

            for (int i = 1; i <= 8; i++)
            {
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"domain_{i}";
                table.Columns.Add(col);
            }

            return table;
        }

        private DataTable FillMergeFiledData(ParameterRec data)
        {
            DataTable dt = CreateMergeFiledTable();

            foreach (string classID in dicClassNameByID.Keys)
            {
                DataRow row = dt.NewRow();

                row["school_name"] = School.ChineseName;
                row["class_name"] = dicClassNameByID[classID];

                // 班級領域清單
                List<string> listDomain = dicDomainNameByClassID[classID];
                // 根據領域對照表做排序
                listDomain.Sort(delegate (string a, string b)
                {
                    int aIndex = listDomainName.FindIndex(domain => domain == a);
                    int bIndex = listDomainName.FindIndex(domain => domain == b);

                    if (aIndex > bIndex)
                    {
                        return 1;
                    }
                    else if (aIndex == bIndex)
                    {
                        return 0;
                    }
                    else
                    {
                        return -1;
                    }
                });
                // 目前樣板最多支援8個領域 多的不顯示
                if (listDomain.Count > 8)
                {
                    listDomain.RemoveRange(8, listDomain.Count - 8);
                }

                // 領域
                {
                    int d = 1;
                    foreach (string domain in listDomain)
                    {
                        row[$"domain_{d}"] = domain;

                        d++;
                    }
                }

                // 學生
                Dictionary<string, Dictionary<string, ScoreRec>> dicStuDomainScore = dicClassStuDomainScore[classID];
                int s = 1;

                foreach (string stuID in dicStuDomainScore.Keys)
                {
                    StudentRec stuRec = dicStuRecByID[stuID];
                    row[$"seat_no_{s}"] = stuRec.SeatNo;
                    row[$"name_{s}"] = stuRec.Name;

                    int d = 1;
                    int sc = 0;
                    int passCount = 0;
                    decimal totalPower = 0;
                    decimal totalScore = 0;
                    // 領域成績
                    foreach (string domain in listDomain)
                    {
                        if (dicStuDomainScore[stuID].ContainsKey(domain))
                        {

                            string score = dicStuDomainScore[stuID][domain].Score;
                            string originScore = dicStuDomainScore[stuID][domain].OriginScore;
                            string power = dicStuDomainScore[stuID][domain].Power;
                            row[$"stu_{s}_domain_{d}"] = score == originScore ? score : $"{score}";


                            //totalScore += FloatParser(score) * FloatParser(power);
                            //totalPower += FloatParser(power);

                            // 最後平均使用算術平均不使用加權平均
                            totalScore += FloatParser(score);
                            totalPower += 1;

                            if (FloatParser(score) >= 60)
                            {
                                passCount++;
                            }
                            sc++;
                        }
                        d++;
                    }




                    if (sc > 0)
                    {
                        // 沒有權重就不幫你算
                        if (totalPower > 0)
                        {

                            // 成績計算規則
                            ScoreCalculator defaultScoreCalculator = new ScoreCalculator(null);
                            ScoreCalculator studentCalculator = defaultScoreCalculator;
                            if (calcIDCache.ContainsKey(stuID) && calcCache.ContainsKey(calcIDCache[stuID]))
                                studentCalculator = calcCache[calcIDCache[stuID]];

                            //// 平均成績
                            //row[$"stu_{s}_avg_score"] = studentCalculator.ParseGraduateScore(totalScore / totalPower);   // Math.Round(totalScore / totalPower, 2);
                            List<decimal> scoreList = new List<decimal>();
                            if (StudentSemsDomainScoreSumDict.ContainsKey(stuID) && StudentSemsCreditSumDict.ContainsKey(stuID))
                            {                               
                                foreach(string sms in StudentSemsDomainScoreSumDict[stuID].Keys)
                                {
                                    // 各學期做加權平均
                                    if (StudentSemsCreditSumDict[stuID].ContainsKey(sms))
                                    {
                                        decimal ss = StudentSemsDomainScoreSumDict[stuID][sms];
                                        decimal cc = StudentSemsCreditSumDict[stuID][sms];
                                        if (cc > 0)
                                        {
                                            // 使用領域計算規則進位
                                            scoreList.Add(studentCalculator.ParseDomainScore(ss / cc));
                                        }

                                    }
                                }
                            }

                            // 最後再做算術平均在四捨五入
                            row[$"stu_{s}_avg_score"] = studentCalculator.ParseGraduateScore(scoreList.Average());   // Math.Round(totalScore / totalPower, 2);

                        }
                        // 領域及格數
                        row[$"stu_{s}_pass_count"] = passCount;
                    }
                    s++;
                }

                dt.Rows.Add(row);
            }

            return dt;
        }

        private decimal FloatParser(string data)
        {
            return decimal.Parse(data == "" ? "0" : data);
        }

        private void btnLeave_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private class StudentRec
        {
            public string SeatNo;
            public string Name;
        }

        private class ParameterRec
        {
            //public string SchoolYear;
            //public string Semester;
            public string ClassIDs;
        }

        private class ScoreRec
        {
            /// <summary>
            /// 成績
            /// </summary>
            public string Score;
            /// <summary>
            /// 原始成績
            /// </summary>
            public string OriginScore;
            /// <summary>
            /// 權數 
            /// </summary>
            public string Power;


        }
    }
}
