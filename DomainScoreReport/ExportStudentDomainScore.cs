using System;
using System.Xml;
using System.Data;
using System.Linq;
using System.IO;
using System.ComponentModel;
using System.Windows.Forms;
using System.Collections.Generic;
using FISCA.Data;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Data.Configuration;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace DomainScoreReport
{
    class ExportStudentDomainScore
    {
        private const string ConfigName = "JHEvaluation_Subject_Ordinal";
        private const string ColumnKey = "DomainOrdinal";
        private Document docTemplat;
        private DataTable dtRsp = new DataTable();
        private List<string> listStudentIDs = new List<string>();
        private List<string> listDomainName = new List<string>();
        private Dictionary<string, StudentRec> dicStudentByID = new Dictionary<string, StudentRec>();
        private QueryHelper qh = new QueryHelper();
        private BackgroundWorker bg = new BackgroundWorker();

        public ExportStudentDomainScore(List<string>listIDs)
        {
            listStudentIDs = listIDs;
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

            // 讀取樣板
            Stream stream = new MemoryStream(Properties.Resources.StudentDomainScore_template);
            docTemplat = new Document(stream);

            InitailizeBackGroundWorker();
        }

        private void InitailizeBackGroundWorker()
        {
            bg.WorkerReportsProgress = true;
            bg.DoWork += new DoWorkEventHandler(BackGroundWork_DoWork);
            bg.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackGroundWork_RunWorkerCompleted);
            bg.ProgressChanged += new ProgressChangedEventHandler(BackGroundWork_ProgressChanged);
        }

        private void BackGroundWork_DoWork(object sender, DoWorkEventArgs e)
        {
            int progress = 0;
            BackgroundWorker worker = sender as BackgroundWorker;

            Document doc = new Document();
            doc.RemoveAllChildren();

            // 取得學生領域成績
            bool isSuccess = GetStudentDomainScore();
            if (isSuccess)
            {
                worker.ReportProgress(progress += 10);
                if (dtRsp.Rows.Count > 0)
                {
                    // 資料整理
                    ParseData();
                    worker.ReportProgress(progress += 30);
                    try
                    {

                        DataTable dt = FillMergeFiledData();
                        int n = 70 / dt.Rows.Count;

                        foreach (DataRow row in dt.Rows)
                        {
                            Document eachDoc = new Document();
                            eachDoc.Sections.Clear();
                            eachDoc.Sections.Add(eachDoc.ImportNode(docTemplat.FirstSection, true));
                            eachDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
                            eachDoc.MailMerge.Execute(row);

                            doc.Sections.Add(doc.ImportNode(eachDoc.Sections[0], true));
                            worker.ReportProgress(progress += n);
                        }

                        e.Result = doc;

                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show("無法列印："+ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    worker.ReportProgress(progress += 70);
                    e.Result = null;
                }
            }
        }

        private void BackGroundWork_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("成績預警通知單 產生中:", e.ProgressPercentage);
        }

        private void BackGroundWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document doc = (Document)e.Result;

            if (doc != null)
            {
                string path = $"{Application.StartupPath}\\Reports\\成績預警通知單.docx";
                int i = 1;
                while (File.Exists(path))
                {
                    string docName = Path.GetFileNameWithoutExtension(path);

                    string newPath = $"{Path.GetDirectoryName(path)}\\成績預警通知單{i++}{Path.GetExtension(path)}";

                    path = newPath;
                }

                doc.Save(path, SaveFormat.Docx);
                MotherForm.SetStatusBarMessage("成績預警通知單 產生完成");

                DialogResult result = MsgBox.Show($"{path}\n成績預警通知單產生完成，是否立刻開啟？", "訊息", MessageBoxButtons.YesNo);

                if (DialogResult.Yes == result)
                {
                    System.Diagnostics.Process.Start(path);
                }
            }
            else
            {
                MotherForm.SetStatusBarMessage("成績資料有誤，無法產生成績預警通知單。");
            }
        }

        /// <summary>
        /// 成績單列印
        /// </summary>
        public void Export()
        {
            if (!bg.IsBusy)
            {
                bg.RunWorkerAsync();
            }
        }

        /// <summary>
        /// 取得學生領域成績資料
        /// </summary>
        private bool GetStudentDomainScore()
        {
            string sql = string.Format(@"
WITH target_student AS(
	SELECT
		*
	FROM
		student
	WHERE
		id IN({0})
)
SELECT
    student.id
    , student.name
    , student.seat_no
    , class.class_name
	, sems_subj_score_ext.semester
	, sems_subj_score_ext.school_year
	, array_to_string(xpath('/Domain/@原始成績', subj_score_ele), '')::text AS 原始成績
	, array_to_string(xpath('/Domain/@成績', subj_score_ele), '')::text AS 成績
	, array_to_string(xpath('/Domain/@領域', subj_score_ele), '')::text AS 領域
    , array_to_string(xpath('/Domain/@權數', subj_score_ele), '')::text AS 權數
FROM (
		SELECT 
			sems_subj_score.*
			, unnest(xpath('/root/Domains/Domain', xmlparse(content '<root>' || score_info || '</root>'))) as subj_score_ele
		FROM 
			sems_subj_score 
			INNER JOIN target_student
				ON target_student.id = sems_subj_score.ref_student_id 
	) as sems_subj_score_ext
    LEFT OUTER JOIN student
        ON student.id = sems_subj_score_ext.ref_student_id
    LEFT OUTER JOIN class
        ON class.id = student.ref_class_id
ORDER BY
	sems_subj_score_ext.grade_year
	, school_year
	, semester
    , class.grade_year
    , class.display_order
    , class.class_name
    , student.seat_no
    , student.student_number
    , student.id
            ", string.Join(",", listStudentIDs));

            try
            {
                dtRsp = qh.Select(sql);
                return true;
            }
            catch (Exception error)
            {
                MsgBox.Show("成績單列印失敗：" + error.Message);
                return false;
            }
        }

        /// <summary>
        /// 資料解析
        /// </summary>
        private void ParseData()
        {
            foreach (DataRow row in dtRsp.Rows)
            {
                string stuID = "" + row["id"];

                // 新增學生物件
                if (!dicStudentByID.ContainsKey(stuID))
                {
                    StudentRec stuRec = new StudentRec();
                    stuRec.ID = stuID;
                    stuRec.Name = "" + row["name"];
                    stuRec.SeatNo = "" + row["seat_no"];
                    stuRec.ClassName = "" + row["class_name"];

                    dicStudentByID.Add(stuID, stuRec);
                }
                // 資料整理
                {
                    StudentRec stuRec = dicStudentByID[stuID];

                    string schoolYear = "" + row["school_year"];
                    string semester = "" + row["semester"];
                    string domain = "" + row["領域"];
                    string ssKey = schoolYear + semester;

                    // 沒有領域就忽略
                    //if (domain != "")
                    if (listDomainName.Contains(domain) && domain != "特殊需求")  //如果領域存在領域資料管理，且不是特殊需求才列印，沒有則忽略-Cynthia 2021.08
                    {
                        // 成績
                        if (!stuRec.dicScoreByDomainBySchoolYear.ContainsKey(ssKey))
                        {
                            stuRec.dicScoreByDomainBySchoolYear.Add(ssKey, new Dictionary<string, ScoreRec>());
                        }
                        ScoreRec sr = new ScoreRec();
                        sr.Score = "" + row["成績"];
                        sr.Power = "" + row["權數"];
                        sr.OriginScore = "" + row["原始成績"];

                        stuRec.dicScoreByDomainBySchoolYear[ssKey].Add(domain, sr);
                        // 領域
                        if (!stuRec.listDomainFromStu.Contains(domain))
                        {
                            stuRec.listDomainFromStu.Add(domain);
                        }
                        // 學年度學期
                        if (!stuRec.dicSchoolYear.Keys.Contains(ssKey))
                        {
                            SchoolYearSemester ss = new SchoolYearSemester();
                            ss.SchoolYear = schoolYear;
                            ss.Semester = semester;
                            stuRec.dicSchoolYear.Add(ssKey, ss);
                        }
                    }
                }
            }

            // 領域根據對照表做排序
            foreach (string id in dicStudentByID.Keys)
            {
                dicStudentByID[id].listDomainFromStu.Sort(delegate (string a, string b) {
                    int aIndex = listDomainName.FindIndex(name => name == a);
                    int bIndex = listDomainName.FindIndex(name => name == b);

                    if (aIndex > bIndex)
                    {
                        return 1;
                    }
                    else
                    {
                        return -1;
                    }
                });
            }
        }

        /// <summary>
        /// 將資料解析為 doc merge 格式
        /// </summary>
        private DataTable FillMergeFiledData()
        {
            string schoolName = School.ChineseName;
            DataTable dt = CreateMergeFieldTable();
            
            foreach (string id in dicStudentByID.Keys)
            {
                DataRow row = dt.NewRow();
                row["year"] = DateTime.Now.Year - 1911;
                row["month"] = DateTime.Now.Month;
                row["day"] = DateTime.Now.Day;
                row["school_name"] = schoolName;
                row["class_name"] = dicStudentByID[id].ClassName;
                row["seat_no"] = dicStudentByID[id].SeatNo;
                row["student_name"] = dicStudentByID[id].Name;

                // 領域及格數
                int passCount = 0;
                // 領域加權平均清單
                List<decimal> listDomainAvgScore = new List<decimal>();
                // 領域
                int d = 1;
                foreach (string domain in dicStudentByID[id].listDomainFromStu)
                {
                    // 領域名稱
                    row[$"{d}_domain"] = domain;
                    // 領域個學期分數與權重清單
                    List<ScoreRec> listScore = new List<ScoreRec>();

                    // 學年度學期
                    int s = 1;
                    foreach (string key in dicStudentByID[id].dicSchoolYear.Keys)
                    {
                        SchoolYearSemester ss = dicStudentByID[id].dicSchoolYear[key];
                        if (d == 1)
                        {
                            row[$"{s}_school_year"] = $"{ss.SchoolYear}學年度";
                            row[$"{s}_semester"] = $"第{ss.Semester}學期";
                        }
                        // 成績
                        if (dicStudentByID[id].dicScoreByDomainBySchoolYear.ContainsKey(key))
                        {
                            if (dicStudentByID[id].dicScoreByDomainBySchoolYear[key].ContainsKey(domain))
                            {
                                string score = dicStudentByID[id].dicScoreByDomainBySchoolYear[key][domain].Score;
                                string originScore = dicStudentByID[id].dicScoreByDomainBySchoolYear[key][domain].OriginScore;
                                string power = dicStudentByID[id].dicScoreByDomainBySchoolYear[key][domain].Power;
                                row[$"{d}_domain_{s}_score"] = score == originScore ? score : $"*{score}";
                                row[$"{d}_d_{s}_p"] = power;

                                ScoreRec sr = new ScoreRec();
                                sr.Score = score;
                                sr.Power = power;

                                listScore.Add(sr);
                            }
                        }
                        s++;
                    }

                    // 領域平均
                    if (listScore.Count > 0)
                    {
                        decimal totalScore = 0;
                        int domainScoreCount = 0;
                        foreach (ScoreRec sr in listScore)
                        {
                            totalScore += DecimalParser(sr.Score);
                            domainScoreCount++;
                        }
                        if (totalScore != 0)
                        {
                            decimal avgScore = Math.Round(totalScore / domainScoreCount, 2, MidpointRounding.AwayFromZero);
                            row[$"{d}_domain_avg"] = avgScore;
                            listDomainAvgScore.Add(avgScore);

                            if (avgScore >= 60)
                            {
                                passCount++;
                            }
                        }
                    }

                    // 及格數
                    row["pass_domain_count"] = passCount;
                    d++;
                }

                // 學期各領域平均
                // 個學期領域平均的算術平均
                decimal allTotalScore = 0;
                int scoreCount = 0;
                int sIndex = 1;
                foreach (string key in dicStudentByID[id].dicSchoolYear.Keys)
                {
                    // 學年度學期各領域成績資料
                    List<ScoreRec> listScore = new List<ScoreRec>();
                    foreach (string domain in dicStudentByID[id].listDomainFromStu)
                    {
                        if (dicStudentByID[id].dicScoreByDomainBySchoolYear.ContainsKey(key))
                        {
                            if (dicStudentByID[id].dicScoreByDomainBySchoolYear[key].ContainsKey(domain))
                            {
                                string score = dicStudentByID[id].dicScoreByDomainBySchoolYear[key][domain].Score;
                                string power = dicStudentByID[id].dicScoreByDomainBySchoolYear[key][domain].Power;

                                ScoreRec sr = new ScoreRec();
                                sr.Score = score;
                                sr.Power = power;

                                listScore.Add(sr);
                            }
                        }
                    }

                    if (listScore.Count > 0)
                    {
                        decimal totalScore = 0;
                        decimal totalPower = 0;
                        foreach (ScoreRec sr in listScore)
                        {
                            totalScore += DecimalParser(sr.Score) * DecimalParser(sr.Power);
                            totalPower += DecimalParser(sr.Power);
                        }
                        // 沒有權重就不幫你算
                        if (totalPower != 0)
                        {
                            decimal score = Math.Round(totalScore / totalPower, 2, MidpointRounding.AwayFromZero);
                            row[$"{sIndex}_all_domain_avg"] = score;
                            allTotalScore += score;
                            scoreCount++;
                        }
                    }

                    sIndex++;
                }

                // 總平均(算術平均)
                if (listDomainAvgScore.Count > 0)
                {
                    //foreach (double score in listDomainAvgScore)
                    //{
                    //    totalScore += score;
                    //}
                    decimal avgScore = Math.Round(allTotalScore / scoreCount, 2, MidpointRounding.AwayFromZero);

                    row["all_domain_avg"] = avgScore;
                }

                dt.Rows.Add(row);
            }

            return dt;
        }

        private DataTable CreateMergeFieldTable()
        {
            DataTable table = new DataTable();
            DataColumn col;

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "year";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "month";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "day";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "school_name";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "class_name";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "seat_no";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "student_name";
            table.Columns.Add(col);

            for (int i = 1; i <= 6; i++)
            {
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"{i}_school_year";
                table.Columns.Add(col);
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"{i}_semester";
                table.Columns.Add(col);

                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"{i}_all_domain_avg";
                table.Columns.Add(col);
            }

            for (int i = 1; i <= 10; i++)
            {
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"{i}_domain";
                table.Columns.Add(col);

                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"{i}_domain_avg";
                table.Columns.Add(col);
                for (int n = 1; n <= 6; n++)
                {
                    col = new DataColumn();
                    col.DataType = Type.GetType("System.String");
                    col.ColumnName = $"{i}_domain_{n}_score";
                    table.Columns.Add(col);

                    col = new DataColumn();
                    col.DataType = Type.GetType("System.String");
                    col.ColumnName = $"{i}_d_{n}_p";
                    table.Columns.Add(col);
                }
            }

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = $"all_domain_avg";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "pass_domain_count";
            table.Columns.Add(col);

            return table;
        }

        private decimal DecimalParser(string data)
        {
            return decimal.Parse(data == "" ? "0" : data);
        }

        private float FloatParser2(string data)
        {
            return float.Parse(data == "" ? "0" : data);
        }

        private class StudentRec
        {
            public string ID;
            public string ClassName;
            public string SeatNo;
            public string Name;
            public List<string> listDomainFromStu = new List<string>();
            public Dictionary<string, SchoolYearSemester> dicSchoolYear = new Dictionary<string, SchoolYearSemester>();
            public Dictionary<string, Dictionary<string, ScoreRec>> dicScoreByDomainBySchoolYear = new Dictionary<string, Dictionary<string, ScoreRec>>();
        }

        private class SchoolYearSemester
        {
            public string SchoolYear;
            public string Semester;
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
