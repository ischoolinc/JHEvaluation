using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using K12.Data;
using FISCA.Data;
using Aspose.Words;
using K12.Data.Configuration;
using System.Xml;
using Aspose.Words.Reporting;
using FISCA.Presentation;

namespace DomainScoreReport
{
    public partial class frmExportSchoolDomainScore : BaseForm
    {
        private const string ConfigName = "JHEvaluation_Subject_Ordinal";
        private const string ColumnKey = "DomainOrdinal";
        private QueryHelper qh = new QueryHelper();
        private Document docTemplate;
        /// <summary>
        /// 領域對照表
        /// </summary>
        private List<string> listDomain = new List<string>();
        /// <summary>
        /// 成績資料的領域清單
        /// 根據領域對照表排序
        /// 取前7個
        /// </summary>
        private List<string> listDomainFromData = new List<string>();
        private Dictionary<string, StudentRec> dicStuRecByID = new Dictionary<string, StudentRec>();
        /// <summary>
        /// 領域不及格數
        /// </summary>
        private Dictionary<string, Dictionary<string, int>> dicUnPassCountByDomainGradeYear = new Dictionary<string, Dictionary<string, int>>();
        /// <summary>
        /// 不及格領域數人數
        /// </summary>
        private Dictionary<string, Dictionary<int, int>> dicUnPassCountByUnPassGradeYear = new Dictionary<string, Dictionary<int, int>>();
        private BackgroundWorker bgWorker = new BackgroundWorker();

        public frmExportSchoolDomainScore()
        {
            InitializeComponent();
            InitializeBackgroundWorker();
        }

        private void frmExportSchoolDomainScore_Load(object sender, EventArgs e)
        {
            // Init SchoolYear Semester
            {
                int schoolYear = int.Parse(School.DefaultSchoolYear);
                cbxSchoolYear.Items.Add(schoolYear - 1);
                cbxSchoolYear.Items.Add(schoolYear);
                cbxSchoolYear.Items.Add(schoolYear + 1);
                cbxSchoolYear.SelectedIndex = 1;

                int semester = int.Parse(School.DefaultSemester);
                cbxSemester.Items.Add(1);
                cbxSemester.Items.Add(2);
                cbxSemester.SelectedIndex = semester - 1;
            }
            cbxRange.Items.Add("補考前");
            cbxRange.Items.Add("補考後");
            cbxRange.SelectedIndex = 0;
            // 載入樣板
            Stream stream = new MemoryStream(Properties.Resources.SchoolDomainScore_template);
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
                        listDomain.Add(name);
                    }
                }
            }
        }

        private void InitializeBackgroundWorker()
        {
            bgWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);
            bgWorker.WorkerReportsProgress = true;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (!bgWorker.IsBusy)
            {
                ParameterRec data = new ParameterRec();
                data.SchoolYear = cbxSchoolYear.SelectedItem.ToString();
                data.Semester = cbxSemester.SelectedItem.ToString();
                data.Range = cbxRange.SelectedItem.ToString();

                bgWorker.RunWorkerAsync(data);
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            ParameterRec data = (ParameterRec)e.Argument;
            int progress = 0;

            // 取得各年級領域成績
            DataTable dt = GetDomainScore(data.SchoolYear, data.Semester);
            bgWorker.ReportProgress(progress += 15);

            // 資料解析
            ParseData(dt, data.Range);
            bgWorker.ReportProgress(progress += 15);

            DataTable table = FillMergeFiledData();
            bgWorker.ReportProgress(progress += 25);

            Document doc = new Document();
            doc.Sections.Clear();
            Section section = (Section)doc.ImportNode(docTemplate.FirstSection, true);
            doc.AppendChild(section);
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            doc.MailMerge.Execute(table);
            bgWorker.ReportProgress(progress += 25);

            string path = $"{Application.StartupPath}\\Reports\\領域不及格人數統計表.docx";
            int i = 1;
            while (File.Exists(path))
            {
                string docName = Path.GetFileNameWithoutExtension(path);
                string newPath = $"{Path.GetDirectoryName(path)}\\領域不及格人數統計表{i++}{Path.GetExtension(path)}";
                path = newPath;
            }

            doc.Save(path, SaveFormat.Docx);
            bgWorker.ReportProgress(100);
            e.Result = path;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            string path = (string)e.Result;
            MotherForm.SetStatusBarMessage("領域不及格人數統計表 產生完成");
            DialogResult result = MsgBox.Show($"{path}\n領域不及格人數統計表產生完成，是否立刻開啟？", "訊息", MessageBoxButtons.YesNo);

            if (DialogResult.Yes == result)
            {
                System.Diagnostics.Process.Start(path);
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("領域不及格人數統計表 產生中:", e.ProgressPercentage);
        }

        private DataTable GetDomainScore(string schoolYear, string semester)
        {
            string sql = string.Format(@"
WITH data_row AS (
	SELECT
	    {0}::INT AS school_year
	    , {1}::INT AS semester
), target_student AS(
	SELECT
		*
	FROM
		student
	WHERE
		status IN(1, 2)
)
SELECT
	sems_subj_score_ext.ref_student_id
	, class.grade_year
	, sems_subj_score_ext.school_year
	, sems_subj_score_ext.semester
	, array_to_string(xpath('/Domain/@原始成績', subj_score_ele), '')::text AS 原始成績
	, array_to_string(xpath('/Domain/@成績', subj_score_ele), '')::text AS 成績
	, array_to_string(xpath('/Domain/@補考成績', subj_score_ele), '')::text AS 補考成績
	, array_to_string(xpath('/Domain/@領域', subj_score_ele), '')::text AS 領域
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
	INNER JOIN data_row
	    ON data_row.school_year = sems_subj_score_ext.school_year
	    AND data_row.semester = sems_subj_score_ext.semester
WHERE
    class.grade_year IS NOT NULL
ORDER BY
	sems_subj_score_ext.grade_year
            ", schoolYear, semester);

            return qh.Select(sql);
        }

        private void ParseData(DataTable dt, string range)
        {
            dicUnPassCountByDomainGradeYear = new Dictionary<string, Dictionary<string, int>>();
            dicUnPassCountByUnPassGradeYear = new Dictionary<string, Dictionary<int, int>>();
            dicStuRecByID = new Dictionary<string, StudentRec>();
            listDomainFromData.Clear();

            // 資料整理
            foreach (DataRow row in dt.Rows)
            {
                // 學生資料整理
                string stuID = "" + row["ref_student_id"];
                string domain = "" + row["領域"];

                if (!dicStuRecByID.ContainsKey(stuID))
                {
                    dicStuRecByID.Add(stuID, new StudentRec());
                }
                StudentRec stuRec = dicStuRecByID[stuID];
                stuRec.ID = stuID;
                stuRec.GradeYear = "" + row["grade_year"];

                if (domain != "")
                {
                    if (!stuRec.dicScoreByDomain.ContainsKey(domain))
                    {
                        stuRec.dicScoreByDomain.Add(domain, new DomainRec());
                    }
                    DomainRec domainRec = stuRec.dicScoreByDomain[domain];
                    domainRec.Name = domain;
                    domainRec.OriginScore = "" + row["原始成績"];
                    domainRec.ReTestScore = "" + row["補考成績"];
                    domainRec.Score = "" + row["成績"];

                    // 領域資料整理
                    if (!listDomainFromData.Contains(domain))
                    {
                        listDomainFromData.Add(domain);
                    }
                }
            }

            // 領域資料排序與刪除
            listDomainFromData = DomainListParse(listDomainFromData);

            // 統計「學生領域不及格數」、「年級領域不及格數」
            foreach (string stuID in dicStuRecByID.Keys)
            {
                StudentRec stuRec = dicStuRecByID[stuID];
                string gradeYear = stuRec.GradeYear;

                foreach (string domain in listDomainFromData)
                {
                    if (stuRec.dicScoreByDomain.ContainsKey(domain))
                    {
                        DomainRec domainRec = stuRec.dicScoreByDomain[domain];

                        #region 學生領域不及格數
                        {
                            if (domainRec.OriginScore != "")
                            {
                                stuRec.OriginScoreUnPassCount += FloatParse(domainRec.OriginScore) < 60 ? 1 : 0;
                            }
                            if (domainRec.ReTestScore != "")
                            {
                                stuRec.ReTestScoreUnPassCount += FloatParse(domainRec.ReTestScore) < 60 ? 1 : 0;
                            }
                            if (domainRec.Score != "")
                            {
                                stuRec.ScoreUnPassCount += FloatParse(domainRec.Score) < 60 ? 1 : 0;
                            }
                        }
                        #endregion

                        #region 年級領域不及格數
                        {
                            string score;
                            if (range == "補考前")
                            {
                                score = domainRec.Score;
                            }
                            else
                            {
                                score = domainRec.ReTestScore;
                            }

                            if (score != "")
                            {
                                // 年級
                                if (!dicUnPassCountByDomainGradeYear.ContainsKey(gradeYear))
                                {
                                    dicUnPassCountByDomainGradeYear.Add(stuRec.GradeYear, new Dictionary<string, int>());
                                }

                                // 領域
                                Dictionary<string, int> dicUnPassCountByDomain = dicUnPassCountByDomainGradeYear[gradeYear];
                                if (!dicUnPassCountByDomain.ContainsKey(domain))
                                {
                                    dicUnPassCountByDomain.Add(domain, 0);
                                }

                                dicUnPassCountByDomain[domain] += FloatParse(score) < 60 ? 1 : 0;
                            }
                        }
                        #endregion
                    }
                }
            }

            // 領域不及格數人數統計
            foreach (string stuID in dicStuRecByID.Keys)
            {
                StudentRec stuRec = dicStuRecByID[stuID];
                int unPassCount = 0;

                if (range == "補考前")
                {
                    unPassCount = stuRec.ScoreUnPassCount;
                }
                else
                {
                    unPassCount = stuRec.ReTestScoreUnPassCount;
                }

                if (unPassCount > 0)
                {
                    if (!dicUnPassCountByUnPassGradeYear.ContainsKey(stuRec.GradeYear))
                    {
                        dicUnPassCountByUnPassGradeYear.Add(stuRec.GradeYear, new Dictionary<int, int>());
                    }
                    Dictionary<int, int> dicUnPassCountByUnPass = dicUnPassCountByUnPassGradeYear[stuRec.GradeYear];

                    if (!dicUnPassCountByUnPass.ContainsKey(unPassCount))
                    {
                        dicUnPassCountByUnPass.Add(unPassCount, 0);
                    }
                    dicUnPassCountByUnPass[unPassCount] += 1;
                }
                
            }
        }

        /// <summary>
        /// 解析領域清單
        /// 1. 根據領域對照表排序
        /// 2. 最多七個領域多的刪除
        /// </summary>
        /// <returns></returns>
        private List<string> DomainListParse(List<string> listData)
        {
            listData.Sort(delegate (string a, string b)
            {
                int aIndex = listDomain.FindIndex(domain => domain == a);
                int bIndex = listDomain.FindIndex(domain => domain == b);

                if (aIndex > bIndex)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            });

            if (listData.Count > 7)
            {
                listData.RemoveRange(7, listData.Count - 7);
            }

            return listData;
        }

        private DataTable CreatMergeFiledTable()
        {
            DataTable table = new DataTable();
            DataColumn col;

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "school_year";
            table.Columns.Add(col);

            col = new DataColumn();
            col.DataType = Type.GetType("System.String");
            col.ColumnName = "semester";
            table.Columns.Add(col);

            for (int i = 1; i <= 7; i++)
            {
                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"domain_{i}";
                table.Columns.Add(col);

                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"domain_{i}_unpass_count";
                table.Columns.Add(col);

                col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = $"unpass_{i}_stu_count";
                table.Columns.Add(col);

                for (int y = 1; y <= 3; y++)
                {
                    col = new DataColumn();
                    col.DataType = Type.GetType("System.String");
                    col.ColumnName = $"grade_year_{y}_domain_{i}_unpass_count";
                    table.Columns.Add(col);

                    col = new DataColumn();

                    col.DataType = Type.GetType("System.String");
                    col.ColumnName = $"grade_year_{y}_unpass_{i}_stu_count";
                    table.Columns.Add(col);
                }
            }

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

            return table;
        }

        private DataTable FillMergeFiledData()
        {
            DataTable dt = CreatMergeFiledTable();
            DataRow row = dt.NewRow();

            row["school_year"] = School.DefaultSchoolYear;
            row["semester"] = School.DefaultSemester;
            row["year"] = DateTime.Now.Year;
            row["month"] = DateTime.Now.Month;
            row["day"] = DateTime.Now.Day;

            int d = 1;
            foreach (string domain in listDomainFromData)
            {
                row[$"domain_{d}"] = domain;

                int unPassTotalCount = 0;
                for (int y = 1; y <= 3; y++)
                {
                    string gradeYear = $"{y}";
                    if (dicUnPassCountByDomainGradeYear.ContainsKey(gradeYear))
                    {
                        if (dicUnPassCountByDomainGradeYear[gradeYear].ContainsKey(domain))
                        {
                            row[$"grade_year_{y}_domain_{d}_unpass_count"] = dicUnPassCountByDomainGradeYear[gradeYear][domain];
                            unPassTotalCount += dicUnPassCountByDomainGradeYear[gradeYear][domain];
                        }
                    }
                }
                row[$"domain_{d}_unpass_count"] = unPassTotalCount;

                d++;
            }

            for (int i = 1; i <= 7; i++)
            {
                int unPassTotalCount = 0;
                for (int y = 1; y <= 3; y++)
                {
                    string gradeYear = $"{y}";
                    if (dicUnPassCountByUnPassGradeYear.ContainsKey(gradeYear))
                    {
                        if (dicUnPassCountByUnPassGradeYear[gradeYear].ContainsKey(i))
                        {
                            row[$"grade_year_{y}_unpass_{i}_stu_count"] = dicUnPassCountByUnPassGradeYear[gradeYear][i];
                            unPassTotalCount += dicUnPassCountByUnPassGradeYear[gradeYear][i];
                        }
                        else
                        {
                            row[$"grade_year_{y}_unpass_{i}_stu_count"] = 0;
                        }
                    }
                    else
                    {
                        row[$"grade_year_{y}_unpass_{i}_stu_count"] = 0;
                    }
                }

                row[$"unpass_{i}_stu_count"] = unPassTotalCount;
                
            }

            dt.Rows.Add(row);

            return dt;
        }

        private float FloatParse(string score)
        {
            return float.Parse(score == "" ? "0" : score);
        }

        private void btnLeave_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private class StudentRec
        {
            public string ID;
            public string GradeYear;
            public Dictionary<string, DomainRec> dicScoreByDomain = new Dictionary<string, DomainRec>();
            public int OriginScoreUnPassCount;
            public int ScoreUnPassCount;
            public int ReTestScoreUnPassCount;
        }

        private class DomainRec
        {
            public string Name;
            public string OriginScore;
            public string Score;
            public string ReTestScore;
        }

        private class ParameterRec
        {
            public string SchoolYear;
            public string Semester;
            public string Range;
        }
    }
}
