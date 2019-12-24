using System;
using System.Drawing;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.IO;
using K12.Data;
using FISCA.Data;
using Aspose.Cells;
using K12.Data.Configuration;
using System.Xml;
using FISCA.Presentation;
using FISCA.Presentation.Controls;

namespace DomainScoreReport
{
    public partial class frmExportSchoolDomainScore : BaseForm
    {
        private const string ConfigName = "JHEvaluation_Subject_Ordinal";
        private const string ColumnKey = "DomainOrdinal";
        private QueryHelper qh = new QueryHelper();
        private Workbook wbTemplate;
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
        /// <summary>
        /// 年級清單
        /// </summary>
        private List<string> listGradeYear = new List<string>();
        /// <summary>
        /// 學生領域資料
        /// </summary>
        private Dictionary<string, Dictionary<string, StudentRec>> dicStuRecByIDSemester = new Dictionary<string, Dictionary<string, StudentRec>>();
        /// <summary>
        /// 領域不及格數
        /// </summary>
        private Dictionary<string, Dictionary<string, int>> dicUnPassCountByDomainGradeYear = new Dictionary<string, Dictionary<string, int>>();
        /// <summary>
        /// 總計 領域不及格數
        /// </summary>
        private Dictionary<string, int> dicTotalCountByDomain = new Dictionary<string, int>();
        /// <summary>
        /// 不及格領域數人數
        /// </summary>
        private Dictionary<string, Dictionary<int, int>> dicUnPassCountByUnPassGradeYear = new Dictionary<string, Dictionary<int, int>>();
        /// <summary>
        /// 總計 不及格領域數 總人數
        /// </summary>
        private Dictionary<int, int> dicTotalCountByUnPass = new Dictionary<int, int>();
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
            {
                Stream wbStream = new MemoryStream(Properties.Resources.SchoolDomainScore_template);
                wbTemplate = new Workbook(wbStream);
            }
            
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

            Workbook wb = FillWorkBookData(data);
            bgWorker.ReportProgress(progress += 25);

            string path = $"{Application.StartupPath}\\Reports\\領域不及格人數統計表.xlsx";
            int i = 1;
            while (File.Exists(path))
            {
                string docName = Path.GetFileNameWithoutExtension(path);
                string newPath = $"{Path.GetDirectoryName(path)}\\領域不及格人數統計表{i++}{Path.GetExtension(path)}";
                path = newPath;
            }

            wb.Save(path, SaveFormat.Xlsx);
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

            this.Close();
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("領域不及格人數統計表 產生中:", e.ProgressPercentage);
        }

        private DataTable GetDomainScore(string schoolYear, string semester)
        {
            string condition = "";
            if (cbxIsSchoolYear.Checked)
            {
                condition = string.Format(@"
WHERE
    class.grade_year IS NOT NULL
    AND sems_subj_score_ext.school_year = {0}
                    ", schoolYear);
            }
            else
            {
                condition = string.Format(@"
WHERE
    class.grade_year IS NOT NULL
    AND sems_subj_score_ext.school_year = {0}
    AND sems_subj_score_ext.semester = {1}
                    ", schoolYear, semester);
            }
            string sql = string.Format(@"
WITH target_student AS(
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
    {0}
ORDER BY
	sems_subj_score_ext.grade_year
            ", condition);

            return qh.Select(sql);
        }

        private void ParseData(DataTable dt, string range)
        {
            // 學生資料整理
            foreach (DataRow row in dt.Rows)
            {
                string stuID = "" + row["ref_student_id"];
                string domain = "" + row["領域"];
                string gradeYear = "" + row["grade_year"];
                string semester = "" + row["semester"];

                if (!dicStuRecByIDSemester.ContainsKey(semester))
                {
                    dicStuRecByIDSemester.Add(semester, new Dictionary<string, StudentRec>());
                }
                if (!dicStuRecByIDSemester[semester].ContainsKey(stuID))
                {
                    dicStuRecByIDSemester[semester].Add(stuID, new StudentRec());
                }
                StudentRec stuRec = dicStuRecByIDSemester[semester][stuID];
                stuRec.ID = stuID;
                stuRec.GradeYear = gradeYear;
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

                    // 領域清單
                    if (!listDomainFromData.Contains(domain))
                    {
                        listDomainFromData.Add(domain);
                    }

                    //  年級清單
                    if (!listGradeYear.Contains(gradeYear))
                    {
                        listGradeYear.Add(gradeYear);
                    }
                }
            }

            // 領域資料排序與刪除
            listDomainFromData = DomainListParse(listDomainFromData);

            // 統計「學生領域不及格數」、「年級領域不及格數」(區分補考前後)
            foreach (string semester in dicStuRecByIDSemester.Keys)
            {
                foreach (string stuID in dicStuRecByIDSemester[semester].Keys)
                {
                    StudentRec stuRec = dicStuRecByIDSemester[semester][stuID];
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
                                    score = domainRec.OriginScore;
                                }
                                else
                                {
                                    score = domainRec.Score;
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
            }

            // 領域不及格數人數統計 (區分補考前後)
            foreach (string semester in dicStuRecByIDSemester.Keys)
            {
                foreach (string stuID in dicStuRecByIDSemester[semester].Keys)
                {
                    StudentRec stuRec = dicStuRecByIDSemester[semester][stuID];
                    int unPassCount = 0;

                    if (range == "補考前")
                    {
                        unPassCount = stuRec.OriginScoreUnPassCount;
                    }
                    else
                    {
                        unPassCount = stuRec.ScoreUnPassCount;
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

            // 總計 領域不及格數 (區分補考前後)
            foreach (string domain in listDomainFromData)
            {
                if (!dicTotalCountByDomain.ContainsKey(domain))
                {
                    dicTotalCountByDomain.Add(domain, 0);
                }
            }
            foreach (string gradeyear in dicUnPassCountByDomainGradeYear.Keys)
            {
                foreach (string domain in dicUnPassCountByDomainGradeYear[gradeyear].Keys)
                {
                    dicTotalCountByDomain[domain] += dicUnPassCountByDomainGradeYear[gradeyear][domain];
                }
            }
            // 總計 不及格領域數總人數 (區分補考前後)
            for (int i = 1; i <= listDomainFromData.Count; i++)
            {
                dicTotalCountByUnPass.Add(i, 0);
            }
            foreach (string gradeyear in dicUnPassCountByUnPassGradeYear.Keys)
            {
                foreach (int unpass in dicUnPassCountByUnPassGradeYear[gradeyear].Keys)
                {
                    dicTotalCountByUnPass[unpass] += dicUnPassCountByUnPassGradeYear[gradeyear][unpass];
                }
            }

            // 年級清單排序
            listGradeYear.Sort();
        }

        /// <summary>
        /// 解析領域清單
        /// 根據領域對照表排序
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

            return listData;
        }

        private Workbook FillWorkBookData(ParameterRec data)
        {
            Workbook prototype = new Workbook();
            prototype.Copy(wbTemplate);
            Worksheet ps = prototype.Worksheets[0];

            Workbook wb = new Workbook();
            wb.Copy(prototype);
            Worksheet ws = wb.Worksheets[0];
            
            Range colRange = ps.Cells.CreateRange(2, 1, 1, 1);
            Range rowRange = ps.Cells.CreateRange(3, 1, false);

            int rowIndex = 0;
            int colIndex = 0;
            int domainCount = listDomainFromData.Count;
            
            string title = cbxIsSchoolYear.Checked ? $"{data.SchoolYear}學年度 {data.Semester}學期 領域不及格人數統計表"
                    : $"{data.SchoolYear}學年度 領域不及格人數統計表";
            ws.Cells.Merge(rowIndex, 0, 1, listDomainFromData.Count * 2);
            ws.Cells[rowIndex++, 0].PutValue(title);

            // 各學習領域學生成績評量情形
            {
                ws.Cells.Merge(1, 1, 1, domainCount);

                Range range = ws.Cells.CreateRange(1, 1, 1, domainCount);
                range.SetOutlineBorders(CellBorderType.Thin, Color.Black);
                range.PutValue("各學習領域學生成績評量情形", false, false);

                Cell cell = ws.Cells.GetCell(1, 1);
                Style style = cell.GetStyle();
                style.HorizontalAlignment = TextAlignmentType.Center;
                cell.SetStyle(style);
            }
            // 學生成績評量不及格領域數情形
            {
                ws.Cells.Merge(1, domainCount + 1, 1, domainCount);
                
                Range range = ws.Cells.CreateRange(1, domainCount + 1, 1, domainCount);
                range.SetOutlineBorders(CellBorderType.Thin, Color.Black);
                range.PutValue("學生成績評量不及格領域數情形", false, false);

                Cell cell = ws.Cells.GetCell(1, domainCount + 1);
                Style style = cell.GetStyle();
                style.HorizontalAlignment = TextAlignmentType.Center;
                cell.SetStyle(style);
            }
            rowIndex++;

            colIndex = 1;
            // 領域
            foreach (string domain in listDomainFromData)
            {
                Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                range.CopyStyle(colRange);
                range.ColumnWidth = colRange.ColumnWidth;
                ws.Cells[rowIndex, colIndex++].PutValue(domain);
            }

            // 不及格數
            for (int i = 1; i <= listDomainFromData.Count; i++)
            {
                Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                range.CopyStyle(colRange);
                range.ColumnWidth = colRange.ColumnWidth;
                ws.Cells[rowIndex, colIndex++].PutValue($"{i}個學習領域不及格人數");
            }

            rowIndex++;
            // 年級
            foreach (string gradeYear in listGradeYear)
            {
                colIndex = 0;

                if (rowIndex > 3)
                {
                    ws.Cells.InsertRow(rowIndex);
                }
                ws.Cells.CreateRange(rowIndex, 1, false).CopyStyle(rowRange);
                ws.Cells[rowIndex, colIndex++].PutValue($"{gradeYear}年級");

                // 領域不及格人數
                foreach (string domain in listDomainFromData)
                {
                    Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                    range.CopyStyle(colRange);

                    if (dicUnPassCountByDomainGradeYear[gradeYear].ContainsKey(domain))
                    {
                        int unPassCount = dicUnPassCountByDomainGradeYear[gradeYear][domain];
                        ws.Cells[rowIndex, colIndex++].PutValue(unPassCount);
                    }
                    else
                    {
                        ws.Cells[rowIndex, colIndex++].PutValue(0);
                    }
                }

                // 不及格領域人數
                for (int i = 1; i <= listDomainFromData.Count; i++)
                {
                    Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                    range.CopyStyle(colRange);

                    if (dicUnPassCountByUnPassGradeYear.ContainsKey(gradeYear))
                    {
                        if (dicUnPassCountByUnPassGradeYear[gradeYear].ContainsKey(i))
                        {
                            int unPassCount = dicUnPassCountByUnPassGradeYear[gradeYear][i];
                            ws.Cells[rowIndex, colIndex++].PutValue(unPassCount);
                        }
                        else
                        {
                            ws.Cells[rowIndex, colIndex++].PutValue(0);
                        }
                    }
                }
                rowIndex++;
            }

            colIndex = 1;
            // 總計 領域
            foreach (string domain in listDomainFromData)
            {
                Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                range.CopyStyle(colRange);

                ws.Cells[rowIndex, colIndex++].PutValue(dicTotalCountByDomain[domain]);
            }
            // 總計 不及格領域數
            foreach (int i in dicTotalCountByUnPass.Keys)
            {
                Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                range.CopyStyle(colRange);

                ws.Cells[rowIndex, colIndex++].PutValue(dicTotalCountByUnPass[i]);
            }
            rowIndex++;

            int year = DateTime.Now.Year - 1911;
            int month = DateTime.Now.Month;
            int day = DateTime.Now.Day;

            ws.Cells[rowIndex, 0].PutValue($"列印日期: {year}年{month}月{day}日");
            return wb;
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
            public int OriginScoreUnPassCount = 0;
            public int ScoreUnPassCount = 0;
            public int ReTestScoreUnPassCount = 0;
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
