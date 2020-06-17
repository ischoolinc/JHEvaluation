using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Xml;
using FISCA.Data;
using FISCA.Presentation.Controls;
using Aspose.Cells;
using K12.Data;
using K12.Data.Configuration;

namespace KaoHsingReExamScoreReport.Forms
{
    public partial class ReDomainForTeacherForm : BaseForm
    {
        private float _passScore = 60;
        private List<string> listClassID;
        private List<string> listDomainFromConfig = new List<string>();
        private BackgroundWorker _bgWorker = new BackgroundWorker();
        private Dictionary<string, Dictionary<string, StudentRec>> dicStuRecByClassID;
        private Dictionary<string, List<string>> dicDomainByClassID;
        private Workbook wbTemplate;
        private QueryHelper qh = new QueryHelper();

        public ReDomainForTeacherForm(List<string> ClassIDList)
        {
            InitializeComponent();

            listClassID = ClassIDList;
        }

        private void ReDomainForTeacherForm_Load(object sender, EventArgs e)
        {
            // 預設學年度、學期
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            #region BackGroundWorker Init
            {
                _bgWorker.DoWork += _bgWorker_DoWork;
                _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
                _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
                _bgWorker.WorkerReportsProgress = true;
            }
            #endregion

            #region 取得領域對照表
            {
                ConfigData cd = School.Configuration["JHEvaluation_Subject_Ordinal"];
                {
                    if (cd.Contains("DomainOrdinal"))
                    {
                        XmlElement element = cd.GetXml("DomainOrdinal", XmlHelper.LoadXml("<Domains/>"));
                        foreach (XmlElement domainElement in element.SelectNodes("Domain"))
                        {
                            string name = domainElement.GetAttribute("Name");
                            listDomainFromConfig.Add(name);
                        }
                    }
                }
            }
            #endregion

            // 載入樣板
            wbTemplate = new Workbook(new MemoryStream(Properties.Resources.ReTestDomainReport_ForTeacher));
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;

            _bgWorker.RunWorkerAsync();
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // 取得班級學生領域成績
            GetStuDomainScore();
            _bgWorker.ReportProgress(20);

            // 取得班級領域清單
            GetClassDomainList();
            _bgWorker.ReportProgress(40);

            // 班級領域清單排序
            foreach (string classID in dicDomainByClassID.Keys)
            {
                List<string> listDomain = dicDomainByClassID[classID];

                // 資料排序
                listDomain.Sort(delegate (string a, string b)
                {
                    int aIndex = listDomainFromConfig.FindIndex(name => name == a);
                    int bIndex = listDomainFromConfig.FindIndex(name => name == b);

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
            _bgWorker.ReportProgress(60);

            Workbook wb = new Workbook();
            int sheetIndex = 0;

            // 總表功能          
            Worksheet wst = wb.Worksheets[0];
            wst.Name = "總表";
            wst.Copy(wbTemplate.Worksheets["總表樣板"]);

            // 需要補考
            List<string> domainNameList = new List<string>();
            Dictionary<string, int> domainIdxDict = new Dictionary<string, int>();
            List<string> tmpNameList = new List<string>();
            int colIdx = 4;

            foreach (string classID in listClassID)
            {
                if (dicDomainByClassID.ContainsKey(classID))
                {
                    foreach (string dName in dicDomainByClassID[classID])
                    {
                        if (!tmpNameList.Contains(dName))
                            tmpNameList.Add(dName);
                    }
                }
            }

            if (tmpNameList.Contains("語文"))
                domainNameList.Add("語文");

            //  依對照表順序
            foreach (string name in listDomainFromConfig)
            {
                if (tmpNameList.Contains(name))
                    domainNameList.Add(name);
            }


            // 填入欄位
            foreach (string dName in domainNameList)
            {
                wst.Cells[0, colIdx].PutValue(dName);
                domainIdxDict.Add(dName + "補考", colIdx);
                colIdx++;
            }
            foreach (string dName in domainNameList)
            {
                wst.Cells[0, colIdx].PutValue(dName);
                domainIdxDict.Add(dName + "成績", colIdx);
                colIdx++;
            }

            int rowIdx = 1;
            // 產生資料至總表
            foreach (string classID in listClassID)
            {
                List<string> listDomain = new List<string>();
                if (dicDomainByClassID.ContainsKey(classID))
                {
                    listDomain = dicDomainByClassID[classID];
                    if (dicStuRecByClassID.ContainsKey(classID))
                    {
                        foreach (string stuID in dicStuRecByClassID[classID].Keys)
                        {
                            StudentRec stuRec = dicStuRecByClassID[classID][stuID];

                            bool checkReExam = false;
                            // 檢查是否需要補考
                            foreach (string domain in listDomain)
                            {
                                if (stuRec.dicScoreRecByDomain.ContainsKey(domain))
                                {
                                    if (stuRec.dicScoreRecByDomain[domain].isPass == false)
                                    {
                                        checkReExam = true;
                                    }
                                }
                            }

                            // 需要補考才產生在總表
                            if (checkReExam)
                            {
                                wst.Cells[rowIdx, 0].PutValue(stuRec.GradeYear);
                                wst.Cells[rowIdx, 1].PutValue(stuRec.ClassName);
                                wst.Cells[rowIdx, 2].PutValue(stuRec.SeatNo);
                                wst.Cells[rowIdx, 3].PutValue(stuRec.StudentName);

                                foreach (string domain in listDomain)
                                {
                                    if (stuRec.dicScoreRecByDomain.ContainsKey(domain))
                                    {
                                        if (stuRec.dicScoreRecByDomain[domain].isPass)
                                        {
                                            //
                                        }
                                        else
                                        {
                                            if (domainIdxDict.ContainsKey(domain + "補考"))
                                            {
                                                wst.Cells[rowIdx, domainIdxDict[domain + "補考"]].PutValue("補考");
                                            }

                                        }

                                        if (domainIdxDict.ContainsKey(domain + "成績"))
                                        {
                                            decimal score;

                                            if (decimal.TryParse(stuRec.dicScoreRecByDomain[domain].OriginScore, out score))
                                            {
                                                wst.Cells[rowIdx, domainIdxDict[domain + "成績"]].PutValue(score);
                                            }
                                        }

                                    }
                                }

                                rowIdx++;
                            }


                        }
                    }
                }
            }
            wst.AutoFitColumns();


            sheetIndex = 1;

            // 各班分列
            foreach (string classID in listClassID)
            {
                string className = K12.Data.Class.SelectByID(classID).Name;

                List<string> listDomain = new List<string>();
                if (dicDomainByClassID.ContainsKey(classID))
                {
                    listDomain = dicDomainByClassID[classID];
                }

                if (sheetIndex > 0)
                {
                    wb.Worksheets.Add();
                }
                Worksheet ws = wb.Worksheets[sheetIndex++];
                ws.Name = className;
                ws.Copy(wbTemplate.Worksheets["班級樣板"]);
                Range colRange = wbTemplate.Worksheets["班級樣板"].Cells.CreateRange(2, 2, 1, 1);

                int rowIndex = 0;
                int colIndex = 0;
                int domainCount = listDomain.Count;
                string title = $"{iptSchoolYear.Value}學年度 第{iptSemester.Value}學期 {className}領域補考名單";

                ws.Cells.Merge(rowIndex, colIndex, 1, (domainCount * 2 + 1) < 6 ? 6 : domainCount * 2 + 1);
                ws.Cells[rowIndex, colIndex].PutValue(title);

                if (dicStuRecByClassID.ContainsKey(classID))
                {
                    rowIndex++;
                    colIndex = 2;
                    // 班級領域清單
                    foreach (string domain in listDomain)
                    {
                        ws.Cells[rowIndex, colIndex++].PutValue(domain);
                    }

                    rowIndex++;
                    foreach (string stuID in dicStuRecByClassID[classID].Keys)
                    {
                        colIndex = 0;
                        StudentRec stuRec = dicStuRecByClassID[classID][stuID];
                        Dictionary<string, ScoreRec> dicScoreRecByDomain = stuRec.dicScoreRecByDomain;

                        // 座號
                        {
                            Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                            range.CopyStyle(colRange);
                            ws.Cells[rowIndex, colIndex++].PutValue(stuRec.SeatNo);
                        }

                        // 姓名
                        {
                            Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                            range.CopyStyle(colRange);
                            ws.Cells[rowIndex, colIndex++].PutValue(stuRec.StudentName);
                        }

                        foreach (string domain in listDomain)
                        {
                            Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                            range.CopyStyle(colRange);

                            if (dicScoreRecByDomain.ContainsKey(domain))
                            {
                                string score = dicScoreRecByDomain[domain].OriginScore;
                                string data = dicScoreRecByDomain[domain].isPass ? $"{score}" : $"補考/ {score}";
                                ws.Cells[rowIndex, colIndex].PutValue(data);
                            }

                            colIndex++;
                        }

                        rowIndex++;
                    }
                }
                else
                {
                    ws.Cells[2, 0].PutValue("此班級無學生。");
                    ws.Cells.Merge(2, 0, 1, 3);
                }
            }

            e.Result = wb;
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("領域補考名單產生中 ...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnPrint.Enabled = true;

            Workbook wb = (Workbook)e.Result;

            if (wb != null)
            {
                Utility.CompletedXls("領域補考名單-給導師", wb);
            }

            FISCA.Presentation.MotherForm.SetStatusBarMessage("領域補考名單產生完成.");
        }

        /// <summary>
        /// 取得班級學生領域成績
        /// </summary>
        private void GetStuDomainScore()
        {
            dicStuRecByClassID = new Dictionary<string, Dictionary<string, StudentRec>>();

            #region SQL
            string sql = $@"
WITH target_student AS(
	SELECT
		*
	FROM
		student
	WHERE
		ref_class_id IN ({string.Join(",", listClassID)})
        AND status IN (1, 2)
)
SELECT
    target_student.id
    , target_student.name
    , target_student.seat_no
    , class.id AS class_id
    , class.class_name
    , class.grade_year
	, array_to_string(xpath('/Domain/@原始成績', subj_score_ele), '')::text AS 原始成績
	, array_to_string(xpath('/Domain/@成績', subj_score_ele), '')::text AS 成績
	, array_to_string(xpath('/Domain/@領域', subj_score_ele), '')::text AS 領域
    , array_to_string(xpath('/Domain/@權數', subj_score_ele), '')::text AS 權數
FROM 
    target_student
    LEFT OUTER JOIN (
		SELECT 
			sems_subj_score.*
			, unnest(xpath('/root/Domains/Domain', xmlparse(content '<root>' || score_info || '</root>'))) as subj_score_ele
		FROM 
			sems_subj_score 
			INNER JOIN target_student
				ON target_student.id = sems_subj_score.ref_student_id 
        WHERE
            sems_subj_score.school_year = {iptSchoolYear.Value}
            AND sems_subj_score.semester = {iptSemester.Value}
	) as sems_subj_score_ext
        ON sems_subj_score_ext.ref_student_id = target_student.id
    LEFT OUTER JOIN class
        ON class.id = target_student.ref_class_id
ORDER BY
    class.display_order
    , target_student.seat_no
            ";
            #endregion

            DataTable dt = qh.Select(sql);

            foreach (DataRow row in dt.Rows)
            {
                string classID = "" + row["class_id"];
                string stuID = "" + row["id"];

                // 學生資料整理
                if (!dicStuRecByClassID.ContainsKey(classID))
                {
                    dicStuRecByClassID.Add(classID, new Dictionary<string, StudentRec>());
                }

                Dictionary<string, StudentRec> dicStuRecByID = dicStuRecByClassID[classID];

                if (!dicStuRecByID.ContainsKey(stuID))
                {
                    StudentRec stuRec = new StudentRec();
                    stuRec.GradeYear = "" + row["grade_year"];
                    stuRec.ClassName = "" + row["class_name"];
                    stuRec.SeatNo = "" + row["seat_no"];
                    stuRec.StudentName = "" + row["name"];
                    stuRec.dicScoreRecByDomain = new Dictionary<string, ScoreRec>();

                    dicStuRecByID.Add(stuID, stuRec);
                }

                string domain = "" + row["領域"];

                if (!dicStuRecByID[stuID].dicScoreRecByDomain.ContainsKey(domain))
                {
                    ScoreRec scoreRec = new ScoreRec();
                    scoreRec.Domain = domain;
                    scoreRec.OriginScore = "" + row["原始成績"];
                    scoreRec.isPass = FloatParse("" + row["原始成績"]) >= _passScore;

                    dicStuRecByID[stuID].dicScoreRecByDomain.Add(domain, scoreRec);
                }

            }
        }

        /// <summary>
        /// 取得班級領域清單
        /// </summary>
        private void GetClassDomainList()
        {
            dicDomainByClassID = new Dictionary<string, List<string>>();

            #region SQL
            string sql = $@"
WITH target_student AS(
	SELECT
		*
	FROM
		student
	WHERE
		ref_class_id IN ({string.Join(",", listClassID)})
        AND status IN (1, 2)
)
SELECT DISTINCT
    class.id
	, array_to_string(xpath('/Domain/@領域', subj_score_ele), '')::text AS 領域
    , class.display_order
FROM
    target_student
    LEFT OUTER JOIN (
		SELECT 
			sems_subj_score.*
			, unnest(xpath('/root/Domains/Domain', xmlparse(content '<root>' || score_info || '</root>'))) as subj_score_ele
		FROM 
			sems_subj_score 
			INNER JOIN target_student
				ON target_student.id = sems_subj_score.ref_student_id 
        WHERE
            sems_subj_score.school_year = {iptSchoolYear.Value}
            AND sems_subj_score.semester = {iptSemester.Value}
	) as sems_subj_score_ext
    ON sems_subj_score_ext.ref_student_id = target_student.id
    LEFT OUTER JOIN class
        ON class.id = target_student.ref_class_id
ORDER BY
    class.display_order
";
            #endregion

            DataTable dt = qh.Select(sql);

            foreach (DataRow row in dt.Rows)
            {
                string classID = "" + row["id"];
                string domain = "" + row["領域"];

                if (!dicDomainByClassID.ContainsKey(classID))
                {
                    dicDomainByClassID.Add(classID, new List<string>());
                }

                List<string> listDomain = dicDomainByClassID[classID];

                if (domain.Trim() != "" && !listDomain.Contains(domain))
                {
                    listDomain.Add(domain);
                }
            }
        }

        private static float FloatParse(string score)
        {
            return float.Parse(score == "" ? "0" : score);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private class StudentRec
        {
            public string GradeYear;
            public string ClassName;
            public string SeatNo;
            public string StudentName;
            public Dictionary<string, ScoreRec> dicScoreRecByDomain;
        }

        public class ScoreRec
        {
            public string Domain;
            public string OriginScore;
            public bool isPass;
        }
    }

}
