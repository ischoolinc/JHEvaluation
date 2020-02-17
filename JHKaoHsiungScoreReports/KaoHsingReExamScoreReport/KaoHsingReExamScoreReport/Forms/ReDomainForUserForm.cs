using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Collections.Generic;
using System.ComponentModel;
using K12.Data;
using K12.Data.Configuration;
using FISCA.Data;
using FISCA.Presentation.Controls;
using Aspose.Cells;

namespace KaoHsingReExamScoreReport.Forms
{
    public partial class ReDomainForUserForm : BaseForm
    {
        private float _passScore = 60;
        private List<string> listClassID;
        private List<string> listDomainFromConfig = new List<string>();
        private List<string> listDomain;
        private Dictionary<string, Dictionary<string, int>> dicClassUnPassCountByDomain;
        private QueryHelper qh = new QueryHelper();
        private Workbook wbTemplate;
        private BackgroundWorker _bgWorker = new BackgroundWorker();

        public ReDomainForUserForm(List<string> ClassIDList)
        {
            InitializeComponent();
            listClassID = ClassIDList;
        }

        private void ReDomainForUserForm_Load(object sender, EventArgs e)
        {
            // 預設學年度、學期
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            #region Init BackGroundWorker
            {
                _bgWorker.WorkerReportsProgress = true;
                _bgWorker.DoWork += _bgWorker_DoWork;
                _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
                _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
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
            wbTemplate = new Workbook(new MemoryStream(Properties.Resources.ReTestDomainReport_ForSchool));
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // 取得領域清單
            GetDomainList();
            _bgWorker.ReportProgress(20);

            // 取得學生領域成績
            GetStudentDomainScore();
            _bgWorker.ReportProgress(40);

            // 班級領域清單排序
            listDomain.Sort(delegate (string a, string b) {
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
            _bgWorker.ReportProgress(60);

            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "";
            ws.Copy(wbTemplate.Worksheets[0]);
            Range colRange = wbTemplate.Worksheets[0].Cells.CreateRange(0, 1, 1, 1);

            int rowIndex = 0;
            int colIndex = 0;

            // 領域
            colIndex++;
            foreach (string domain in listDomain)
            {
                Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                range.CopyStyle(colRange);
                ws.Cells[rowIndex, colIndex++].PutValue(domain);
            }

            rowIndex++;
            // 班級領域補考人數
            foreach (string classID in listClassID)
            {
                colIndex = 0;
                // 班級
                {
                    Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                    range.CopyStyle(colRange);
                    string className = K12.Data.Class.SelectByID(classID).Name;
                    ws.Cells[rowIndex, colIndex++].PutValue(className);
                }

                // 領域補考人數
                {
                    if (dicClassUnPassCountByDomain.ContainsKey(classID))
                    {
                        foreach (string domain in listDomain)
                        {
                            Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                            range.CopyStyle(colRange);
                            if (dicClassUnPassCountByDomain[classID].ContainsKey(domain))
                            {
                                int unPassCount = dicClassUnPassCountByDomain[classID][domain];
                                ws.Cells[rowIndex, colIndex++].PutValue(unPassCount);
                            }
                            else
                            {
                                ws.Cells[rowIndex, colIndex++].PutValue(0);
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < listDomain.Count; i++)
                        {
                            Range range = ws.Cells.CreateRange(rowIndex, colIndex, 1, 1);
                            range.CopyStyle(colRange);
                            ws.Cells[rowIndex, colIndex++].PutValue(0);
                        }
                    }
                }

                rowIndex++;
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
                Utility.CompletedXls("領域補考名單-給試務", wb);
            }
                
            FISCA.Presentation.MotherForm.SetStatusBarMessage("領域補考名單產生完成.");
        }

        /// <summary>
        /// 取得選取班級領域清單
        /// </summary>
        private void GetDomainList()
        {
            listDomain = new List<string>();

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
	array_to_string(xpath('/Domain/@領域', subj_score_ele), '')::text AS 領域
FROM (
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
";
            #endregion

            DataTable dt = qh.Select(sql);

            foreach (DataRow row in dt.Rows)
            {
                string domain = "" + row["領域"];

                if (!listDomain.Contains(domain))
                {
                    listDomain.Add(domain);
                }
            }
        }

        /// <summary>
        /// 取得學生領域成績 
        /// </summary>
        private void GetStudentDomainScore()
        {
            dicClassUnPassCountByDomain = new Dictionary<string, Dictionary<string, int>>();

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
    student.id
    , class.id AS class_id
    , class.class_name
    , class.grade_year
	, array_to_string(xpath('/Domain/@原始成績', subj_score_ele), '')::text AS 原始成績
	, array_to_string(xpath('/Domain/@成績', subj_score_ele), '')::text AS 成績
	, array_to_string(xpath('/Domain/@領域', subj_score_ele), '')::text AS 領域
FROM (
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
    LEFT OUTER JOIN student
        ON student.id = sems_subj_score_ext.ref_student_id
    LEFT OUTER JOIN class
        ON class.id = student.ref_class_id
ORDER BY
    class.display_order
";
            #endregion

            DataTable dt = qh.Select(sql);

            foreach (DataRow row in dt.Rows)
            {
                string classID = "" + row["class_id"];
                string domain = "" + row["領域"];
                float originScore = FloatParse("" + row["原始成績"]);

                if (!dicClassUnPassCountByDomain.ContainsKey(classID))
                {
                    dicClassUnPassCountByDomain.Add(classID, new Dictionary<string, int>());
                }

                Dictionary<string, int> dicDomainUnPassCount = dicClassUnPassCountByDomain[classID];

                if (!dicDomainUnPassCount.ContainsKey(domain))
                {
                    dicDomainUnPassCount.Add(domain, 0);
                }

                if (originScore < _passScore)
                {
                    dicDomainUnPassCount[domain]++;
                }
            }
        }

        private static float FloatParse(string score)
        {
            return float.Parse(score == "" ? "0" : score);
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;
            _bgWorker.RunWorkerAsync();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
