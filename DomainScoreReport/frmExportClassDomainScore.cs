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
using K12.Data.Configuration;
using FISCA.Data;
using Aspose.Words;
using System.IO;
using System.Xml;
using Aspose.Words.Reporting;
using FISCA.Presentation;

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
        private Dictionary<string, Dictionary<string, Dictionary<string, float>>> dicClassStuDomainScore = new Dictionary<string, Dictionary<string, Dictionary<string, float>>>();
        private Dictionary<string, string> dicClassNameByID = new Dictionary<string, string>();
        private Dictionary<string, StudentRec> dicStuRecByID = new Dictionary<string, StudentRec>();
        private Dictionary<string, List<string>> dicDomainNameByClassID = new Dictionary<string, List<string>>();
        private BackgroundWorker bgWorker = new BackgroundWorker();
        private int percentage = 0;

        public frmExportClassDomainScore(List<string>listIDs)
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
            // init schoolYear semester
            int schoolYear = int.Parse(School.DefaultSchoolYear);
            int semester = int.Parse(School.DefaultSemester);

            cbxSchoolYear.Items.Add(schoolYear - 1);
            cbxSchoolYear.Items.Add(schoolYear);
            cbxSchoolYear.Items.Add(schoolYear + 1);
            cbxSchoolYear.SelectedIndex = 1;

            cbxSemester.Items.Add(1);
            cbxSemester.Items.Add(2);
            cbxSemester.SelectedIndex = semester - 1;

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
            if (!bgWorker.IsBusy)
            {
                ParameterRec data = new ParameterRec();
                data.SchoolYear = cbxSchoolYear.SelectedItem.ToString();
                data.Semester = cbxSemester.SelectedItem.ToString();
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

            int p = 70 / table.Rows.Count;
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
		ref_class_id IN ({2})
), domain_score AS(
    SELECT
        student.id
        , student.name
        , student.seat_no
        , class.id AS class_id
        , class.class_name
	    --, sems_subj_score_ext.ref_student_id
	    , sems_subj_score_ext.semester
	    , sems_subj_score_ext.school_year
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
	    ) as sems_subj_score_ext
        LEFT OUTER JOIN student
            ON student.id = sems_subj_score_ext.ref_student_id
        LEFT OUTER JOIN class
            ON class.id = student.ref_class_id
        INNER JOIN data_row
    	    ON data_row.school_year = sems_subj_score_ext.school_year
    	    AND data_row.semester = sems_subj_score_ext.semester
    ORDER BY
	    sems_subj_score_ext.grade_year 
        , class.display_order
        , student.seat_no
) -- 篩除沒有 domain 的資料 
SELECT
    *
FROM
    domain_score
WHERE
    領域 <> ''
                ", data.SchoolYear, data.Semester, data.ClassIDs);

            return qh.Select(sql);
        }

        private void ParseData(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                string classID = "" + row["class_id"];
                string studentID = "" + row["id"];
                string domain = "" + row["領域"];
                float score = float.Parse("" + row["成績"] == "" ? "0" : "" + row["成績"]);

                // 成績資料
                if (!dicClassStuDomainScore.ContainsKey(classID))
                {
                    dicClassStuDomainScore.Add(classID, new Dictionary<string, Dictionary<string, float>>());
                }
                if (!dicClassStuDomainScore[classID].ContainsKey(studentID))
                {
                    dicClassStuDomainScore[classID].Add(studentID, new Dictionary<string, float>());
                }
                dicClassStuDomainScore[classID][studentID].Add(domain, score);
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

            for (int i = 1; i <= 40; i++)
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
                row["school_year"] = data.SchoolYear;
                row["semester"] = data.Semester;
                row["class_name"] = dicClassNameByID[classID];

                // 班級領域清單
                List<string> listDomain = dicDomainNameByClassID[classID];
                // 根據領域對照表做排序
                listDomain.Sort(delegate(string a, string b)
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
                Dictionary<string, Dictionary<string, float>> dicStuDomainScore = dicClassStuDomainScore[classID];
                int s = 1;
                foreach (string stuID in dicStuDomainScore.Keys)
                {
                    StudentRec stuRec = dicStuRecByID[stuID];
                    row[$"seat_no_{s}"] = stuRec.SeatNo;
                    row[$"name_{s}"] = stuRec.Name;

                    int d = 1;
                    int sc = 0;
                    int passCount = 0;
                    float totalScore = 0;
                    // 領域成績
                    foreach (string domain in listDomain)
                    {
                        if (dicStuDomainScore[stuID].ContainsKey(domain))
                        {
                            float score = dicStuDomainScore[stuID][domain];
                            row[$"stu_{s}_domain_{d}"] = score;
                            totalScore += score;

                            if (score >= 60)
                            {
                                passCount++;
                            }
                            sc++;
                        }
                        d++;
                    }
                    if (sc > 0)
                    {
                        // 平均成績
                        row[$"stu_{s}_avg_score"] = Math.Round(totalScore / sc, 2);
                        // 領域及格數
                        row[$"stu_{s}_pass_count"] = passCount;
                    }
                    s++;
                }

                dt.Rows.Add(row);
            }

            return dt;
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
            public string SchoolYear;
            public string Semester;
            public string ClassIDs;
        }
    }
}
