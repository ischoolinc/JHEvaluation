using System;
using System.Xml;
using System.Data;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using FISCA.Data;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Data.Configuration;
using Aspose.Words;
using Campus.Report;

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
            Stream stream = new MemoryStream(Properties.Resources.StudentDoaminScore_template);
            docTemplat = new Document(stream);
        }

        /// <summary>
        /// 成績單列印
        /// </summary>
        public void Export()
        {
            bool isSuccess = GetStudentDomainScore();
            if (isSuccess)
            {
                if (dtRsp.Rows.Count > 0)
                {
                    ParseData();

                    List<string> listPath = new List<string>();
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    DataTable dt = ParseMergeSource();
                    foreach (DataRow row in dt.Rows)
                    {
                        Document doc = (Document)docTemplat.Clone(true);
                        doc.MailMerge.Execute(row);
                        string _path = $"{path}\\{ row["student_name"]}_成績預警通知單.docx";
                        doc.Save(_path, SaveFormat.Docx);
                        listPath.Add(_path);
                        //ReportSaver.SaveDocument(doc, $"{"" + row["student_name"]}_成績預警通知單");
                    }

                    DialogResult result =　MsgBox.Show("成績預警通知單產生完成，是否立刻開啟？", "訊息", MessageBoxButtons.YesNo);

                    if (DialogResult.Yes == result)
                    {
                        foreach (string _path in listPath)
                        {
                            using (FileStream fs = new FileStream(_path, FileMode.Open));
                        }
                    }

                }
                else
                {
                    MsgBox.Show("沒有資料。");
                }
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
ORDER BY
	sems_subj_score_ext.grade_year
	, school_year
	, semester
            ", string.Join(",", listStudentIDs));

            try
            {
                dtRsp = qh.Select(sql);
                return true;
            }
            catch (Exception error)
            {
                MsgBox.Show("成績單列印失敗:" + error.Message);
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
                    string score = "" + row["成績"];
                    string ssKey = schoolYear + semester;

                    // 成績
                    if (!stuRec.dicScoreByDomainBySchoolYear.ContainsKey(ssKey))
                    {
                        stuRec.dicScoreByDomainBySchoolYear.Add(ssKey, new Dictionary<string, string>());
                    }
                    stuRec.dicScoreByDomainBySchoolYear[ssKey].Add(domain, score);
                    // 領域
                    if (!stuRec.dicDomainByName.Keys.Contains(domain))
                    {
                        DomainRec domainRec = new DomainRec();
                        domainRec.Name = domain;
                        stuRec.dicDomainByName.Add(domain, domainRec);
                    }
                    stuRec.dicDomainByName[domain].TotalScore += float.Parse("" + row["成績"] == "" ? "0" : "" + row["成績"]);
                    stuRec.dicDomainByName[domain].ScoreCount += 1;
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

            // 計算平均
            foreach (string id in dicStudentByID.Keys)
            {
                foreach (DomainRec doc in dicStudentByID[id].dicDomainByName.Values)
                {
                    doc.Calc();
                }
            }
        }

        /// <summary>
        /// 將資料解析為 doc merge 格式
        /// </summary>
        private DataTable ParseMergeSource()
        {
            string schoolName = School.ChineseName;
            DataTable dt = MakeDataTable();
            
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

                int s = 1;
                double tTotalScore = 0;
                int tCount = 0;
                // 學年度學期
                foreach (string key in dicStudentByID[id].dicSchoolYear.Keys)
                {
                    SchoolYearSemester ss = dicStudentByID[id].dicSchoolYear[key];
                    row[$"{s}_school_year"] = $"{ss.SchoolYear}學年度";
                    row[$"{s}_semester"] = $"第{ss.Semester}學期";

                    
                    int sd = 1;
                    float sTotalScore = 0;
                    int sCount = 0;
                    // 領域
                    foreach (string domainName in dicStudentByID[id].dicDomainByName.Keys)
                    {
                        // 成績
                        if (dicStudentByID[id].dicScoreByDomainBySchoolYear[key].ContainsKey(domainName))
                        {
                            row[$"{sd}_domain_{s}_score"] = dicStudentByID[id].dicScoreByDomainBySchoolYear[key][domainName];
                            sTotalScore += float.Parse(dicStudentByID[id].dicScoreByDomainBySchoolYear[key][domainName]);
                            sCount++;
                        }
                        else
                        {
                            row[$"{sd}_domain_{s}_score"] = "";
                        }
                        sd++;
                    }

                    // 學期領域平均
                    if (sTotalScore > 0 && sCount > 0)
                    {
                        double avg = Math.Round(sTotalScore / sCount, 2);
                        row[$"{s}_all_domain_avg"] = avg;
                        tTotalScore += avg;
                        tCount++;
                    }
                    s++;
                }

                // 總平均
                if (tTotalScore > 0 && tCount > 0)
                {
                    row["all_domain_avg"] = Math.Round(tTotalScore / tCount, 2);
                }

                int d = 0;
                int passDomainCount = 0;
                foreach (string key in dicStudentByID[id].dicDomainByName.Keys)
                {
                    d++;
                    DomainRec dr = dicStudentByID[id].dicDomainByName[key];
                    // 領域名稱
                    row[$"{d}_domain"] = dr.Name;
                    // 領域平均
                    row[$"{d}_domain_avg"] = dr.AvgScore;

                    if (dr.AvgScore > 60)
                    {
                        passDomainCount++;
                    }
                }
                row["pass_domain_count"] = passDomainCount;

                dt.Rows.Add(row);
            }

            return dt;
        }

        private DataTable MakeDataTable()
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
    }

    class StudentRec
    {
        public string ID;
        public string ClassName;
        public string SeatNo;
        public string Name;
        public Dictionary<string, DomainRec> dicDomainByName = new Dictionary<string, DomainRec>();
        public Dictionary<string, SchoolYearSemester> dicSchoolYear = new Dictionary<string, SchoolYearSemester>();
        public Dictionary<string, Dictionary<string, string>> dicScoreByDomainBySchoolYear = new Dictionary<string, Dictionary<string, string>>();
        public Dictionary<string, string> dicAvgScoreByDomain = new Dictionary<string, string>();
    }

    class SchoolYearSemester
    {
        public string SchoolYear;
        public string Semester;
    }

    class DomainRec
    {
        public string Name;
        public float TotalScore;
        public float ScoreCount;
        public double AvgScore;

        public void Calc()
        {
            AvgScore = Math.Round(TotalScore / ScoreCount, 2);
        }
    }
}
