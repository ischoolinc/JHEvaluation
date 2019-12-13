using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using Aspose.Words;
using System.IO;

namespace HsinChuExamScoreClassFixedRank.DAO
{
    /// <summary>
    /// 資料存取使用
    /// </summary>
    public class DataAccess
    {

        /// <summary>
        /// 匯出合併欄位總表Word
        /// </summary>
        public static void ExportMappingFieldWord()
        {

        }


        /// <summary>
        /// 透過班級ID、學年度、學期 取得學生修課試別、領域、科目
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, List<string>>> GetExamDomainSubjectDictByClass(string SchoolYear, string Semester, List<string> ClassIDList, string ExamID)
        {
            Dictionary<string, Dictionary<string, List<string>>> value = new Dictionary<string, Dictionary<string, List<string>>>();

            if (ClassIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "SELECT te_include.ref_exam_id AS exam_id,course.domain,course.subject FROM sc_attend INNER JOIN course ON sc_attend.ref_course_id = course.id INNER JOIN student ON sc_attend.ref_student_id = student.id  INNER JOIN te_include ON course.ref_exam_template_id = te_include.ref_exam_template_id WHERE student.status = 1 AND course.ref_class_id IN(" + string.Join(",", ClassIDList.ToArray()) + ") AND course.school_year = " + SchoolYear + " AND course.semester = " + Semester + " AND te_include.ref_exam_id = " + ExamID + " ORDER BY domain,subject";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string SubjectName = "";
                    // 科目名稱空白略過
                    if (dr["subject"] == null)
                        continue;

                    SubjectName = dr["subject"].ToString();

                    if (string.IsNullOrWhiteSpace(SubjectName))
                        continue;

                    string exam_id = dr["exam_id"].ToString();
                    // 領域空白為彈性課程
                    string domain = "彈性課程";
                    if (dr["domain"] != null && dr["domain"].ToString() != "")
                    {
                        domain = dr["domain"].ToString();
                    }

                    string subject = SubjectName;// dr["subject"].ToString();

                    // 試別
                    if (!value.ContainsKey(exam_id))
                    {
                        value.Add(exam_id, new Dictionary<string, List<string>>());
                    }

                    // 領域
                    if (!value[exam_id].ContainsKey(domain))
                        value[exam_id].Add(domain, new List<string>());

                    // 科目
                    if (!value[exam_id][domain].Contains(subject))
                        value[exam_id][domain].Add(subject);
                }
            }

            return value;
        }


        /// <summary>
        /// 取得評分樣板上定期，平時是100-定期
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, decimal> GetScorePercentageHS()
        {
            Dictionary<string, decimal> returnData = new Dictionary<string, decimal>();
            FISCA.Data.QueryHelper qh1 = new FISCA.Data.QueryHelper();
            string query1 = @"SELECT id,CAST(regexp_replace( xpath_string(exam_template.extension,'/Extension/ScorePercentage'), '^$', '0') as integer) AS ScorePercentage  FROM exam_template";
            System.Data.DataTable dt1 = qh1.Select(query1);

            foreach (System.Data.DataRow dr in dt1.Rows)
            {
                string id = dr["id"].ToString();
                decimal sp = 50;
                if (decimal.TryParse(dr["ScorePercentage"].ToString(), out sp))
                    returnData.Add(id, sp);
                else
                    returnData.Add(id, 50);

            }
            return returnData;
        }

        public static List<ClassInfo> GetClassStudentsByClassID(List<string> ClassIDs)
        {
            List<ClassInfo> value = new List<ClassInfo>();
            // SELECT class.id AS class_id,class.class_name,class.grade_year,student.seat_no,student.name AS student_name FROM student INNER  JOIN Class ON student.ref_class_id = class.id WHERE student.status = 1 AND class.id IN (10,5) ORDER BY class.grade_year,class.display_order,class.class_name,student.seat_no
            Dictionary<string, ClassInfo> tmpDict = new Dictionary<string, ClassInfo>();
            if (ClassIDs.Count > 0)
            {

                QueryHelper qh = new QueryHelper();
                string query = "SELECT " +
                    "class.id AS class_id" +
                    ",class.class_name" +
                    ",class.grade_year" +
                    ",student.seat_no" +
                    ",student.name AS student_name" +
                    ",student.id AS student_id" +
                    " FROM student" +
                    " INNER  JOIN Class" +
                    " ON student.ref_class_id = class.id" +
                    " WHERE " +
                    "student.status = 1 AND" +
                    " class.id IN (" + string.Join(",", ClassIDs.ToArray()) + ") ORDER BY class.grade_year,class.display_order,class.class_name,student.seat_no";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string class_id = dr["class_id"].ToString();

                    if (!tmpDict.ContainsKey(class_id))
                    {
                        ClassInfo ci = new ClassInfo();
                        ci.ClassID = class_id;
                        ci.ClassName = dr["class_name"].ToString();
                        int gr = 0;
                        int.TryParse(dr["grade_year"].ToString(), out gr);
                        ci.GradeYear = gr;
                        ci.Students = new List<StudentInfo>();
                        tmpDict.Add(class_id, ci);
                    }

                    // 加入學生
                    StudentInfo si = new StudentInfo();
                    si.Name = dr["student_name"].ToString();
                    si.SeatNo = dr["seat_no"].ToString();
                    si.StudentID = dr["student_id"].ToString();
                    si.ClassID = class_id;

                    if (tmpDict.ContainsKey(class_id))
                    {
                        tmpDict[class_id].Students.Add(si);
                    }
                }
            }

            foreach (string cid in tmpDict.Keys)
            {
                value.Add(tmpDict[cid]);
            }
            return value;
        }


        /// <summary>
        /// 載入學生成績
        /// </summary>
        /// <param name="ClassInfoList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="ExamID"></param>
        /// <returns></returns>
        public static List<ClassInfo> LoadClassStudentScore(List<ClassInfo> ClassInfoList, string SchoolYear, string Semester, string ExamID, Dictionary<string, decimal> ScorePercentageHS, List<string> ClassIDList)
        {
            // 學生領域科目成績
            Dictionary<string, Dictionary<string, DomainInfo>> tmpDomainDict = new Dictionary<string, Dictionary<string, DomainInfo>>();
            if (ClassIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "SELECT " +
                    "student.id AS student_id" +
                    ",course.domain" +
                    ",course.subject" +
                    ",course.credit" +
                    ",sce_take.ref_exam_id AS exam_id" +
                    ",course.ref_exam_template_id AS template_id" +
                    ",array_to_string(xpath('/Extension/AssignmentScore/text()',xmlparse(content sce_take.extension)),'') AS assignment_score" +
                    ",array_to_string(xpath('/Extension/Score/text()',xmlparse(content sce_take.extension)),'') AS f_score" +
                    " FROM sc_attend INNER JOIN course" +
                    " ON sc_attend.ref_course_id = course.id INNER JOIN student" +
                    " ON sc_attend.ref_student_id = student.id INNER JOIN sce_take" +
                    " ON sce_take.ref_sc_attend_id = sc_attend.id " +
                    " INNER JOIN te_include ON course.ref_exam_template_id = te_include.ref_exam_template_id AND sce_take.ref_exam_id = te_include.ref_exam_id" +
                    " WHERE" +
                    " te_include.ref_exam_id = " + ExamID + " AND course.school_year = " + SchoolYear + " AND course.semester = " + Semester + " AND course.ref_class_id IN(" + string.Join(",", ClassIDList.ToArray()) + ") AND student.status = 1";

                DataTable dt = qh.Select(query);

                foreach (DataRow dr in dt.Rows)
                {
                    string SubjectName = "";
                    // 科目名稱空白不處理
                    if (dr["subject"] == null)
                    {
                        continue;
                    }
                    SubjectName = dr["subject"].ToString();

                    if (string.IsNullOrWhiteSpace(SubjectName))
                        continue;

                    string student_id = dr["student_id"].ToString();
                    if (!tmpDomainDict.ContainsKey(student_id))
                        tmpDomainDict.Add(student_id, new Dictionary<string, DomainInfo>());

                    string dName = "彈性課程";
                    if (dr["domain"] != null && dr["domain"].ToString() != "")
                        dName = dr["domain"].ToString();

                    if (!tmpDomainDict[student_id].ContainsKey(dName))
                    {
                        DomainInfo di = new DomainInfo();
                        di.Name = dName;
                        di.SubjectInfoList = new List<SubjectInfo>();
                        tmpDomainDict[student_id].Add(dName, di);
                    }

                    // 定期比例預設
                    decimal sfp = 50 * 0.01M, sfa = 50 * 0.01M;


                    string template_id = "";

                    if (dr["template_id"] != null && dr["template_id"].ToString() != "")
                        template_id = dr["template_id"].ToString();

                    // 處理評量比例
                    if (ScorePercentageHS.ContainsKey(template_id))
                    {
                        sfp = ScorePercentageHS[template_id] * 0.01M;
                        sfa = (100 - ScorePercentageHS[template_id]) * 0.01M;
                    }

                    // 科目成績
                    SubjectInfo si = new SubjectInfo();
                    si.Name = SubjectName;//dr["subject"].ToString();
                    si.DomainName = dName;

                    if (dr["credit"] != null)
                    {
                        decimal cr;
                        if (decimal.TryParse(dr["credit"].ToString(), out cr))
                        {
                            si.Credit = cr;
                        }
                        else
                            si.Credit = null;
                    }
                    else
                        si.Credit = null;


                    si.ScoreAP = sfa;
                    si.ScoreFP = sfp;

                    if (dr["assignment_score"] != null && dr["assignment_score"].ToString() != "")
                    {
                        decimal sa;
                        if (decimal.TryParse(dr["assignment_score"].ToString(), out sa))
                        {
                            si.ScoreA = sa;
                        }

                    }
                    else
                    {
                        si.ScoreA = null;
                    }

                    if (dr["f_score"] != null && dr["f_score"].ToString() != "")
                    {
                        decimal sf;
                        if (decimal.TryParse(dr["f_score"].ToString(), out sf))
                        {
                            si.ScoreF = sf;
                        }

                    }
                    else
                    {
                        si.ScoreF = null;
                    }

                    // 計算科目評量總成績
                    si.CalcScore();

                    if (tmpDomainDict[student_id].ContainsKey(si.DomainName))
                    {
                        tmpDomainDict[student_id][si.DomainName].SubjectInfoList.Add(si);
                    }

                }
            }

            // 資料回填
            foreach (ClassInfo ci in ClassInfoList)
            {
                foreach (StudentInfo si in ci.Students)
                {
                    if (tmpDomainDict.ContainsKey(si.StudentID))
                    {
                        // 先清空
                        si.DomainInfoList.Clear();
                        foreach (string doName in tmpDomainDict[si.StudentID].Keys)
                        {
                            // 計算領域成績
                            tmpDomainDict[si.StudentID][doName].CalScore();

                            // 回寫
                            si.DomainInfoList.Add(tmpDomainDict[si.StudentID][doName]);
                        }
                    }
                }
            }

            return ClassInfoList;
        }


        /// <summary>
        /// 取得學生定期評量固定排名
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="ExamID"></param>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static DataTable GetStudentExamRankMatrix(string SchoolYear, string Semester, string ExamID, List<string> ClassIDList)
        {
            DataTable value = new DataTable();
            QueryHelper qh = new QueryHelper();
            string query = @"
SELECT 
	rank_matrix.id AS rank_matrix_id
	, rank_matrix.school_year
	, rank_matrix.semester
	, rank_matrix.grade_year
	, rank_matrix.item_type
	, rank_matrix.ref_exam_id
	, rank_matrix.item_name
	, rank_matrix.rank_type
	, rank_matrix.rank_name
	, class.class_name
	, student.seat_no
	, student.student_number
	, student.name
	, rank_detail.ref_student_id
	, rank_detail.rank
	, rank_detail.pr
	, rank_detail.percentile
    , rank_matrix.avg_top_25
    , rank_matrix.avg_top_50
    , rank_matrix.avg
    , rank_matrix.avg_bottom_50
    , rank_matrix.avg_bottom_25
    , rank_detail.score
FROM 
	rank_matrix
	LEFT OUTER JOIN rank_detail
		ON rank_detail.ref_matrix_id = rank_matrix.id
	LEFT OUTER JOIN student
		ON student.id = rank_detail.ref_student_id
	LEFT OUTER JOIN class
		ON class.id = student.ref_class_id
WHERE
	rank_matrix.is_alive = true
	AND rank_matrix.school_year = '" + SchoolYear + @"'
    AND rank_matrix.semester = '" + Semester + @"'
	AND rank_matrix.item_type like '定期評量%'
	AND rank_matrix.ref_exam_id = '" + ExamID + @"'
    AND class.id IN (" + string.Join(",", ClassIDList.ToArray()) + ") " +
    "ORDER BY rank_matrix.id" +
    ", rank_detail.rank	" +
    ", class.grade_year" +
    ", class.display_order" +
    ", class.class_name" +
    ", student.seat_no" +
    ", student.id";

            value = qh.Select(query);

            return value;
        }


        public static Dictionary<string, string> GetClassTeacherNameDictByClassID(List<string> ClassIDList)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            if (ClassIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "SELECT " +
                    "class.id AS class_id" +
                    ",teacher.teacher_name" +
                    ",teacher.nickname " +
                    "FROM " +
                    "class INNER JOIN teacher " +
                    "ON class.ref_teacher_id = teacher.id " +
                    "WHERE teacher.status = 1 AND class.id IN(" + string.Join(",", ClassIDList.ToArray()) + ")";

                DataTable dt = qh.Select(query);

                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string class_id = dr["class_id"].ToString();
                        string teacher_name = "", nickname = "";
                        if (dr["teacher_name"] != null)
                            teacher_name = dr["teacher_name"].ToString();

                        if (dr["nickname"] != null)
                        {
                            nickname = dr["nickname"].ToString();
                        }

                        if (!value.ContainsKey(class_id))
                        {
                            if (!string.IsNullOrWhiteSpace(nickname))
                            {
                                teacher_name = teacher_name + "(" + nickname + ")";
                            }
                            value.Add(class_id, teacher_name);
                        }
                    }
                }

            }
            return value;
        }


        /// <summary>
        /// 取得班級年級五標及組距
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="ExamID"></param>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static DataTable GetClassExamRankMatrix(string SchoolYear, string Semester, string ExamID, List<string> ClassIDList)
        {
            DataTable value = new DataTable();
            QueryHelper qh = new QueryHelper();
            // 定期評量:item_type:領域成績、科目成績,rank_type:班排名,年排名
            string query = @"
SELECT 
	DISTINCT rank_matrix.id AS rank_matrix_id
	, class.id AS class_id
	, rank_matrix.school_year
	, rank_matrix.semester
	, rank_matrix.grade_year
	, rank_matrix.item_type
	, rank_matrix.ref_exam_id
	, rank_matrix.item_name
	, rank_matrix.rank_type
	, rank_matrix.rank_name
	, rank_matrix.matrix_count
    , rank_matrix.level_gte100
    , rank_matrix.level_90
    , rank_matrix.level_80
    , rank_matrix.level_70
    , rank_matrix.level_60
    , rank_matrix.level_50
    , rank_matrix.level_40
    , rank_matrix.level_30
    , rank_matrix.level_20
    , rank_matrix.level_10
    , rank_matrix.level_lt10
    , rank_matrix.avg_top_25
    , rank_matrix.avg_top_50
    , rank_matrix.avg
    , rank_matrix.avg_bottom_50
    , rank_matrix.avg_bottom_25   
FROM 
	rank_matrix
	LEFT OUTER JOIN rank_detail
		ON rank_detail.ref_matrix_id = rank_matrix.id
	LEFT OUTER JOIN student
		ON student.id = rank_detail.ref_student_id
	LEFT OUTER JOIN class
		ON class.id = student.ref_class_id
WHERE
	rank_matrix.is_alive = true AND rank_matrix.school_year = " + SchoolYear + @" 
	AND rank_matrix.semester = " + Semester + @" AND rank_matrix.item_type IN('定期評量/領域成績','定期評量/科目成績','定期評量_定期/領域成績','定期評量_定期/科目成績')
	AND rank_type IN('班排名','年排名') 
	AND rank_matrix.ref_exam_id = " + ExamID + @" AND class.id IN(" + string.Join(",", ClassIDList.ToArray()) + ")";

            value = qh.Select(query);

            return value;
        }

    }
}
