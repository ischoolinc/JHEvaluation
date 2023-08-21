using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using Aspose.Words;
using System.IO;


namespace HsinChuSemesterClassFixedRank.DAO
{
    /// <summary>
    /// 資料存取使用
    /// </summary>
    public class DataAccess
    {

        public static Dictionary<string, string> ClassTag1Dict = new Dictionary<string, string>();
        public static Dictionary<string, string> ClassTag2Dict = new Dictionary<string, string>();

        /// <summary>
        /// 透過 ClassID 取得班導師
        /// </summary>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
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
                    "class LEFT JOIN teacher " +
                    "ON class.ref_teacher_id = teacher.id " +
                    "WHERE teacher.status = 1 AND class.id IN(" + string.Join(",", ClassIDList.ToArray()) + ") " +
                    "ORDER BY class.grade_year,class.display_order,class_name";

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
        /// 取得學生學期排名、五標與組距資料值
        ///  </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, Dictionary<string, string>>> GetSemsScoreRankMatrixDataValue(string SchoolYear, string Semester, List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> value = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            // 沒有學生不處理
            if (StudentIDList.Count == 0)
                return value;

            List<string> r2List = new List<string>();
            r2List.Add("rank");
            r2List.Add("matrix_count");
            r2List.Add("pr");
            r2List.Add("percentile");
            r2List.Add("avg_top_25");
            r2List.Add("avg_top_50");
            r2List.Add("avg");
            r2List.Add("avg_bottom_50");
            r2List.Add("avg_bottom_25");
            r2List.Add("level_gte100");
            r2List.Add("level_90");
            r2List.Add("level_80");
            r2List.Add("level_70");
            r2List.Add("level_60");
            r2List.Add("level_50");
            r2List.Add("level_40");
            r2List.Add("level_30");
            r2List.Add("level_20");
            r2List.Add("level_10");
            r2List.Add("level_lt10");

            // 需要四捨五入
            List<string> r2ListNP = new List<string>();
            r2ListNP.Add("avg_top_25");
            r2ListNP.Add("avg_top_50");
            r2ListNP.Add("avg");
            r2ListNP.Add("avg_bottom_50");
            r2ListNP.Add("avg_bottom_25");




            QueryHelper qh = new QueryHelper();
            string query = "" +
               " SELECT " +
" 	rank_matrix.id AS rank_matrix_id" +
" 	, rank_matrix.school_year" +
" 	, rank_matrix.semester" +
" 	, rank_matrix.grade_year" +
" 	, rank_matrix.item_type" +
" 	, rank_matrix.ref_exam_id AS exam_id" +
" 	, rank_matrix.item_name" +
" 	, rank_matrix.rank_type" +
" 	, rank_matrix.rank_name" +
" 	, class.class_name" +
" 	, student.seat_no" +
" 	, student.student_number" +
" 	, student.name" +
" 	, rank_detail.ref_student_id AS student_id " +
" 	, rank_detail.rank" +
"   , rank_matrix.matrix_count " +
" 	, rank_detail.pr" +
" 	, rank_detail.percentile" +
"   , rank_matrix.avg_top_25" +
"   , rank_matrix.avg_top_50" +
"   , rank_matrix.avg" +
"   , rank_matrix.avg_bottom_50" +
"   , rank_matrix.avg_bottom_25" +
" 	, rank_matrix.level_gte100" +
" 	, rank_matrix.level_90" +
" 	, rank_matrix.level_80" +
" 	, rank_matrix.level_70" +
" 	, rank_matrix.level_60" +
" 	, rank_matrix.level_50" +
" 	, rank_matrix.level_40" +
" 	, rank_matrix.level_30" +
" 	, rank_matrix.level_20" +
" 	, rank_matrix.level_10" +
" 	, rank_matrix.level_lt10" +
" FROM " +
" 	rank_matrix" +
" 	LEFT OUTER JOIN rank_detail" +
" 		ON rank_detail.ref_matrix_id = rank_matrix.id" +
" 	LEFT OUTER JOIN student" +
" 		ON student.id = rank_detail.ref_student_id" +
" 	LEFT OUTER JOIN class" +
" 		ON class.id = student.ref_class_id" +
" WHERE" +
" 	rank_matrix.is_alive = true" +
" 	AND rank_matrix.school_year = " + SchoolYear +
"     AND rank_matrix.semester = " + Semester +
" 	AND rank_matrix.item_type like '學期%'" +
" 	AND rank_matrix.ref_exam_id = -1 " +
"     AND ref_student_id IN (" + string.Join(",", StudentIDList.ToArray()) + ") " +
" ORDER BY " +
" 	rank_matrix.id" +
" 	, rank_detail.rank" +
" 	, class.grade_year" +
" 	, class.display_order" +
" 	, class.class_name" +
" 	, student.seat_no" +
" 	, student.id";

            DataTable dt = qh.Select(query);
            //dt.TableName = "d5";
            //dt.WriteXmlSchema(Application.StartupPath + "\\d5s.xml");
            //dt.WriteXml(Application.StartupPath + "\\d5d.xml");

            //StudentTag1Dict.Clear();
            //StudentTag2Dict.Clear();
            // student id key
            // key = item_type + item_name +  rank_name 
            foreach (DataRow dr in dt.Rows)
            {
                string sid = dr["student_id"].ToString();
                if (!value.ContainsKey(sid))
                    value.Add(sid, new Dictionary<string, Dictionary<string, string>>());

                string key = dr["item_type"].ToString() + "_" + dr["item_name"].ToString() + "_" + dr["rank_type"].ToString();

                //if (key == "學期/總計成績_課程學習總成績_類別1排名")
                //{
                //    if (dr["rank_name"] != null)
                //    {
                //        if (!StudentTag1Dict.ContainsKey(sid))
                //            StudentTag1Dict.Add(sid, dr["rank_name"].ToString());
                //    }
                //}

                //if (key == "學期/總計成績_課程學習總成績_類別2排名")
                //{
                //    if (dr["rank_name"] != null)
                //    {
                //        if (!StudentTag2Dict.ContainsKey(sid))
                //            StudentTag2Dict.Add(sid, dr["rank_name"].ToString());
                //    }
                //}


                if (!value[sid].ContainsKey(key))
                    value[sid].Add(key, new Dictionary<string, string>());

                foreach (string r2 in r2List)
                {
                    string dValue = "";
                    if (dr[r2] != null)
                    {
                        if (r2ListNP.Contains(r2))
                        {
                            decimal dd;
                            if (decimal.TryParse(dr[r2].ToString(), out dd))
                            {
                                dValue = Math.Round(dd, 2, MidpointRounding.AwayFromZero).ToString();
                            }

                        }
                        else
                        {
                            dValue = dr[r2].ToString();
                        }
                    }

                    if (!value[sid][key].ContainsKey(r2))
                        value[sid][key].Add(r2, dValue);

                }
            }

            return value;
        }


        /// <summary>
        /// 取得班級學期排名、五標與組距資料值
        ///  </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, Dictionary<string, string>>> GetClassSemsScoreRankMatrixDataValue(string SchoolYear, string Semester, List<string> ClassIDList)
        {
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> value = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            // 沒有學生不處理
            if (ClassIDList.Count == 0)
                return value;

            List<string> r2List = new List<string>();
            r2List.Add("matrix_count");
            r2List.Add("avg_top_25");
            r2List.Add("avg_top_50");
            r2List.Add("avg");
            r2List.Add("avg_bottom_50");
            r2List.Add("avg_bottom_25");
            r2List.Add("pr_88");
            r2List.Add("pr_75");
            r2List.Add("pr_50");
            r2List.Add("pr_25");
            r2List.Add("pr_12");
            r2List.Add("std_dev_pop");

            r2List.Add("level_gte100");
            r2List.Add("level_90");
            r2List.Add("level_80");
            r2List.Add("level_70");
            r2List.Add("level_60");
            r2List.Add("level_50");
            r2List.Add("level_40");
            r2List.Add("level_30");
            r2List.Add("level_20");
            r2List.Add("level_10");
            r2List.Add("level_lt10");

            // 需要四捨五入
            List<string> r2ListNP = new List<string>();
            r2ListNP.Add("avg_top_25");
            r2ListNP.Add("avg_top_50");
            r2ListNP.Add("avg");
            r2ListNP.Add("avg_bottom_50");
            r2ListNP.Add("avg_bottom_25");
            r2ListNP.Add("pr_88");
            r2ListNP.Add("pr_75");
            r2ListNP.Add("pr_50");
            r2ListNP.Add("pr_25");
            r2ListNP.Add("pr_12");
            r2ListNP.Add("std_dev_pop");




            QueryHelper qh = new QueryHelper();
            string query = "" +
               " SELECT DISTINCT " +
" 	rank_matrix.id AS rank_matrix_id" +
" 	, rank_matrix.school_year" +
" 	, rank_matrix.semester" +
" 	, rank_matrix.grade_year" +
" 	, rank_matrix.item_type" +
" 	, rank_matrix.item_name" +
" 	, rank_matrix.rank_type" +
" 	, rank_matrix.rank_name" +
" 	, class.id AS class_id " +
" 	, class.class_name" +
"   , rank_matrix.matrix_count " +
"   , rank_matrix.avg_top_25" +
"   , rank_matrix.avg_top_50" +
"   , rank_matrix.avg" +
"   , rank_matrix.avg_bottom_50" +
"   , rank_matrix.avg_bottom_25" +
"   , rank_matrix.pr_88" +
"   , rank_matrix.pr_75" +
"   , rank_matrix.pr_50" +
"   , rank_matrix.pr_25" +
"   , rank_matrix.pr_12" +
"   , rank_matrix.std_dev_pop" +
" 	, rank_matrix.level_gte100" +
" 	, rank_matrix.level_90" +
" 	, rank_matrix.level_80" +
" 	, rank_matrix.level_70" +
" 	, rank_matrix.level_60" +
" 	, rank_matrix.level_50" +
" 	, rank_matrix.level_40" +
" 	, rank_matrix.level_30" +
" 	, rank_matrix.level_20" +
" 	, rank_matrix.level_10" +
" 	, rank_matrix.level_lt10" +
" FROM " +
" 	rank_matrix" +
" 	LEFT OUTER JOIN rank_detail" +
" 		ON rank_detail.ref_matrix_id = rank_matrix.id" +
" 	LEFT OUTER JOIN student" +
" 		ON student.id = rank_detail.ref_student_id" +
" 	LEFT OUTER JOIN class" +
" 		ON class.id = student.ref_class_id" +
" WHERE" +
" 	rank_matrix.is_alive = true" +
" 	AND rank_matrix.school_year = " + SchoolYear +
"     AND rank_matrix.semester = " + Semester +
" 	AND rank_matrix.item_type like '學期%'" +
" 	AND rank_matrix.ref_exam_id = -1 " +
"     AND student.ref_class_id IN (" + string.Join(",", ClassIDList.ToArray()) + ") ";

            DataTable dt = qh.Select(query);
            //dt.TableName = "d5";
            //dt.WriteXmlSchema(Application.StartupPath + "\\d5s.xml");
            //dt.WriteXml(Application.StartupPath + "\\d5d.xml");

            ClassTag1Dict.Clear();
            ClassTag2Dict.Clear();

            // student id key
            // key = item_type + item_name +  rank_name 
            foreach (DataRow dr in dt.Rows)
            {
                string sid = dr["class_id"].ToString();
                if (!value.ContainsKey(sid))
                    value.Add(sid, new Dictionary<string, Dictionary<string, string>>());

                string key = dr["item_type"].ToString() + "_" + dr["item_name"].ToString() + "_" + dr["rank_type"].ToString();

                if (key == "學期/總計成績_課程學習總成績_類別1排名")
                {
                    if (dr["rank_name"] != null)
                    {
                        if (!ClassTag1Dict.ContainsKey(sid))
                            ClassTag1Dict.Add(sid, dr["rank_name"].ToString());

                    }
                }

                if (key == "學期/總計成績_課程學習總成績_類別2排名")
                {
                    if (dr["rank_name"] != null)
                    {
                        if (!ClassTag2Dict.ContainsKey(sid))
                            ClassTag2Dict.Add(sid, dr["rank_name"].ToString());
                    }
                }


                if (!value[sid].ContainsKey(key))
                    value[sid].Add(key, new Dictionary<string, string>());

                foreach (string r2 in r2List)
                {
                    string dValue = "";
                    if (dr[r2] != null)
                    {
                        if (r2ListNP.Contains(r2))
                        {
                            decimal dd;
                            if (decimal.TryParse(dr[r2].ToString(), out dd))
                            {
                                dValue = Math.Round(dd, 2, MidpointRounding.AwayFromZero).ToString();
                            }

                        }
                        else
                        {
                            dValue = dr[r2].ToString();
                        }
                    }

                    if (!value[sid][key].ContainsKey(r2))
                        value[sid][key].Add(r2, dValue);

                }
            }

            return value;
        }


        /// <summary>
        /// 取得科目資料管理
        /// </summary>
        /// <returns></returns>
        public static List<string> GetSubjectList()
        {
            QueryHelper queryHelper = new QueryHelper();
            DataTable dataTable;
            List<string> subjectNameList = new List<string>();
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

            foreach (DataRow dtRow in dataTable.Rows)
            {
                subjectNameList.Add(dtRow["subject_name"] + "");
            }
            return subjectNameList;
        }
    }

}
