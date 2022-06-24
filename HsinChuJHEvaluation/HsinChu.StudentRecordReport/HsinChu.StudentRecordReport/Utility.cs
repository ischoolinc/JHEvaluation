using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;

namespace HsinChu.StudentRecordReport
{
    public class Utility
    {
        /// <summary>
        /// 透過學生編號,取得特定學年度學期服務學習時數
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// /// <returns></returns>
        public static Dictionary<string, Dictionary<string, string>> GetServiceLearningDetail(List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, string>> retVal = new Dictionary<string, Dictionary<string, string>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,school_year,semester,sum(hours) as hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') group by ref_student_id,school_year,semester order by school_year,semester;";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["ref_student_id"].ToString();
                    string key1 = dr["school_year"].ToString() + "_" + dr["semester"].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, new Dictionary<string, string>());

                    if (!retVal[sid].ContainsKey(key1))
                        retVal[sid].Add(key1, "0");

                    retVal[sid][key1] = dr["hours"].ToString();
                }
            }
            return retVal;
        }

        /// <summary>
        /// 科目排序
        /// </summary>
        /// <returns></returns>
        public static List<string> GetSubjectOrder()
        {
            List<string> result = new List<string>();
            QueryHelper qh = new QueryHelper();
            string sql =
                @"
WITH subject_mapping AS 
(
SELECT
    unnest(xpath('//Subjects/Subject/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS subject_name
FROM  
    list 
WHERE name  ='JHEvaluation_Subject_Ordinal'
)SELECT
		replace (subject_name ,'&amp;amp;','&') AS subject_name
	FROM  subject_mapping
";

            DataTable dt = qh.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                string subject = dr["subject_name"].ToString();
                result.Add(subject);
            }
            return result;
        }

    }
}
