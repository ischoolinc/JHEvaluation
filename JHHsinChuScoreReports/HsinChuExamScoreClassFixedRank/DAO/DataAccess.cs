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
        /// 透過班級ID、學年度、學期 取得試別、領域、科目
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static Dictionary<string,Dictionary<string,List<string>>> GetExamDomainSubjectDictByClass(string SchoolYear,string Semester,List<string> ClassIDList)
        {
            Dictionary<string, Dictionary<string, List<string>>> value = new Dictionary<string, Dictionary<string, List<string>>>();

            if (ClassIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "SELECT DISTINCT te_include.ref_exam_id AS exam_id,course.domain,course.subject FROM sc_attend INNER JOIN course ON sc_attend.ref_course_id = course.id INNER JOIN student ON sc_attend.ref_student_id = student.id  INNER JOIN te_include ON course.ref_exam_template_id = te_include.ref_exam_template_id WHERE student.status = 1 AND student.ref_class_id IN("+string.Join(",",ClassIDList.ToArray())+") AND course.school_year = "+SchoolYear+" AND course.semester = "+Semester+" ORDER BY exam_id,domain,subject";

                DataTable dt = qh.Select(query);
                foreach(DataRow dr in dt.Rows)
                {
                    string exam_id = dr["exam_id"].ToString();
                    // 領域空白為彈性課程
                    string domain = "彈性課程";
                    if (dr["domain"] != null && dr["domain"].ToString() != "")
                    {
                        domain = dr["domain"].ToString();
                    }

                    string subject = dr["subject"].ToString();

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
    }
}
