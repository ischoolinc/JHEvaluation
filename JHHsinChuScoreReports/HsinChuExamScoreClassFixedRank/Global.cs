using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;

namespace HsinChuExamScoreClassFixedRank
{
    public class Global
    {
        public const string _UDTTableName = "ischool.HsinChuExamScoreClassFixedRan.configure";

        public static string _ProjectName = "國中新竹班級評量成績單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static string _UserConfTypeName = "使用者選擇設定檔";

        public static int _SelSchoolYear;
        public static int _SelSemester;
        public static string _SelExamID = "";
        public static List<string> _SelStudentIDList = new List<string>();

        /// <summary>
        /// 設定領域名稱
        /// </summary>
        public static void SetDomainList()
        {
            DomainNameList.Clear();

            // 從學生修課動態取得科目領域名稱
            if (_SelStudentIDList.Count > 0 && _SelSchoolYear > 0 && _SelSemester > 0 && _SelExamID != "")
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "SELECT DISTINCT " +
                    "domain " +
                    "FROM " +
                    "sc_attend " +
                    "INNER JOIN " +
                    "course " +
                    "ON sc_attend.ref_course_id=course.id " +
                    "INNER JOIN te_include " +
                    "ON course.ref_exam_template_id = te_include.ref_exam_template_id " +
                    "WHERE sc_attend.ref_student_id IN(" + string.Join(",", _SelStudentIDList.ToArray()) + ") " +
                    "AND course.school_year=" + _SelSchoolYear + " " +
                    "AND course.semester=" + _SelSemester + " " +
                    "AND te_include.ref_exam_id = " + _SelExamID + " AND domain <>'';";
                DataTable dt = qh.Select(strSQL);

                foreach (DataRow dr in dt.Rows)
                {
                    string domain = dr["domain"].ToString();
                    if (!DomainNameList.Contains(domain))
                        DomainNameList.Add(domain);
                }
            }
            else
            {
                // 預設
                DomainNameList.Add("語文");
                DomainNameList.Add("數學");
                DomainNameList.Add("社會");
                DomainNameList.Add("自然與生活科技");
                DomainNameList.Add("自然科學");
                DomainNameList.Add("藝術");
                DomainNameList.Add("健康與體育");
                DomainNameList.Add("藝術與人文");
                DomainNameList.Add("綜合活動");
                DomainNameList.Add("彈性課程");
                DomainNameList.Add("科技");
                DomainNameList.Add("特殊需求");
            }

            if (!DomainNameList.Contains("彈性課程"))
                DomainNameList.Add("彈性課程");
        }

        /// <summary>
        /// 設定檔預設名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> DefaultConfigNameList()
        {
            List<string> retVal = new List<string>();
            retVal.Add("領域成績單");
            retVal.Add("科目成績單");         
            return retVal;
        }


        /// <summary>
        /// 固定領域名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> DomainNameList = new List<string>();

    }
}
