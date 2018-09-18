using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;

namespace JHEvaluation.StudentSemesterScoreReport
{
    public class Utility
    {
        /// <summary>
        /// 透過學生編號,取得特定學年度學期服務學習時數
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// /// <returns></returns>
        public static Dictionary<string, decimal> GetServiceLearningDetail(List<string> StudentIDList, int SchoolYear, int Semester)
        {
            Dictionary<string, decimal> retVal = new Dictionary<string, decimal>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,occur_date,reason,hours from $k12.service.learning.record where school_year=" + SchoolYear + " and semester=" + Semester + " and ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "');";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    decimal hr;

                    string sid = dr[0].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, 0);

                    if (decimal.TryParse(dr["hours"].ToString(), out hr))
                        retVal[sid] += hr;
                }
            }
            return retVal;
        }

        //2018.09.16 [ischoolKingdom] Vicky依據[05-02][02]學期成績證明單 服務學習時數顯示處理 項目，新增服務學習學年累計時數。
        /// <summary>
        /// 透過學生編號,取得特定學年度累計服務學習時數
        /// </summary>
        public static Dictionary<string, decimal> GetServiceLearningDetail(List<string> StudentIDList, int SchoolYear)
        {
            Dictionary<string, decimal> _retVal = new Dictionary<string, decimal>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,occur_date,reason,hours from $k12.service.learning.record where school_year=" + SchoolYear  + " and ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "');";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    decimal hr;

                    string sid = dr[0].ToString();
                    if (!_retVal.ContainsKey(sid))
                        _retVal.Add(sid, 0);

                    if (decimal.TryParse(dr["hours"].ToString(), out hr))
                        _retVal[sid] += hr;
                }
            }
            return _retVal;
        }
    }
}
