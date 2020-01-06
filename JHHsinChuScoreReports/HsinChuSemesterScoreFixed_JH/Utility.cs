using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;
using K12.Data;
using System.IO;

namespace HsinChuSemesterScoreFixed_JH
{
    class Utility
    {
        /// <summary>
        /// 透過 ClassID 取得學生為一般，依照年級、班級名稱、座號排序
        /// </summary>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static List<string> GetClassStudentIDList1ByClassID(List<string> ClassIDList)
        {
            List<string> retVal = new List<string>();
            QueryHelper qh = new QueryHelper();
            string query = "select student.id from student inner join class on student.ref_class_id = class.id where student.status=1 and class.id in(" + string.Join(",", ClassIDList.ToArray()) + ") order by class.grade_year,class.class_name,student.seat_no";
            DataTable dt = new DataTable();
            dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
                retVal.Add(dr[0].ToString());

            dt.Clear();
            return retVal;
        }
    }
}
