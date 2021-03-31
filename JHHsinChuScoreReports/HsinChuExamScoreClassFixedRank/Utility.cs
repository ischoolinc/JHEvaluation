using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;
using System.IO;

namespace HsinChuExamScoreClassFixedRank
{
    public class Utility
    {
//        /// <summary>
//        /// 取得領域排序
//        /// </summary>
//        /// <returns></returns>
//        public static List<string> GetDominOrder()
//        {
//            List<string> result = new List<string>();
//            QueryHelper qh = new QueryHelper();
//            string sql =
//                @"
//SELECT
//unnest(xpath('/Configurations/Configuration/Domains/Domain/@Name', xmlparse(content replace(  replace(content,'&lt;', '<'),'&gt;', '>')))) as domain_name
//FROM 
//	list
//WHERE  name = 'JHEvaluation_Subject_Ordinal' 

//";
//            DataTable dt = qh.Select(sql);

//            foreach (DataRow dr in dt.Rows)
//            {
//                string domain = dr["domain_name"] + "";
//                result.Add(domain);
//            }
//            return result;
//        }


        /// <summary>
        /// 取得科目排序
        /// </summary>
        /// <returns></returns>
        public static List<string> GetSubjectOrder()
        {
            List<string> result = new List<string>();
            QueryHelper qh = new QueryHelper();
            string sql =
                @"
                    SELECT
                    unnest(xpath('/Configurations/Configuration/Subjects/Subject/@Name', xmlparse(content replace(  replace(content,'&lt;', '<'),'&gt;', '>')))) as subject_name
                    FROM 
	                    list
                    WHERE  name = 'JHEvaluation_Subject_Ordinal'";

            DataTable dt = qh.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                string subject = dr["subject_name"] + "";
                result.Add(subject);
            }
            return result;
        }
    }
}
