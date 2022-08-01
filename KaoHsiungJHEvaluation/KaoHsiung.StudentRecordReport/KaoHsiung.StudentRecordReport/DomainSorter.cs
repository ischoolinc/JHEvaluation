using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace KaoHsiung.StudentRecordReport
{
    public class DomainSorter
    {
        //private static List<string> _list = new List<string>(new string[] { "國語文", "英語", "數學", "社會", "藝術與人文", "自然與生活科技", "健康與體育", "綜合活動", "彈性課程" });
        //private static List<string> _list = new List<string>(new string[] { "國語文","英語","數學","社會","自然科學","自然與生活科技","藝術","藝術與人文","健康與體育","綜合活動","科技","實用語文","實用數學","社會適應","生活教育","休閒教育","職業教育","特殊需求","體育專業","藝術才能專長","彈性課程" });
        private static List<string> _list = GetDomainList();
        /// <summary>
        /// 取得資料庫內中英文對照資料
        /// </summary>
        private static List<string> GetDomainList()
        {
            List<string> domainList = new List<string>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string subjectQuery = @"
WITH subject_mapping AS 
(
SELECT
    unnest(xpath('//Subjects/Subject/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS subject_name
	,unnest(xpath('//Subjects/Subject/@EnglishName',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS subject_english_name
FROM  
    list 
WHERE name  ='JHEvaluation_Subject_Ordinal'
)SELECT
		replace (subject_name ,'&amp;amp;','&') AS subject_name
		, replace (subject_english_name ,'&amp;amp;','&') AS subject_english_name
	FROM  subject_mapping";
                //DataTable dt = qh.Select(subjectQuery);

                //foreach (DataRow dr in dt.Rows)
                //{
                //    string subject = dr["subject_name"].ToString();
                //    string subjectEng = dr["subject_english_name"].ToString();
                //    if (!_SubjectNameDict.ContainsKey(subject))
                //        _SubjectNameDict.Add(subject, subjectEng);
                //}


                string domainQuery = @"
WITH domain_mapping AS 
(
	SELECT
		unnest(xpath('//Domains/Domain/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS domain_name
		, unnest(xpath('//Domains/Domain/@EnglishName',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS domain_EnglishName
	FROM  
	    list 
	WHERE name  ='JHEvaluation_Subject_Ordinal'
)SELECT
		replace (domain_name ,'&amp;amp;','&') AS domain_name
		,replace (domain_EnglishName ,'&amp;amp;','&') AS domain_EnglishName
	FROM  domain_mapping";
                DataTable dt2 = qh.Select(domainQuery);

                foreach (DataRow dr in dt2.Rows)
                {
                    string domain = dr["domain_name"].ToString();
                    string domainEng = dr["domain_EnglishName"].ToString();
                    domainList.Add(domain);
                }
            }
            catch (Exception ex)
            {

            }
            return domainList;
        }



        public static int Sort1(string x, string y)
        {
            int ix = _list.IndexOf(x);
            int iy = _list.IndexOf(y);

            if (ix >= 0 && iy >= 0) return ix.CompareTo(iy);
            else if (ix >= 0) return -1;
            else if (iy >= 0) return 1;
            else return x.CompareTo(y);
        }
    }
}
