using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Xml.Linq;
using System.Data;

namespace JHEvaluation.StudentScoreSummaryReport
{
    /// <summary>
    /// 領域與科目中英文對照
    /// </summary>
    public class SubjDomainEngNameMapping
    {
        private Dictionary<string, string> _DomainNameDict;
        private Dictionary<string, string> _SubjectNameDict;

        public SubjDomainEngNameMapping()
        {
            _DomainNameDict = new Dictionary<string, string>();
            _SubjectNameDict = new Dictionary<string, string>();
            GetData();
        }

        /// <summary>
        /// 取得資料庫內中英文對照資料
        /// </summary>
        private void GetData()
        {
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
                DataTable dt = qh.Select(subjectQuery);
                
                foreach (DataRow dr in dt.Rows)
                {
                    string subject = dr["subject_name"].ToString();
                    string subjectEng = dr["subject_english_name"].ToString();
                    if (!_SubjectNameDict.ContainsKey(subject))
                        _SubjectNameDict.Add(subject,subjectEng);
                }


                string domainQuery = @"
WITH  domain_mapping AS 
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
                    if (!_DomainNameDict.ContainsKey(domain))
                        _DomainNameDict.Add(domain, domainEng);
                }
            }
            catch (Exception ex)
            { 
                
            }

        }


        private string GetAttribute(string name, XElement elm)
        {
            string retVal = "";
            if (elm.Attribute(name) != null)
                retVal = elm.Attribute(name).Value;

            return retVal;
        }

        /// <summary>
        /// 取得領域英文名稱
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public string GetDomainEngName(string name)
        {
            string retVal = "";
            if (_DomainNameDict.ContainsKey(name))
                retVal = _DomainNameDict[name];
            return retVal;
        }

        /// <summary>
        /// 取得科目英文名稱
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public string GetSubjectEngName(string name)
        {
            string retVal = "";
            if (_SubjectNameDict.ContainsKey(name))
                retVal = _SubjectNameDict[name];
            return retVal;
        }
    }
}
