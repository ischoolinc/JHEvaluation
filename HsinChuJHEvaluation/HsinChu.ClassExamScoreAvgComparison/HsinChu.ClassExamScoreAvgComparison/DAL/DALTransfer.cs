using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace HsinChu.ClassExamScoreAvgComparison.DAL
{
    class DALTransfer
    {
        /// <summary>
        /// 班級依DisplayOrder排序(沒有 DisplayOrder按班級名稱
        /// </summary>
        /// <param name="classRecordList"></param>
        /// <returns></returns>
        public static List<JHSchool.Data.JHClassRecord> ClassRecordSortByDisplayOrder(List<JHSchool.Data.JHClassRecord> classRecordList)
        {
            List<JHSchool.Data.JHClassRecord> HasClassDisplayOrder = new List<JHSchool.Data.JHClassRecord>();
            List<JHSchool.Data.JHClassRecord> NoClassDisplayOrder = new List<JHSchool.Data.JHClassRecord>();
            List<JHSchool.Data.JHClassRecord> returnValue = new List<JHSchool.Data.JHClassRecord>();

            foreach (JHSchool.Data.JHClassRecord cr in classRecordList)
            {
                if (string.IsNullOrEmpty(cr.DisplayOrder))
                    NoClassDisplayOrder.Add(cr);
                else
                    HasClassDisplayOrder.Add(cr);
            }
            HasClassDisplayOrder.Sort(new Comparison<JHSchool.Data.JHClassRecord>(ClassReocrdSorter1));
            NoClassDisplayOrder.Sort(new Comparison<JHSchool.Data.JHClassRecord>(ClassReocrdSorter2));

            foreach (JHSchool.Data.JHClassRecord cr in HasClassDisplayOrder)
                returnValue.Add(cr);

            foreach (JHSchool.Data.JHClassRecord cr in NoClassDisplayOrder)
                returnValue.Add(cr);

            return returnValue;
        }

        private  static int ClassReocrdSorter1(JHSchool.Data.JHClassRecord x, JHSchool.Data.JHClassRecord y)
        {
            int intX;
            int intY;
            int.TryParse(x.DisplayOrder , out intX);
            int.TryParse(y.DisplayOrder, out intY);
            return intX.CompareTo(intY);
        }

        private static int ClassReocrdSorter2(JHSchool.Data.JHClassRecord x, JHSchool.Data.JHClassRecord y)
        {
            return x.Name.CompareTo(y.Name);        
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
                subjectNameList.Add(dtRow["subject_name"].ToString());
            }
            return subjectNameList;
        }
    }
}
