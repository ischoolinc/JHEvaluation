using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HsinChuSemesterScoreFixed_JH
{
    public class Global
    {
        /// <summary>
        /// UDT TableName
        /// </summary>
        public const string _UDTTableName = "ischool.HsinChuSemesterScoreFixed_JH.configure";

        public static string _ProjectName = "國中新竹學期成績單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static string _UserConfTypeName = "使用者選擇設定檔";

        public static int _SelSchoolYear;
        public static int _SelSemester;
        public static string _SelExamID = "";
        public static string _SelRefsExamID = "";

        public static List<string> _SelStudentIDList = new List<string>();

        /// <summary>
        /// 設定檔預設名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> DefaultConfigNameList()
        {
            List<string> retVal = new List<string>();
            retVal.Add("領域成績單");
            retVal.Add("科目成績單");
            retVal.Add("科目及領域成績單_領域組距");
            retVal.Add("科目及領域成績單_科目組距");
            return retVal;
        }

        /// <summary>
        /// 取得獎懲名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> GetDisciplineNameList()
        {
            return new string[] { "大功", "小功", "嘉獎", "大過", "小過", "警告" }.ToList();
        }

        /// <summary>
        /// 匯出合併欄位總表Word
        /// </summary>
        public static void ExportMappingFieldWord()
        {

        }
    }
}
