using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using FISCA.Data;
using System.Data;
using Aspose.Words;
using HsinChuExamScore_JH.DAO;

namespace HsinChuExamScore_JH
{
    public class Global
    {
        #region 設定檔記錄用

        /// <summary>
        /// UDT TableName
        /// </summary>
        public const string _UDTTableName = "ischool.新竹國中評量成績通知單.configure";

        public static string _ProjectName = "國中新竹評量成績單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static string _UserConfTypeName = "使用者選擇設定檔";

        public static int _SelSchoolYear;
        public static int _SelSemester;
        public static string _SelExamID = "";
        public static string _SelRefsExamID = "";

        public static List<string> _SelStudentIDList = new List<string>();
        // public static List<string> UserDefineFields = new List<string>();
        /// <summary>
        /// 設定檔預設名稱 
        /// </summary>
        /// <returns></returns>

        /// <summary>
        /// 進位四捨五入位數
        /// </summary>
        public static int parseNumebr = 2;

        public static List<string> DefaultConfigNameList()
        {
            List<string> retVal = new List<string>();
            retVal.Add("領域成績單");
            retVal.Add("科目成績單");
            retVal.Add("科目及領域成績單_領域組距");
            retVal.Add("科目及領域成績單_科目組距");
            return retVal;
        }

        #endregion


        /// <summary>
        /// 用傳入字串 去確認是否有含彈性課程 
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static Boolean CheckIfContainFlex(string text)
        {
            if (text.Contains("不含彈性"))
            {
                return false;
            }
            else
            {
                return true;
            }
        }



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

                DomainNameList.Sort(new StringComparer(Utility.GetDominOrder().ToArray()));
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
        /// 固定領域名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> DomainNameList = new List<string>();

        /// <summary>
        /// 取得獎懲名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> GetDisciplineNameList()
        {
            return new string[] { "大功", "小功", "嘉獎", "大過", "小過", "警告" }.ToList();
        }

        /// <summary>
        /// 取得自訂欄位
        /// </summary>
        public static List<string> GetUserDefineFields()
        {

            List<string> UserDefineFields = new List<string>();
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select("SELECT fieldname FROM $stud.userdefinedata GROUP BY fieldname");
            foreach (DataRow dr in dt.Rows)
            {
                UserDefineFields.Add(dr["fieldname"].ToString());
            }
            return UserDefineFields;
        }


        /// <summary>
        /// 匯出合併欄位總表Word 註:20200318為符合南榮自訂欄位需求增加自訂欄位變數
        /// </summary>
        public static void ExportMappingFieldWord()
        {
            #region 儲存檔案
            string inputReportName = "國中評量成績單合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.新竹_評量成績合併欄位總表_動態版));

            try
            {
                #region 動態產生合併欄位
                // 讀取總表檔案並動態加入合併欄位

                List<string> ScoreComposition = new List<string>();
                ScoreComposition.Add("科目成績");
                ScoreComposition.Add("科目定期成績");
                ScoreComposition.Add("參考科目成績");
                ScoreComposition.Add("參考科目定期成績");

                DocumentBuilder builder = new DocumentBuilder(tempDoc);
                builder.Write("=== 國中評量成績單合併欄位總表 ===");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("學生電子報表識別");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 系統編號" + " \\* MERGEFORMAT ", "«學生電子報表識別" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("學校名稱");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學校名稱" + " \\* MERGEFORMAT ", "«學校名稱" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("學年度");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學年度" + " \\* MERGEFORMAT ", "«學年度" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("學期");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學期" + " \\* MERGEFORMAT ", "«學期" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("試別名稱");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 試別名稱" + " \\* MERGEFORMAT ", "«試別名稱" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("參考試別");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 參考試別" + " \\* MERGEFORMAT ", "«參考試別" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("班級");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 班級" + " \\* MERGEFORMAT ", "«班級" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("學號");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + " \\* MERGEFORMAT ", "«學號" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("座號");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + " \\* MERGEFORMAT ", "«座號" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("姓名");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + " \\* MERGEFORMAT ", "«姓名" + "»");
                builder.EndRow();

                // Jean 新增自訂欄位 *不存在記憶體裡因為 這裡是靜態空間 如行政人員及時增加自訂欄位 會抓不到最新的
                foreach (string userDefine in GetUserDefineFields())

                {
                    builder.InsertCell();
                    builder.Write(userDefine + "-自訂欄位");
                    builder.InsertCell();
                    builder.InsertField($"MERGEFIELD {userDefine}-自訂欄位  \\* MERGEFORMAT ", $"«{userDefine}-自訂欄位»");
                    builder.EndRow();
                }

                // Jean增加總計成績 算術平均 算數總分
                
                string[] itemNames = new string[] { "平均", "總分" };
                Dictionary<string, string> itemTypesMapping = new Dictionary<string, string>() {
                        {"定期評量_定期/總計成績","成績"},
                        {"定期評量/總計成績","定期成績"}}; // 因資料庫內 itemTypes用詞過長 轉換為較短並符合其他合併欄位命名方式的欄位

                string[] martrixCols = new string[] { "名次", "PR值", "百分比", "母體頂標", "母體前標", "母體平均", "母體後標", "母體底標", "母體人數", "母體新頂標", "母體新前標", "母體新均標", "母體新後標", "母體新底標", "母體標準差" }; // 固定排名統計值
                string[] _rankTypes = { "班排名", "年排名", "類別1排名", "類別2排名" };

            
                //  原本的變數改成用迴圈寫 

                string[] _itemNames = new string[] { "加權平均", "加權總分" };
                //string[] _rankTypes = new string[] { "班排名", "年排名", "類別1排名", "類別2排名" };
                string[] matrixColumns = new string[] { "名次", "PR值", "百分比", "母體頂標", "母體前標", "母體平均", "母體後標", "母體底標", "母體人數", "母體新頂標", "母體新前標", "母體新均標", "母體新後標", "母體新底標", "母體標準差" }; // 固定排名統計值


             


                // 科目加權平均
                foreach (string composition in ScoreComposition) //  1. 科目成績   2.科目定期成績  3.參考科目成績  4.參考科目定期成績
                {
                                                         
                    foreach (string itemName in itemNames) //平均 總分 => 基本上這兩個 就是從科目來的
                    {

                          
                        //if (!composition.Contains("參考"))
                        //{
                        //    builder.InsertCell();
                        //    builder.Write($"{composition}{itemName}");
                        //    builder.InsertCell();
                        //    builder.InsertField($"MERGEFIELD  {composition}{itemName}" + " \\* MERGEFORMAT ", $"« {composition}{itemName}" + "»");
                        //    builder.EndRow();
                        //}

                        foreach (string rankType in _rankTypes) // 年排名 班排名 類別1排名 類別2排名 
                        {
                            builder.InsertCell();
                            builder.Write($"------------【{composition} {itemName} {rankType}】------------");
                            builder.InsertCell();
                            builder.Write($"---------------------------------------------------------------");
                            builder.EndRow();

                            // foreach (string itemType in itemTypesMapping.Keys)
                            {  // 處理成績
                                foreach (string matrixCol in martrixCols)
                                {
                                    builder.InsertCell();
                                    builder.Write($"{composition}{itemName}{rankType}{matrixCol}");
                                    builder.InsertCell();
                                    builder.InsertField($"MERGEFIELD {composition}{itemName}{rankType}{matrixCol}   \\* MERGEFORMAT ", $"«{composition}{itemName}{rankType}{matrixCol}»");
                                    builder.EndRow();
                                }
                            }
                       
                        }
                                                                     
                    }
                    
                    // 加權平均
                    foreach (string itemName in _itemNames)  //"加權平均" ,  "加權總分" 
                    {
                        // 處理成績
                        //if (!composition.Contains("參考"))
                        //{
                        //    builder.InsertCell();
                        //    builder.Write($"{composition}{itemName}");
                        //    builder.InsertCell();
                        //    builder.InsertField($"MERGEFIELD  {composition}{itemName}" + " \\* MERGEFORMAT ", $"«{composition}{itemName}" + "»");
                        //    builder.EndRow();
                        //}
                        // 處理 固定排名相關資料
                        foreach (string ranktype in _rankTypes)  //"班排名", "年排名" , "類別1排名" , "類別2排名"
                        {

                            builder.InsertCell();
                            builder.Write($"------------【{composition} {itemName} {ranktype}】------------");
                            builder.InsertCell();
                            builder.Write($"--------------------------------------------------------------");
                            builder.EndRow();


                            foreach (string matrixCol in matrixColumns) // "名次", "PR值", "百分比" , "母體頂標" , "母體前標",  "母體平均", "母體後標" , "母體底標" , "母體人數" 
                            {

                                // if (!rankType.Contains("參考"))
                                {
                                    builder.InsertCell();
                                    builder.Write($"{composition}{itemName}{ranktype}{matrixCol}");
                                    builder.InsertCell();
                                    builder.InsertField($"MERGEFIELD  {composition}{itemName}{ranktype}{matrixCol}" + " \\* MERGEFORMAT ", $"«{composition}{itemName}{ranktype}{matrixCol}" + "»");
                                    builder.EndRow();
                                }
                            }
                        }
                    }
                
                    // 領域加權平均
                    //    if (!rt.Contains("參考"))
                    //    {
                    //        builder.InsertCell();
                    //        builder.Write(rt + "加權平均");
                    //        builder.InsertCell();
                    //        builder.InsertField("MERGEFIELD " + rt + "加權平均" + " \\* MERGEFORMAT ", "«DA" + "»");
                    //        builder.EndRow();
                    //    }

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名名次" + " \\* MERGEFORMAT ", "«DACR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名PR值" + " \\* MERGEFORMAT ", "«DACPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名百分比" + " \\* MERGEFORMAT ", "«DACP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名母體頂標" + " \\* MERGEFORMAT ", "«DAC25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名母體前標" + " \\* MERGEFORMAT ", "«DAC50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名母體平均" + " \\* MERGEFORMAT ", "«DACA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名母體後標" + " \\* MERGEFORMAT ", "«DAC50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名母體底標" + " \\* MERGEFORMAT ", "«DAC25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均班排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均班排名母體人數" + " \\* MERGEFORMAT ", "«DACMC" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名名次" + " \\* MERGEFORMAT ", "«DAYR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名PR值" + " \\* MERGEFORMAT ", "«DAYPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名百分比" + " \\* MERGEFORMAT ", "«DAYP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名母體頂標" + " \\* MERGEFORMAT ", "«DAY25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名母體前標" + " \\* MERGEFORMAT ", "«DAY50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名母體平均" + " \\* MERGEFORMAT ", "«DAYA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名母體後標" + " \\* MERGEFORMAT ", "«DAY50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名母體底標" + " \\* MERGEFORMAT ", "«DAY25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均年排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均年排名母體人數" + " \\* MERGEFORMAT ", "«DAYMC" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名名次" + " \\* MERGEFORMAT ", "«D1TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名PR值" + " \\* MERGEFORMAT ", "«D1TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名百分比" + " \\* MERGEFORMAT ", "«D1TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名母體頂標" + " \\* MERGEFORMAT ", "«DA1T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名母體前標" + " \\* MERGEFORMAT ", "«DA1T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名母體平均" + " \\* MERGEFORMAT ", "«DA1TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名母體後標" + " \\* MERGEFORMAT ", "«DA1T50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名母體底標" + " \\* MERGEFORMAT ", "«DA1T25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別1排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別1排名母體人數" + " \\* MERGEFORMAT ", "«DA1TMC" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名名次" + " \\* MERGEFORMAT ", "«D2TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名PR值" + " \\* MERGEFORMAT ", "«D2TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名百分比" + " \\* MERGEFORMAT ", "«D2TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名母體頂標" + " \\* MERGEFORMAT ", "«DA2T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名母體前標" + " \\* MERGEFORMAT ", "«DA2T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名母體平均" + " \\* MERGEFORMAT ", "«DA2TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名母體後標" + " \\* MERGEFORMAT ", "«DA2T50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名母體底標" + " \\* MERGEFORMAT ", "«DA2T25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權平均類別2排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權平均類別2排名母體人數" + " \\* MERGEFORMAT ", "«DA2T2MC" + "»");
                    //    builder.EndRow();


                    //    if (!rt.Contains("參考"))
                    //    {

                    //        // 平均 
                    //        builder.InsertCell();
                    //        builder.Write(rt + "平均");
                    //        builder.InsertCell();
                    //        builder.InsertField("MERGEFIELD " + rt + "平均" + " \\* MERGEFORMAT ", "«DA" + "»");
                    //        builder.EndRow();
                    //    }

                    //    //因固定排名沒有計算領域總計成績(平均、總分、加權平均、加權總分) ====> 所以先拿掉
                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名名次" + " \\* MERGEFORMAT ", "«DACR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名PR值" + " \\* MERGEFORMAT ", "«DACPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名百分比" + " \\* MERGEFORMAT ", "«DACP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名母體頂標" + " \\* MERGEFORMAT ", "«DAC25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名母體前標" + " \\* MERGEFORMAT ", "«DAC50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名母體平均" + " \\* MERGEFORMAT ", "«DACA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名母體後標" + " \\* MERGEFORMAT ", "«DAC50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名母體底標" + " \\* MERGEFORMAT ", "«DAC25B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均班排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均班排名母體人數" + " \\* MERGEFORMAT ", "«DACMC" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名名次" + " \\* MERGEFORMAT ", "«DAYR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名PR值" + " \\* MERGEFORMAT ", "«DAYPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名百分比" + " \\* MERGEFORMAT ", "«DAYP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名母體頂標" + " \\* MERGEFORMAT ", "«DAY25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名母體前標" + " \\* MERGEFORMAT ", "«DAY50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名母體平均" + " \\* MERGEFORMAT ", "«DAYA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名母體後標" + " \\* MERGEFORMAT ", "«DAY50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名母體底標" + " \\* MERGEFORMAT ", "«DAY25B" + "»");
                    //    builder.EndRow();



                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均年排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均年排名母體人數" + " \\* MERGEFORMAT ", "«DAYMC" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名名次" + " \\* MERGEFORMAT ", "«D1TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名PR值" + " \\* MERGEFORMAT ", "«D1TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名百分比" + " \\* MERGEFORMAT ", "«D1TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名母體頂標" + " \\* MERGEFORMAT ", "«DA1T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名母體前標" + " \\* MERGEFORMAT ", "«DA1T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名母體平均" + " \\* MERGEFORMAT ", "«DA1TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名母體後標" + " \\* MERGEFORMAT ", "«DA1T50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名母體底標" + " \\* MERGEFORMAT ", "«DA1T25B" + "»");
                    //    builder.EndRow();




                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別1排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別1排名母體人數" + " \\* MERGEFORMAT ", "«DA1TMC" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名名次" + " \\* MERGEFORMAT ", "«D2TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名PR值" + " \\* MERGEFORMAT ", "«D2TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名百分比" + " \\* MERGEFORMAT ", "«D2TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名母體頂標" + " \\* MERGEFORMAT ", "«DA2T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名母體前標" + " \\* MERGEFORMAT ", "«DA2T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名母體平均" + " \\* MERGEFORMAT ", "«DA2TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名母體後標" + " \\* MERGEFORMAT ", "«DA2T50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名母體底標" + " \\* MERGEFORMAT ", "«DA2T25B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "平均類別2排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "平均類別2排名母體人數" + " \\* MERGEFORMAT ", "«DA2TMC" + "»");
                    //    builder.EndRow();

                    //    if (!rt.Contains("參考"))
                    //    {
                    //        // 加權總分
                    //        builder.InsertCell();
                    //        builder.Write(rt + "加權總分");
                    //        builder.InsertCell();
                    //        builder.InsertField("MERGEFIELD " + rt + "加權總分" + " \\* MERGEFORMAT ", "«DA" + "»");
                    //        builder.EndRow();
                    //    }

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名名次" + " \\* MERGEFORMAT ", "«DACR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名PR值" + " \\* MERGEFORMAT ", "«DACPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名百分比" + " \\* MERGEFORMAT ", "«DACP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名母體頂標" + " \\* MERGEFORMAT ", "«DAC25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名母體前標" + " \\* MERGEFORMAT ", "«DAC50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名母體平均" + " \\* MERGEFORMAT ", "«DACA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名母體後標" + " \\* MERGEFORMAT ", "«DAC50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名母體底標" + " \\* MERGEFORMAT ", "«DAC25B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分班排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分班排名母體人數" + " \\* MERGEFORMAT ", "«DAC25B" + "»");
                    //    builder.EndRow();



                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名名次" + " \\* MERGEFORMAT ", "«DAYR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名PR值" + " \\* MERGEFORMAT ", "«DAYPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名百分比" + " \\* MERGEFORMAT ", "«DAYP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名母體頂標" + " \\* MERGEFORMAT ", "«DAY25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名母體前標" + " \\* MERGEFORMAT ", "«DAY50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名母體平均" + " \\* MERGEFORMAT ", "«DAYA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名母體後標" + " \\* MERGEFORMAT ", "«DAY50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名母體底標" + " \\* MERGEFORMAT ", "«DAY25B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分年排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分年排名母體人數" + " \\* MERGEFORMAT ", "«DAYMC" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名名次" + " \\* MERGEFORMAT ", "«D1TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名PR值" + " \\* MERGEFORMAT ", "«D1TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名百分比" + " \\* MERGEFORMAT ", "«D1TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名母體頂標" + " \\* MERGEFORMAT ", "«DA1T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名母體前標" + " \\* MERGEFORMAT ", "«DA1T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名母體平均" + " \\* MERGEFORMAT ", "«DA1TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名母體後標" + " \\* MERGEFORMAT ", "«DA1T50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名母體底標" + " \\* MERGEFORMAT ", "«DA1T25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別1排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別1排名母體人數" + " \\* MERGEFORMAT ", "«DA1TMC" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名名次" + " \\* MERGEFORMAT ", "«D2TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名PR值" + " \\* MERGEFORMAT ", "«D2TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名百分比" + " \\* MERGEFORMAT ", "«D2TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名母體頂標" + " \\* MERGEFORMAT ", "«DA2T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名母體前標" + " \\* MERGEFORMAT ", "«DA2T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名母體平均" + " \\* MERGEFORMAT ", "«DA2TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名母體後標" + " \\* MERGEFORMAT ", "«DA2T50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名母體底標" + " \\* MERGEFORMAT ", "«DA2T25B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "加權總分類別2排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "加權總分類別2排名母體人數" + " \\* MERGEFORMAT ", "«DA2TMC" + "»");
                    //    builder.EndRow();


                    //    if (!rt.Contains("參考"))
                    //    {
                    //        // "總分
                    //        builder.InsertCell();
                    //        builder.Write(rt + "總分");
                    //        builder.InsertCell();
                    //        builder.InsertField("MERGEFIELD " + rt + "總分" + " \\* MERGEFORMAT ", "«DA" + "»");
                    //        builder.EndRow();
                    //    }

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名名次" + " \\* MERGEFORMAT ", "«DACR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名PR值" + " \\* MERGEFORMAT ", "«DACPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名百分比" + " \\* MERGEFORMAT ", "«DACP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名母體頂標" + " \\* MERGEFORMAT ", "«DAC25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名母體前標" + " \\* MERGEFORMAT ", "«DAC50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名母體平均" + " \\* MERGEFORMAT ", "«DACA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名母體後標" + " \\* MERGEFORMAT ", "«DAC50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名母體底標" + " \\* MERGEFORMAT ", "«DAC25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分班排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分班排名母體人數" + " \\* MERGEFORMAT ", "«DAC25B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名名次" + " \\* MERGEFORMAT ", "«DAYR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名PR值" + " \\* MERGEFORMAT ", "«DAYPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名百分比" + " \\* MERGEFORMAT ", "«DAYP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名母體頂標" + " \\* MERGEFORMAT ", "«DAY25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名母體前標" + " \\* MERGEFORMAT ", "«DAY50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名母體平均" + " \\* MERGEFORMAT ", "«DAYA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名母體後標" + " \\* MERGEFORMAT ", "«DAY50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名母體底標" + " \\* MERGEFORMAT ", "«DAY25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分年排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分年排名母體人數" + " \\* MERGEFORMAT ", "«DAYMC" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名名次" + " \\* MERGEFORMAT ", "«D1TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名PR值" + " \\* MERGEFORMAT ", "«D1TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名百分比" + " \\* MERGEFORMAT ", "«D1TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名母體頂標" + " \\* MERGEFORMAT ", "«DA1T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名母體前標" + " \\* MERGEFORMAT ", "«DA1T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名母體平均" + " \\* MERGEFORMAT ", "«DA1TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名母體後標" + " \\* MERGEFORMAT ", "«DA1T50B " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名母體底標" + " \\* MERGEFORMAT ", "«DA1T25B" + "»");
                    //    builder.EndRow();


                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別1排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別1排名母體人數" + " \\* MERGEFORMAT ", "«DA1TMC" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名名次");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名名次" + " \\* MERGEFORMAT ", "«D2TR " + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名PR值");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名PR值" + " \\* MERGEFORMAT ", "«D2TPR" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名百分比");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名百分比" + " \\* MERGEFORMAT ", "«D2TP" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名母體頂標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名母體頂標" + " \\* MERGEFORMAT ", "«DA2T25T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名母體前標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名母體前標" + " \\* MERGEFORMAT ", "«DA2T50T" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名母體平均");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名母體平均" + " \\* MERGEFORMAT ", "«DA2TA" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名母體後標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名母體後標" + " \\* MERGEFORMAT ", "«DA2T50B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名母體底標");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名母體底標" + " \\* MERGEFORMAT ", "«DA2T25B" + "»");
                    //    builder.EndRow();

                    //    builder.InsertCell();
                    //    builder.Write(rt + "總分類別2排名母體人數");
                    //    builder.InsertCell();
                    //    builder.InsertField("MERGEFIELD " + rt + "總分類別2排名母體人數" + " \\* MERGEFORMAT ", "«DA2TMC" + "»");
                    //    builder.EndRow();
                }



                // todo  2020 將部分科目相關欄位改寫為 迴圈下去跑 Jean 
                // 科目相關資訊
                // * 科目
                // 1.
                // a. 加權平均 b.加權平均(不含彈性) c.加權總分 d .加權總分(不含彈性)

                /*===========================以下為原本有的先保留==============================*/

                builder.InsertCell();
                builder.Write($"--------------------------------------------------------------");
                builder.InsertCell();
                builder.Write($"--------------------------------------------------------------");
                builder.EndRow();



                builder.InsertCell();
                builder.Write("領域成績加權平均(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 領域成績加權平均(不含彈性)" + " \\* MERGEFORMAT ", "«DAN" + "»");
                builder.EndRow();


                builder.InsertCell();
                builder.Write("領域成績加權總分(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 領域成績加權總分(不含彈性)" + " \\* MERGEFORMAT ", "«DAN" + "»");
                builder.EndRow();

                /*=================================以上為原本有的先保留================================*/




                /*===========================以上科目相關改為迴圈下去跑==================================*/


                // 個總排列組合
                // AM是算術平均的縮寫
                string scoreTarget = "科目_S";
                string[] ScoreCompositions = new string[] { "定期成績_F", "平時成績_A", "成績_T" };
                string[] ScoreCaculateWays = new string[] { "平均_AMCF", "平均(不含彈性)_AMNF", "總分_ATCF", "總分(不含彈性)_STNF" };

                string scoreTargetName = scoreTarget.Split('_')[0];
                string scoreTargetAbbr = scoreTarget.Split('_')[1];

                foreach (string scoreComposition in ScoreCompositions)
                {
                    string scoreCompositionName = scoreComposition.Split('_')[0];
                    string scoreCompositionAbbr = scoreComposition.Split('_')[1];
                    foreach (string scoreCaculateWay in ScoreCaculateWays)
                    {
                        string scoreCaculateWayName = scoreCaculateWay.Split('_')[0];
                        string scoreCaculateWayAbbr = scoreCaculateWay.Split('_')[1];

                        builder.InsertCell();
                        // builder.Write("科目定期評量加權平均");
                        builder.Write($"{scoreTargetName}{scoreCompositionName}{scoreCaculateWayName}");
                        builder.InsertCell();
                        // builder.InsertField("MERGEFIELD 科目定期評量加權平均" + " \\* MERGEFORMAT ", "«SF" + "»");
                        builder.InsertField($"MERGEFIELD {scoreTargetName}{scoreCompositionName}{scoreCaculateWayName}" + " \\* MERGEFORMAT ", $"«{scoreTargetAbbr}{scoreCaculateWayAbbr}»");
                        builder.EndRow();
                    }
                }




                builder.InsertCell();
                builder.Write("科目定期成績加權平均");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目定期成績加權平均" + " \\* MERGEFORMAT ", "«SF" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目平時成績加權平均");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目平時成績加權平均" + " \\* MERGEFORMAT ", "«SA" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目成績加權平均");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目成績加權平均" + " \\* MERGEFORMAT ", "«ST" + "»");
                builder.EndRow();


                builder.InsertCell();
                builder.Write("科目定期成績加權平均(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目定期成績加權平均(不含彈性)" + " \\* MERGEFORMAT ", "«SFN" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目平時成績加權平均(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目平時成績加權平均(不含彈性)" + " \\* MERGEFORMAT ", "«SAN" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目成績加權平均(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目成績加權平均(不含彈性)" + " \\* MERGEFORMAT ", "«STN" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("領域成績加權總分");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 領域成績加權總分" + " \\* MERGEFORMAT ", "«DA" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目定期成績加權總分");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目定期成績加權總分" + " \\* MERGEFORMAT ", "«SF" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目平時成績加權總分");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目平時成績加權總分" + " \\* MERGEFORMAT ", "«SA" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目成績加權總分");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目成績加權總分" + " \\* MERGEFORMAT ", "«ST" + "»");
                builder.EndRow();


                builder.InsertCell();
                builder.Write("科目定期成績加權總分(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目定期成績加權總分(不含彈性)" + " \\* MERGEFORMAT ", "«SFN" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目平時成績加權總分(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目平時成績加權總分(不含彈性)" + " \\* MERGEFORMAT ", "«SAN" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("科目總成績加權總分(不含彈性課程)");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目成績加權總分(不含彈性)" + " \\* MERGEFORMAT ", "«STN" + "»");
                builder.EndRow();

                /*===========================以上科目相關改為迴圈下去跑==============================*/



                builder.InsertCell();
                builder.Write("缺曠紀錄");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 缺曠紀錄" + " \\* MERGEFORMAT ", "«缺曠紀錄" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("服務學習時數");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 服務學習時數" + " \\* MERGEFORMAT ", "«SL" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("校長");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 校長" + " \\* MERGEFORMAT ", "«校長" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("教務主任");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 教務主任" + " \\* MERGEFORMAT ", "«教務主任" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("班導師");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 班導師" + " \\* MERGEFORMAT ", "«班導師" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("成績校正日期");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 成績校正日期" + " \\* MERGEFORMAT ", "«成績校正日期" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("區間開始日期");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 區間開始日期" + " \\* MERGEFORMAT ", "«區間開始日期" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("區間結束日期");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 區間結束日期" + " \\* MERGEFORMAT ", "«區間結束日期" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("大功區間統計");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 大功區間統計" + " \\* MERGEFORMAT ", "«D1A" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("小功區間統計");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 小功區間統計" + " \\* MERGEFORMAT ", "«D1B" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("嘉獎區間統計");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 嘉獎區間統計" + " \\* MERGEFORMAT ", "«D1C" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("大過區間統計");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 大過區間統計" + " \\* MERGEFORMAT ", "«D2A" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("小過區間統計");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 小過區間統計" + " \\* MERGEFORMAT ", "«D2B" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("警告區間統計");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 警告區間統計" + " \\* MERGEFORMAT ", "«D2C" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("監護人姓名");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 監護人姓名" + " \\* MERGEFORMAT ", "«監護人姓名" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("父親姓名");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 父親姓名" + " \\* MERGEFORMAT ", "«父親姓名" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("母親姓名");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 母親姓名" + " \\* MERGEFORMAT ", "«母親姓名" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("戶籍地址");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 戶籍地址" + " \\* MERGEFORMAT ", "«戶籍地址" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("聯絡地址");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 聯絡地址" + " \\* MERGEFORMAT ", "«聯絡地址" + "»");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("其他地址");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 其他地址" + " \\* MERGEFORMAT ", "«其他地址" + "»");
                builder.EndRow();
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();

                // 處理領域成績
                builder.Write("領域成績");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("領域加權平均");
                builder.InsertCell();
                builder.Write("領域定期加權平均");
                builder.InsertCell();
                builder.Write("領域平時加權平均");
                builder.InsertCell();
                builder.Write("領域加權平均等第");
                builder.InsertCell();
                builder.Write("領域定期加權平均等第");
                builder.InsertCell();
                builder.Write("領域平時加權平均等第");
                builder.InsertCell();
                builder.Write("領域權數");
                builder.InsertCell();
                builder.Write("領域班排名名次");
                builder.InsertCell();
                builder.Write("領域班排名PR值");
                builder.InsertCell();
                builder.Write("領域班排名百分比");
                builder.InsertCell();
                builder.Write("領域班排名母體平均");
                builder.InsertCell();
                builder.Write("領域年排名名次");
                builder.InsertCell();
                builder.Write("領域年排名PR值");
                builder.InsertCell();
                builder.Write("領域年排名百分比");

                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域加權平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DS" + "»");
                    mName1 = domainName + "_領域定期加權平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DF" + "»");
                    mName1 = domainName + "_領域平時加權平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DA" + "»");
                    mName1 = domainName + "_領域加權平均等第";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DS" + "»");
                    mName1 = domainName + "_領域定期加權平均等第";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DF" + "»");
                    mName1 = domainName + "_領域平時加權平均等第";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DA" + "»");
                    mName1 = domainName + "_領域權數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC" + "»");
                    mName1 = domainName + "_領域班排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCR" + "»");
                    mName1 = domainName + "_領域班排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR" + "»");
                    mName1 = domainName + "_領域班排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCP" + "»");
                    mName1 = domainName + "_領域班排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCA" + "»");
                    mName1 = domainName + "_領域年排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYR" + "»");
                    mName1 = domainName + "_領域年排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR" + "»");
                    mName1 = domainName + "_領域年排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYP" + "»");

                    builder.EndRow();
                }
                builder.EndTable(); // 領域成績

                builder.Writeln();
                builder.Writeln();

                // 領域成績排名1
                builder.Write("領域成績排名");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("領域年排名名次");
                builder.InsertCell();
                builder.Write("領域年排名PR值");
                builder.InsertCell();
                builder.Write("領域年排名百分比");
                builder.InsertCell();
                builder.Write("領域班排名名次");
                builder.InsertCell();
                builder.Write("領域班排名PR值");
                builder.InsertCell();
                builder.Write("領域班排名百分比");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域年排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYR" + "»");
                    mName1 = domainName + "_領域年排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR" + "»");
                    mName1 = domainName + "_領域年排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYP" + "»");
                    mName1 = domainName + "_領域班排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCR" + "»");
                    mName1 = domainName + "_領域班排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR" + "»");
                    mName1 = domainName + "_領域班排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCP" + "»");
                    builder.EndRow();
                }

                builder.EndTable(); // 領域成績排名1

                builder.Writeln();
                builder.Writeln();

                // 領域成績排名2
                builder.Write("領域成績排名");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("領域類別1排名名次");
                builder.InsertCell();
                builder.Write("領域類別1排名PR值");
                builder.InsertCell();
                builder.Write("領域類別1排名百分比");
                builder.InsertCell();
                builder.Write("領域類別2排名名次");
                builder.InsertCell();
                builder.Write("領域類別2排名PR值");
                builder.InsertCell();
                builder.Write("領域類別2排名百分比");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域類別1排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TR" + "»");
                    mName1 = domainName + "_領域類別1排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR" + "»");
                    mName1 = domainName + "_領域類別1排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TP" + "»");
                    mName1 = domainName + "_領域類別2排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TR" + "»");
                    mName1 = domainName + "_領域類別2排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR" + "»");
                    mName1 = domainName + "_領域類別2排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TP" + "»");
                    builder.EndRow();

                }
                builder.EndTable(); // 領域成績排名2

                builder.Writeln();
                builder.Writeln();

                // 領域成績五標(年排名及班排名)
                builder.Write("領域成績五標、標準差(年排名及班排名)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("領域年排名母體頂標");
                builder.InsertCell();
                builder.Write("領域年排名母體前標");
                builder.InsertCell();
                builder.Write("領域年排名母體平均");
                builder.InsertCell();
                builder.Write("領域年排名母體後標");
                builder.InsertCell();
                builder.Write("領域年排名母體底標");
                builder.InsertCell();
                builder.Write("領域年排名母體人數");
                builder.InsertCell();
                builder.Write("領域年排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域年排名母體新前標");
                builder.InsertCell();
                builder.Write("領域年排名母體新均標");
                builder.InsertCell();
                builder.Write("領域年排名母體新後標");
                builder.InsertCell();
                builder.Write("領域年排名母體新底標");
                builder.InsertCell();
                builder.Write("領域年排名母體標準差");

                builder.InsertCell();
                builder.Write("領域班排名母體頂標");
                builder.InsertCell();
                builder.Write("領域班排名母體前標");
                builder.InsertCell();
                builder.Write("領域班排名母體平均");
                builder.InsertCell();
                builder.Write("領域班排名母體後標");
                builder.InsertCell();
                builder.Write("領域班排名母體底標");
                builder.InsertCell();
                builder.Write("領域班排名母體人數");
                builder.InsertCell();
                builder.Write("領域班排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域班排名母體新前標");
                builder.InsertCell();
                builder.Write("領域班排名母體新均標");
                builder.InsertCell();
                builder.Write("領域班排名母體新後標");
                builder.InsertCell();
                builder.Write("領域班排名母體新底標");
                builder.InsertCell();
                builder.Write("領域班排名母體標準差");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域年排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY25T" + "»");
                    mName1 = domainName + "_領域年排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY50T" + "»");
                    mName1 = domainName + "_領域年排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYA" + "»");
                    mName1 = domainName + "_領域年排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY50B" + "»");
                    mName1 = domainName + "_領域年排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY50B" + "»");
                    mName1 = domainName + "_領域年排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY25B" + "»");

                    mName1 = domainName + "_領域年排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR88" + "»");
                    mName1 = domainName + "_領域年排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR75" + "»");
                    mName1 = domainName + "_領域年排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR50" + "»");
                    mName1 = domainName + "_領域年排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR25" + "»");
                    mName1 = domainName + "_領域年排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR12" + "»");
                    mName1 = domainName + "_領域年排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYSTD" + "»");

                    mName1 = domainName + "_領域班排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC25T" + "»");
                    mName1 = domainName + "_領域班排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC50T" + "»");
                    mName1 = domainName + "_領域班排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCA" + "»");
                    mName1 = domainName + "_領域班排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC50B" + "»");
                    mName1 = domainName + "_領域班排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC25B" + "»");
                    mName1 = domainName + "_領域班排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC2MC" + "»");

                    mName1 = domainName + "_領域班排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR88" + "»");
                    mName1 = domainName + "_領域班排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR75" + "»");
                    mName1 = domainName + "_領域班排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR50" + "»");
                    mName1 = domainName + "_領域班排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR25" + "»");
                    mName1 = domainName + "_領域班排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR12" + "»");
                    mName1 = domainName + "_領域班排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCSTD" + "»");

                    builder.EndRow();
                }

                builder.EndTable(); // 領域成績五標(年排名及班排名)

                builder.Writeln();
                builder.Writeln();

                // 領域成績五標(類別1排名及類別2排名)
                builder.Write("領域成績五標、標準差(類別1排名及類別2排名)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("領域類別1排名母體頂標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體前標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體平均");
                builder.InsertCell();
                builder.Write("領域類別1排名母體後標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體底標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體人數");
                builder.InsertCell();
                builder.Write("領域類別1排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體新前標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體新均標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體新後標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體新底標");
                builder.InsertCell();
                builder.Write("領域類別1排名母體標準差");

                builder.InsertCell();
                builder.Write("領域類別2排名母體頂標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體前標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體平均");
                builder.InsertCell();
                builder.Write("領域類別2排名母體後標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體底標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體人數");
                builder.InsertCell();
                builder.Write("領域類別2排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體新前標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體新均標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體新後標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體新底標");
                builder.InsertCell();
                builder.Write("領域類別2排名母體標準差");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域類別1排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T25T" + "»");
                    mName1 = domainName + "_領域類別1排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T50T" + "»");
                    mName1 = domainName + "_領域類別1排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TA" + "»");
                    mName1 = domainName + "_領域類別1排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T50B" + "»");
                    mName1 = domainName + "_領域類別1排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T25B" + "»");
                    mName1 = domainName + "_領域類別1排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TMC" + "»");

                    mName1 = domainName + "_領域類別1排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR88" + "»");
                    mName1 = domainName + "_領域類別1排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR75" + "»");
                    mName1 = domainName + "_領域類別1排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR50" + "»");
                    mName1 = domainName + "_領域類別1排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR25" + "»");
                    mName1 = domainName + "_領域類別1排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR12" + "»");
                    mName1 = domainName + "_領域類別1排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TSTD" + "»");

                    mName1 = domainName + "_領域類別2排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T25T" + "»");
                    mName1 = domainName + "_領域類別2排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T50T" + "»");
                    mName1 = domainName + "_領域類別2排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TA" + "»");
                    mName1 = domainName + "_領域類別2排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T50B" + "»");
                    mName1 = domainName + "_領域類別2排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T25B" + "»");
                    mName1 = domainName + "_領域類別2排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TMC" + "»");

                    mName1 = domainName + "_領域類別2排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR88" + "»");
                    mName1 = domainName + "_領域類別2排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR75" + "»");
                    mName1 = domainName + "_領域類別2排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR50" + "»");
                    mName1 = domainName + "_領域類別2排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR25" + "»");
                    mName1 = domainName + "_領域類別2排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR12" + "»");
                    mName1 = domainName + "_領域類別2排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TSTD" + "»");
                    builder.EndRow();
                }

                builder.EndTable(); // 領域成績五標(類別1排名及類別2排名)

                builder.Writeln();
                builder.Writeln();


                // 領域成績定期
                // 處理領域定期成績
                builder.Write("領域定期成績");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("領域定期班排名名次");
                builder.InsertCell();
                builder.Write("領域定期班排名PR值");
                builder.InsertCell();
                builder.Write("領域定期班排名百分比");
                builder.InsertCell();
                builder.Write("領域定期班排名母體平均");
                builder.InsertCell();
                builder.Write("領域定期年排名名次");
                builder.InsertCell();
                builder.Write("領域定期年排名PR值");
                builder.InsertCell();
                builder.Write("領域定期年排名百分比");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";

                    mName1 = domainName + "_領域定期班排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCR" + "»");
                    mName1 = domainName + "_領域定期班排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR" + "»");
                    mName1 = domainName + "_領域定期班排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCP" + "»");
                    mName1 = domainName + "_領域定期班排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCA" + "»");
                    mName1 = domainName + "_領域定期年排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYR" + "»");
                    mName1 = domainName + "_領域定期年排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR" + "»");
                    mName1 = domainName + "_領域定期年排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYP" + "»");

                    builder.EndRow();
                }
                builder.EndTable(); // 領域定期成績

                builder.Writeln();
                builder.Writeln();

                // 領域定期成績排名1
                builder.Write("領域定期成績排名");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域定期名稱");
                builder.InsertCell();
                builder.Write("領域定期年排名名次");
                builder.InsertCell();
                builder.Write("領域定期年排名PR值");
                builder.InsertCell();
                builder.Write("領域定期年排名百分比");
                builder.InsertCell();
                builder.Write("領域定期班排名名次");
                builder.InsertCell();
                builder.Write("領域定期班排名PR值");
                builder.InsertCell();
                builder.Write("領域定期班排名百分比");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域定期年排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYR" + "»");
                    mName1 = domainName + "_領域定期年排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR" + "»");
                    mName1 = domainName + "_領域定期年排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYP" + "»");
                    mName1 = domainName + "_領域定期班排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCR" + "»");
                    mName1 = domainName + "_領域定期班排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR" + "»");
                    mName1 = domainName + "_領域定期班排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCP" + "»");
                    builder.EndRow();
                }

                builder.EndTable(); // 領域定期成績排名1

                builder.Writeln();
                builder.Writeln();

                // 領域定期成績排名2
                builder.Write("領域定期成績排名");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域定期名稱");
                builder.InsertCell();
                builder.Write("領域定期類別1排名名次");
                builder.InsertCell();
                builder.Write("領域定期類別1排名PR值");
                builder.InsertCell();
                builder.Write("領域定期類別1排名百分比");
                builder.InsertCell();
                builder.Write("領域定期類別2排名名次");
                builder.InsertCell();
                builder.Write("領域定期類別2排名PR值");
                builder.InsertCell();
                builder.Write("領域定期類別2排名百分比");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域定期類別1排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TR" + "»");
                    mName1 = domainName + "_領域定期類別1排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR" + "»");
                    mName1 = domainName + "_領域定期類別1排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TP" + "»");
                    mName1 = domainName + "_領域定期類別2排名名次";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TR" + "»");
                    mName1 = domainName + "_領域定期類別2排名PR值";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR" + "»");
                    mName1 = domainName + "_領域定期類別2排名百分比";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TP" + "»");
                    builder.EndRow();

                }
                builder.EndTable(); // 領域定期成績排名2

                builder.Writeln();
                builder.Writeln();

                // 領域定期成績五標(年排名及班排名)
                builder.Write("領域定期成績五標、標準差(年排名及班排名)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域定期名稱");
                builder.InsertCell();
                builder.Write("領域定期年排名母體頂標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體前標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體平均");
                builder.InsertCell();
                builder.Write("領域定期年排名母體後標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體底標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體人數");
                builder.InsertCell();
                builder.Write("領域定期年排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體新前標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體新均標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體新後標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體新底標");
                builder.InsertCell();
                builder.Write("領域定期年排名母體標準差");

                builder.InsertCell();
                builder.Write("領域定期班排名母體頂標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體前標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體平均");
                builder.InsertCell();
                builder.Write("領域定期班排名母體後標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體底標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體人數");
                builder.InsertCell();
                builder.Write("領域定期班排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體新前標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體新均標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體新後標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體新底標");
                builder.InsertCell();
                builder.Write("領域定期班排名母體標準差");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域定期年排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY25T" + "»");
                    mName1 = domainName + "_領域定期年排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY50T" + "»");
                    mName1 = domainName + "_領域定期年排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYA" + "»");
                    mName1 = domainName + "_領域定期年排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY50B" + "»");
                    mName1 = domainName + "_領域定期年排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY25B" + "»");
                    mName1 = domainName + "_領域定期年排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DY25B" + "»");
                    mName1 = domainName + "_領域定期年排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR88" + "»");
                    mName1 = domainName + "_領域定期年排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR75" + "»");
                    mName1 = domainName + "_領域定期年排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR50" + "»");
                    mName1 = domainName + "_領域定期年排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR25" + "»");
                    mName1 = domainName + "_領域定期年排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYPR12" + "»");
                    mName1 = domainName + "_領域定期年排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DYSTD" + "»");

                    mName1 = domainName + "_領域定期班排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC25T" + "»");
                    mName1 = domainName + "_領域定期班排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC50T" + "»");
                    mName1 = domainName + "_領域定期班排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCA" + "»");
                    mName1 = domainName + "_領域定期班排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC50B" + "»");
                    mName1 = domainName + "_領域定期班排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC25B" + "»");
                    mName1 = domainName + "_領域定期班排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DC2MC" + "»");

                    mName1 = domainName + "_領域定期班排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR88" + "»");
                    mName1 = domainName + "_領域定期班排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR75" + "»");
                    mName1 = domainName + "_領域定期班排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR50" + "»");
                    mName1 = domainName + "_領域定期班排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR25" + "»");
                    mName1 = domainName + "_領域定期班排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCPR12" + "»");
                    mName1 = domainName + "_領域定期班排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DCSTD" + "»");
                    builder.EndRow();
                }

                builder.EndTable(); // 領域定期成績五標(年排名及班排名)

                builder.Writeln();
                builder.Writeln();

                // 領域定期成績五標(類別1排名及類別2排名)
                builder.Write("領域定期成績五標、標準差(類別1排名及類別2排名)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域定期名稱");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體頂標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體前標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體平均");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體後標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體底標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體人數");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體新前標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體新均標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體新後標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體新底標");
                builder.InsertCell();
                builder.Write("領域定期類別1排名母體標準差");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體頂標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體前標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體平均");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體後標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體底標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體人數");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體新頂標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體新前標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體新均標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體新後標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體新底標");
                builder.InsertCell();
                builder.Write("領域定期類別2排名母體標準差");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    string mName1 = "";
                    mName1 = domainName + "_領域定期類別1排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T25T" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T50T" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TA" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T50B" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1T25B" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TMC" + "»");

                    mName1 = domainName + "_領域定期類別1排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR88" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR75" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR50" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR25" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TPR12" + "»");
                    mName1 = domainName + "_領域定期類別1排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D1TSTD" + "»");

                    mName1 = domainName + "_領域定期類別2排名母體頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T25T" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T50T" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體平均";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TA" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T50B" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2T25B" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體人數";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TMC" + "»");

                    mName1 = domainName + "_領域定期類別2排名母體新頂標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR88" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體新前標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR75" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體新均標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR50" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體新後標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR25" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體新底標";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TPR12" + "»");
                    mName1 = domainName + "_領域定期類別2排名母體標準差";
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«D2TSTD" + "»");

                    builder.EndRow();
                }

                builder.EndTable(); // 領域定期成績五標(類別1排名及類別2排名)

                builder.Writeln();
                builder.Writeln();


                builder.Writeln("序列化科目資料(年排名及班排名)");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目班排名名次");
                    builder.InsertCell();
                    builder.Write("科目班排名PR值");
                    builder.InsertCell();
                    builder.Write("科目班排名百分比");
                    builder.InsertCell();
                    builder.Write("科目年排名名次");
                    builder.InsertCell();
                    builder.Write("科目年排名PR值");
                    builder.InsertCell();
                    builder.Write("科目年排名百分比");
                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目班排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCR" + "»");
                        mName1 = domainName + "_科目班排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR" + "»");
                        mName1 = domainName + "_科目班排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCP" + "»");
                        mName1 = domainName + "_科目年排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYR" + "»");
                        mName1 = domainName + "_科目年排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR" + "»");
                        mName1 = domainName + "_科目年排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYP" + "»");
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("序列化科目資料(類別1排名及類別2排名)");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目類別1排名名次");
                    builder.InsertCell();
                    builder.Write("科目類別1排名PR值");
                    builder.InsertCell();
                    builder.Write("科目類別1排名百分比");
                    builder.InsertCell();
                    builder.Write("科目類別2排名名次");
                    builder.InsertCell();
                    builder.Write("科目類別2排名PR值");
                    builder.InsertCell();
                    builder.Write("科目類別2排名百分比");
                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目類別1排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TR" + "»");
                        mName1 = domainName + "_科目類別1排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR" + "»");
                        mName1 = domainName + "_科目類別1排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TP" + "»");
                        mName1 = domainName + "_科目類別2排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TR" + "»");
                        mName1 = domainName + "_科目類別2排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR" + "»");
                        mName1 = domainName + "_科目類別2排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TP" + "»");
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("序列化科目資料五標、標準差(年排名及班排名)");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目年排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目年排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體人數");
                    builder.InsertCell();
                    builder.Write("科目年排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目年排名母體標準差");

                    builder.InsertCell();
                    builder.Write("科目班排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目班排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體人數");
                    builder.InsertCell();
                    builder.Write("科目班排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目班排名母體標準差");

                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目年排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY25T" + "»");
                        mName1 = domainName + "_科目年排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY50T" + "»");
                        mName1 = domainName + "_科目年排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYA" + "»");
                        mName1 = domainName + "_科目年排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY50B" + "»");
                        mName1 = domainName + "_科目年排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY25B" + "»");
                        mName1 = domainName + "_科目年排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYMC" + "»");

                        mName1 = domainName + "_科目年排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR88" + "»");
                        mName1 = domainName + "_科目年排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR75" + "»");
                        mName1 = domainName + "_科目年排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR50" + "»");
                        mName1 = domainName + "_科目年排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR25" + "»");
                        mName1 = domainName + "_科目年排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR12" + "»");
                        mName1 = domainName + "_科目年排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYMSTD" + "»");


                        mName1 = domainName + "_科目班排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC25T" + "»");
                        mName1 = domainName + "_科目班排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC50T" + "»");
                        mName1 = domainName + "_科目班排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCA" + "»");
                        mName1 = domainName + "_科目班排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC50B" + "»");
                        mName1 = domainName + "_科目班排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC25B" + "»");
                        mName1 = domainName + "_科目班排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCMC" + "»");

                        mName1 = domainName + "_科目班排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR88" + "»");
                        mName1 = domainName + "_科目班排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR75" + "»");
                        mName1 = domainName + "_科目班排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR50" + "»");
                        mName1 = domainName + "_科目班排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR25" + "»");
                        mName1 = domainName + "_科目班排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR12" + "»");
                        mName1 = domainName + "_科目班排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCMSTD" + "»");
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("序列化科目資料五標、標準差(類別1排名及類別2排名)");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體母數");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目類別1排名母體標準差");

                    builder.InsertCell();
                    builder.Write("科目類別2排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體人數");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目類別2排名母體標準差");
                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目類別1排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T25T" + "»");
                        mName1 = domainName + "_科目類別1排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T50T" + "»");
                        mName1 = domainName + "_科目類別1排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TA" + "»");
                        mName1 = domainName + "_科目類別1排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T50B" + "»");
                        mName1 = domainName + "_科目類別1排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T25B" + "»");
                        mName1 = domainName + "_科目類別1排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T25B" + "»");

                        mName1 = domainName + "_科目類別1排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR88" + "»");
                        mName1 = domainName + "_科目類別1排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR75" + "»");
                        mName1 = domainName + "_科目類別1排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR50" + "»");
                        mName1 = domainName + "_科目類別1排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR25" + "»");
                        mName1 = domainName + "_科目類別1排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR12" + "»");
                        mName1 = domainName + "_科目類別1排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TSTD" + "»");

                        mName1 = domainName + "_科目類別2排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T25T" + "»");
                        mName1 = domainName + "_科目類別2排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T50T" + "»");
                        mName1 = domainName + "_科目類別2排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TA" + "»");
                        mName1 = domainName + "_科目類別2排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T50B" + "»");
                        mName1 = domainName + "_科目類別2排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T25B" + "»");

                        mName1 = domainName + "_科目類別2排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TMC" + "»");

                        mName1 = domainName + "_科目類別2排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR88" + "»");
                        mName1 = domainName + "_科目類別2排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR75" + "»");
                        mName1 = domainName + "_科目類別2排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR50" + "»");
                        mName1 = domainName + "_科目類別2排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR25" + "»");
                        mName1 = domainName + "_科目類別2排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR12" + "»");
                        mName1 = domainName + "_科目類別2排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TSTD" + "»");

                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                // 科目定期
                builder.Writeln("序列化科目定期評量資料(年排名及班排名)");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目定期班排名名次");
                    builder.InsertCell();
                    builder.Write("科目定期班排名PR值");
                    builder.InsertCell();
                    builder.Write("科目定期班排名百分比");
                    builder.InsertCell();
                    builder.Write("科目定期年排名名次");
                    builder.InsertCell();
                    builder.Write("科目定期年排名PR值");
                    builder.InsertCell();
                    builder.Write("科目定期年排名百分比");
                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目定期班排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCR" + "»");
                        mName1 = domainName + "_科目定期班排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR" + "»");
                        mName1 = domainName + "_科目定期班排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCP" + "»");
                        mName1 = domainName + "_科目定期年排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYR" + "»");
                        mName1 = domainName + "_科目定期年排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR" + "»");
                        mName1 = domainName + "_科目定期年排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYP" + "»");
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("序列化科目定期評量資料(類別1排名及類別2排名)");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名名次");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名PR值");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名百分比");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名名次");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名PR值");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名百分比");
                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目定期類別1排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TR" + "»");
                        mName1 = domainName + "_科目定期類別1排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR" + "»");
                        mName1 = domainName + "_科目定期類別1排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TP" + "»");
                        mName1 = domainName + "_科目定期類別2排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TR" + "»");
                        mName1 = domainName + "_科目定期類別2排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR" + "»");
                        mName1 = domainName + "_科目定期類別2排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TP" + "»");
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }


                builder.Writeln("序列化科目定期評量資料五標、標準差(年排名)");
                builder.Writeln("※Ａ++~B適用於「定期評量排名擴充功能」模組");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體人數");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目定期年排名母體標準差");
                    builder.InsertCell();
                    builder.Write("科目定期年排名A++");
                    builder.InsertCell();
                    builder.Write("科目定期年排名A+");
                    builder.InsertCell();
                    builder.Write("科目定期年排名A");
                    builder.InsertCell();
                    builder.Write("科目定期年排名B++");
                    builder.InsertCell();
                    builder.Write("科目定期年排名B+");
                    builder.InsertCell();
                    builder.Write("科目定期年排名B");


                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目定期年排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY25T" + "»");
                        mName1 = domainName + "_科目定期年排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY50T" + "»");
                        mName1 = domainName + "_科目定期年排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYA" + "»");
                        mName1 = domainName + "_科目定期年排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY50B" + "»");
                        mName1 = domainName + "_科目定期年排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SY25B" + "»");
                        mName1 = domainName + "_科目定期年排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYMC" + "»");
                        mName1 = domainName + "_科目定期年排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR88" + "»");
                        mName1 = domainName + "_科目定期年排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR75" + "»");
                        mName1 = domainName + "_科目定期年排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR50" + "»");
                        mName1 = domainName + "_科目定期年排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR25" + "»");
                        mName1 = domainName + "_科目定期年排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR12" + "»");
                        mName1 = domainName + "_科目定期年排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYMSTD" + "»");

                        mName1 = domainName + "_科目定期年排名A++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYA++" + "»");
                        mName1 = domainName + "_科目定期年排名A+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYA+" + "»");
                        mName1 = domainName + "_科目定期年排名A" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYA" + "»");
                        mName1 = domainName + "_科目定期年排名B++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYB++" + "»");
                        mName1 = domainName + "_科目定期年排名B+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYB+" + "»");
                        mName1 = domainName + "_科目定期年排名B" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYB" + "»");

                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("序列化科目定期評量資料五標、標準差(班排名)");
                builder.Writeln("※Ａ++~B適用於「定期評量排名擴充功能」模組");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");

                    builder.InsertCell();
                    builder.Write("科目定期班排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體人數");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目定期班排名母體標準差");

                    builder.InsertCell();
                    builder.Write("科目定期班排名A++");
                    builder.InsertCell();
                    builder.Write("科目定期班排名A+");
                    builder.InsertCell();
                    builder.Write("科目定期班排名A");
                    builder.InsertCell();
                    builder.Write("科目定期班排名B++");
                    builder.InsertCell();
                    builder.Write("科目定期班排名B+");
                    builder.InsertCell();
                    builder.Write("科目定期班排名B");

                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");

                        mName1 = domainName + "_科目定期班排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC25T" + "»");
                        mName1 = domainName + "_科目定期班排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC50T" + "»");
                        mName1 = domainName + "_科目定期班排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCA" + "»");
                        mName1 = domainName + "_科目定期班排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC50B" + "»");
                        mName1 = domainName + "_科目定期班排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC25B" + "»");
                        mName1 = domainName + "_科目定期班排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCMC" + "»");

                        mName1 = domainName + "_科目定期班排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR88" + "»");
                        mName1 = domainName + "_科目定期班排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR75" + "»");
                        mName1 = domainName + "_科目定期班排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR50" + "»");
                        mName1 = domainName + "_科目定期班排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR25" + "»");
                        mName1 = domainName + "_科目定期班排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR12" + "»");
                        mName1 = domainName + "_科目定期班排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCMSTD" + "»");

                        mName1 = domainName + "_科目定期班排名A++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCA++" + "»");
                        mName1 = domainName + "_科目定期班排名A+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCA+" + "»");
                        mName1 = domainName + "_科目定期班排名A" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCA" + "»");
                        mName1 = domainName + "_科目定期班排名B++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCB++" + "»");
                        mName1 = domainName + "_科目定期班排名B+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCB+" + "»");
                        mName1 = domainName + "_科目定期班排名B" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCB" + "»");

                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("序列化科目定期評量資料五標、標準差(類別1排名)");
                builder.Writeln("※Ａ++~B適用於「定期評量排名擴充功能」模組");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體人數");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名母體標準差");

                    builder.InsertCell();
                    builder.Write("科目定期類別1排名A++");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名A+");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名A");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名B++");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名B+");
                    builder.InsertCell();
                    builder.Write("科目定期類別1排名B");
                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T25T" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T50T" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TA" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T50B" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1T25B" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TMC" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR88" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR75" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR50" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR25" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TPR12" + "»");
                        mName1 = domainName + "_科目定期類別1排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S1TMSTD" + "»");

                        mName1 = domainName + "_科目定期類別1排名A++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST1A++" + "»");
                        mName1 = domainName + "_科目定期類別1排名A+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST1A+" + "»");
                        mName1 = domainName + "_科目定期類別1排名A" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST1A" + "»");
                        mName1 = domainName + "_科目定期類別1排名B++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST1B++" + "»");
                        mName1 = domainName + "_科目定期類別1排名B+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST1B+" + "»");
                        mName1 = domainName + "_科目定期類別1排名B" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST1B" + "»");
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("序列化科目定期評量資料五標、標準差(類別2排名)");
                builder.Writeln("※Ａ++~B適用於「定期評量排名擴充功能」模組");
                foreach (string domainName in DomainNameList)
                {
                    builder.Write(domainName + "領域");
                    string mName1 = "";

                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體頂標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體前標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體後標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體底標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體人數");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體新頂標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體新前標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體新均標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體新後標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體新底標");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名母體標準差");

                    builder.InsertCell();
                    builder.Write("科目定期類別2排名A++");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名A+");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名A");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名B++");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名B+");
                    builder.InsertCell();
                    builder.Write("科目定期類別2排名B");
                    builder.EndRow();

                    for (int sj = 1; sj <= 12; sj++)
                    {
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");

                        mName1 = domainName + "_科目定期類別2排名母體頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T25T" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T50T" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TA" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T50B" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2T25B" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體人數" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TMC" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體新頂標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR88" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體新前標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR75" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體新均標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR50" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體新後標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR25" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體新底標" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TPR12" + "»");
                        mName1 = domainName + "_科目定期類別2排名母體標準差" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«S2TMSTD" + "»");

                        mName1 = domainName + "_科目定期類別2排名A++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST2A++" + "»");
                        mName1 = domainName + "_科目定期類別2排名A+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST2A+" + "»");
                        mName1 = domainName + "_科目定期類別2排名A" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST2A" + "»");
                        mName1 = domainName + "_科目定期類別2排名B++" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST2B++" + "»");
                        mName1 = domainName + "_科目定期類別2排名B+" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST2B+" + "»");
                        mName1 = domainName + "_科目定期類別2排名B" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST2B" + "»");

                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }



                List<string> tmpRNameList = new List<string>();
                tmpRNameList.Add("_R0_9");
                tmpRNameList.Add("_R10_19");
                tmpRNameList.Add("_R20_29");
                tmpRNameList.Add("_R30_39");
                tmpRNameList.Add("_R40_49");
                tmpRNameList.Add("_R50_59");
                tmpRNameList.Add("_R60_69");
                tmpRNameList.Add("_R70_79");
                tmpRNameList.Add("_R80_89");
                tmpRNameList.Add("_R90_99");
                tmpRNameList.Add("_R100_u");

                builder.Writeln("領域成績組距(總成績)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("班級_" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "班級_" + domainName + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("年級" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "年級_" + domainName + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("類別1_" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "類別1_" + domainName + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("類別2_" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "類別2_" + domainName + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }


                builder.EndTable();
                builder.Writeln();
                builder.Writeln();


                builder.Writeln("領域成績組距(定期)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("班級_" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "班級_" + domainName + "F" + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("年級" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "年級_" + domainName + "F" + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("類別1_" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "類別1_" + domainName + itName + "F";
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }

                foreach (string domainName in DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write("類別2_" + domainName);

                    string mName1 = "";

                    foreach (string itName in tmpRNameList)
                    {
                        mName1 = "類別2_" + domainName + itName + "F";
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«DR" + "»");
                    }
                    builder.EndRow();
                }

                builder.EndTable();
                builder.Writeln();
                builder.Writeln();



                foreach (string domainName in DomainNameList)
                {

                    builder.Writeln(domainName + "領域");
                    builder.Writeln("※「科目定期評量自訂等第」適用於「定期評量排名擴充功能」模組");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目權數");
                    builder.InsertCell();
                    builder.Write("科目權數含括號");
                    builder.InsertCell();
                    builder.Write("科目定期評量");
                    builder.InsertCell();
                    builder.Write("科目平時評量");
                    builder.InsertCell();
                    builder.Write("科目總成績");
                    builder.InsertCell();
                    builder.Write("科目定期評量等第");
                    builder.InsertCell();
                    builder.Write("科目平時評量等第");
                    builder.InsertCell();
                    builder.Write("科目總成績等第");
                    builder.InsertCell();
                    builder.Write("科目定期評量自訂等第(年排名)");
                    builder.InsertCell();
                    builder.Write("科目定期評量自訂等第(班排名)");
                    builder.InsertCell();
                    builder.Write("科目定期評量自訂等第(類別1排名)");
                    builder.InsertCell();
                    builder.Write("科目定期評量自訂等第(類別2排名)");
                    builder.InsertCell();
                    builder.Write("科目文字評量");
                    builder.InsertCell();
                    builder.Write("科目班排名名次");
                    builder.InsertCell();
                    builder.Write("科目班排名PR值");
                    builder.InsertCell();
                    builder.Write("科目班排名百分比");
                    builder.InsertCell();
                    builder.Write("科目班排名母體平均");
                    builder.InsertCell();
                    builder.Write("科目年排名名次");
                    builder.InsertCell();
                    builder.Write("科目年排名PR值");
                    builder.InsertCell();
                    builder.Write("科目年排名百分比");
                    builder.InsertCell();
                    //builder.Write("科目年排名百分比");
                    //builder.InsertCell();
                    builder.Write("科目年排母體平均");
                    builder.InsertCell();
                    builder.Write("參考試別科目定期評量");
                    builder.InsertCell();
                    builder.Write("參考試別科目平時評量");
                    builder.InsertCell();
                    builder.Write("參考試別科目總成績");
                    builder.EndRow();
                    // todo 新增參考識別功能變數
                    for (int sj = 1; sj <= 12; sj++)
                    {
                        string mName1 = "";
                        mName1 = domainName + "_科目名稱" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SN" + "»");
                        mName1 = domainName + "_科目權數" + sj;

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC" + "»");

                        mName1 = domainName + "_科目權數含括號" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC" + "»");

                        mName1 = domainName + "_科目定期評量" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SF" + "»");
                        mName1 = domainName + "_科目平時評量" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SA" + "»");
                        mName1 = domainName + "_科目總成績" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SC" + "»");
                        mName1 = domainName + "_科目定期評量等第" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SF" + "»");
                        mName1 = domainName + "_科目平時評量等第" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SA" + "»");
                        mName1 = domainName + "_科目總成績等第" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SST" + "»");

                        mName1 = domainName + "_科目定期評量自訂等第(年排名)" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SFY" + "»");
                        mName1 = domainName + "_科目定期評量自訂等第(班排名)" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SFC" + "»");
                        mName1 = domainName + "_科目定期評量自訂等第(類別1排名)" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SFG" + "»");
                        mName1 = domainName + "_科目定期評量自訂等第(類別2排名)" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SFP" + "»");

                        mName1 = domainName + "_科目文字評量" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«ST" + "»");
                        mName1 = domainName + "_科目班排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCR" + "»");
                        mName1 = domainName + "_科目班排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCPR" + "»");
                        mName1 = domainName + "_科目班排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCP" + "»");
                        mName1 = domainName + "_科目班排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SCA" + "»");
                        mName1 = domainName + "_科目年排名名次" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYR" + "»");
                        mName1 = domainName + "_科目年排名PR值" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYPR" + "»");
                        mName1 = domainName + "_科目年排名百分比" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYP" + "»");
                        mName1 = domainName + "_科目年排名母體平均" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«SYA" + "»");
                        // Jean 新增參考識別
                        mName1 = domainName + "_參考試別科目定期評量" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«RSF" + "»");
                        mName1 = domainName + "_參考試別科目平時評量" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«RSA" + "»");
                        mName1 = domainName + "_參考試別科目總成績" + sj;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + mName1 + " \\* MERGEFORMAT ", "«TSC" + "»");
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();
                    builder.Writeln();
                }

                builder.Writeln("科目班級組距(總成績)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "s班級_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "s班級_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();

                builder.Writeln("科目班級組距(定期)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "sf班級_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "sf班級_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();


                builder.Writeln("科目年級組距(總成績)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "s年級_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "s年級_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();

                builder.Writeln("科目年級組距(定期)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "sf年級_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "sf年級_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();

                builder.Writeln("科目類別1組距(總成績)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "s類別1_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "s類別1_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();

                builder.Writeln("科目類別1組距(定期)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "sf類別1_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "sf類別1_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();

                builder.Writeln("科目類別2組距(總成績)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "s類別2_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "s類別2_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();

                builder.Writeln("科目類別2組距(定期)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("領域名稱");
                builder.InsertCell();
                builder.Write("0-9");
                builder.InsertCell();
                builder.Write("10-19");
                builder.InsertCell();
                builder.Write("20-29");
                builder.InsertCell();
                builder.Write("30-39");
                builder.InsertCell();
                builder.Write("40-49");
                builder.InsertCell();
                builder.Write("50-59");
                builder.InsertCell();
                builder.Write("60-69");
                builder.InsertCell();
                builder.Write("70-79");
                builder.InsertCell();
                builder.Write("80-89");
                builder.InsertCell();
                builder.Write("90-99");
                builder.InsertCell();
                builder.Write("100以上");
                builder.EndRow();


                for (int ss = 1; ss <= 30; ss++)
                {
                    string sName1 = "";
                    sName1 = "sf類別2_科目名稱" + ss;
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«N" + "»");

                    foreach (string itName in tmpRNameList)
                    {

                        sName1 = "sf類別2_科目" + ss + itName;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + sName1 + " \\* MERGEFORMAT ", "«SR" + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
                builder.Writeln();


                // 領域-科目1,2,3 組距
                List<string> rrKNameList = new List<string>();
                rrKNameList.Add("班級_");
                rrKNameList.Add("年級_");
                rrKNameList.Add("類別1_");
                rrKNameList.Add("類別2_");

                // 總成績
                foreach (string rk in rrKNameList)
                {
                    foreach (string dname in Global.DomainNameList)
                    {
                        builder.Writeln(dname + "領域-科目 " + rk + "組距");
                        builder.StartTable();
                        builder.InsertCell();
                        builder.Write("領域科目名稱");
                        builder.InsertCell();
                        builder.Write("0-9");
                        builder.InsertCell();
                        builder.Write("10-19");
                        builder.InsertCell();
                        builder.Write("20-29");
                        builder.InsertCell();
                        builder.Write("30-39");
                        builder.InsertCell();
                        builder.Write("40-49");
                        builder.InsertCell();
                        builder.Write("50-59");
                        builder.InsertCell();
                        builder.Write("60-69");
                        builder.InsertCell();
                        builder.Write("70-79");
                        builder.InsertCell();
                        builder.Write("80-89");
                        builder.InsertCell();
                        builder.Write("90-99");
                        builder.InsertCell();
                        builder.Write("100以上");
                        builder.EndRow();

                        for (int i = 1; i <= 5; i++)
                        {
                            //                    班級_語文領域_科目名稱7
                            //班級_語文領域_科目7_R100_u

                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + rk + dname + "領域_科目名稱" + i + " \\* MERGEFORMAT ", "«N" + i + "»");

                            foreach (string r in tmpRNameList)
                            {
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + rk + dname + "領域_科目" + i + r + " \\* MERGEFORMAT ", "«R" + i + "»");
                            }

                            builder.EndRow();
                        }

                        builder.EndTable();
                        builder.Writeln();
                    }
                }

                // 定期成績
                foreach (string rk in rrKNameList)
                {
                    foreach (string dname in Global.DomainNameList)
                    {
                        builder.Writeln(dname + "領域-科目定期 " + rk + "組距");
                        builder.StartTable();
                        builder.InsertCell();
                        builder.Write("領域科目名稱");
                        builder.InsertCell();
                        builder.Write("0-9");
                        builder.InsertCell();
                        builder.Write("10-19");
                        builder.InsertCell();
                        builder.Write("20-29");
                        builder.InsertCell();
                        builder.Write("30-39");
                        builder.InsertCell();
                        builder.Write("40-49");
                        builder.InsertCell();
                        builder.Write("50-59");
                        builder.InsertCell();
                        builder.Write("60-69");
                        builder.InsertCell();
                        builder.Write("70-79");
                        builder.InsertCell();
                        builder.Write("80-89");
                        builder.InsertCell();
                        builder.Write("90-99");
                        builder.InsertCell();
                        builder.Write("100以上");
                        builder.EndRow();

                        for (int i = 1; i <= 5; i++)
                        {
                            //                    班級_語文領域_科目名稱7
                            //班級_語文領域_科目7F_R100_u

                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + rk + dname + "領域_科目名稱" + i + " \\* MERGEFORMAT ", "«N" + i + "»");

                            foreach (string r in tmpRNameList)
                            {
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + rk + dname + "領域_科目" + i + "F" + r + " \\* MERGEFORMAT ", "«R" + i + "»");
                            }

                            builder.EndRow();
                        }

                        builder.EndTable();
                        builder.Writeln();
                    }
                }

                //組距 (程式用詞)^^^(顯示用詞)
                string[] intervals = new string[] {  "R0_9^0-9",
                                                    "R10_19^10-19",
                                                    "R20_29^20-29",
                                                    "R30_39^30-39",
                                                    "R40_49^40-49",
                                                    "R50_59^50-59",
                                                    "R60_69^60-69",
                                                    "R70_79^70-79",
                                                    "R80_89^80-89",
                                                    "R90_99^90-99",
                                                     "R100_u^100以上"};


                // 總計成績 平均/總分 
                foreach (string rankType in _rankTypes) //"班排名", "年排名", "類別1排名", "類別2排名" 
                {
                    builder.Writeln("總計成績-科目定期 " + rankType + " 組距");
                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("");
                    builder.InsertCell();
                    builder.Write("0-9");
                    builder.InsertCell();
                    builder.Write("10-19");
                    builder.InsertCell();
                    builder.Write("20-29");
                    builder.InsertCell();
                    builder.Write("30-39");
                    builder.InsertCell();
                    builder.Write("40-49");
                    builder.InsertCell();
                    builder.Write("50-59");
                    builder.InsertCell();
                    builder.Write("60-69");
                    builder.InsertCell();
                    builder.Write("70-79");
                    builder.InsertCell();
                    builder.Write("80-89");
                    builder.InsertCell();
                    builder.Write("90-99");
                    builder.InsertCell();
                    builder.Write("100以上");
                    builder.EndRow();

                    builder.InsertCell();
                    builder.Write($"算術平均"); // 算數

                    int num = 1;
                    foreach (string interval in intervals)
                    {
                        string intervalPname = interval.Split('^')[0];
                        string intervalUserName = interval.Split('^')[1];
                        builder.InsertCell();
                        builder.InsertField($"MERGEFIELD  科目{"定期成績"}{"平均"}{rankType}{intervalPname}   \\* MERGEFORMAT ", $"«R{intervalUserName}»");
                        num++;

                        // builder.InsenumrtField("MERGEFIELD " + rankType + dname + "領域_科目" + i + "F" + r + " \\* MERGEFORMAT ", "«R" + i + "»");
                    }
                    builder.EndRow();

                    builder.EndTable();
                    builder.Writeln();
                }
                builder.Writeln();
                builder.Writeln();

                #endregion
                tempDoc.Save(path, SaveFormat.Doc);

                //System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                //stream.Write(Properties.Resources.新竹評量成績合併欄位總表, 0, Properties.Resources.新竹評量成績合併欄位總表.Length);
                //stream.Flush();
                //stream.Close();
                System.Diagnostics.Process.Start(path);

            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        tempDoc.Save(sd.FileName, SaveFormat.Doc);
                        //System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        //stream.Write(Properties.Resources.新竹評量成績合併欄位總表, 0, Properties.Resources.新竹評量成績合併欄位總表.Length);
                        //stream.Flush();
                        //stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }
    }
}
