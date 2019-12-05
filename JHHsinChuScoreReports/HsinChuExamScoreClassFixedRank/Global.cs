using Aspose.Words;
using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace HsinChuExamScoreClassFixedRank
{
    public class Global
    {
        public const string _UDTTableName = "ischool.HsinChuExamScoreClassFixedRank.configure";

        public static string _ProjectName = "國中新竹班級評量成績單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static string _UserConfTypeName = "使用者選擇設定檔";

        public static string _SelSchoolYear;
        public static string _SelSemester;
        public static string _SelExamID = "";
        public static List<string> _SelStudentIDList = new List<string>();
        public static List<string> _SelClassIDList = new List<string>();
        public static Dictionary<string, List<string>> DomainSubjectDict = new Dictionary<string, List<string>>();

        /// <summary>
        /// 有勾選領域
        /// </summary>
        public static List<string> SelOnlyDomainList = new List<string>();

        /// <summary>
        /// 只有勾選科目
        /// </summary>
        public static List<string> SelOnlySubjectList = new List<string>();


        /// <summary>
        /// 進位四捨五入位數
        /// </summary>
        public static int parseNumebr = 2;

        /// <summary>
        /// 設定領域名稱
        /// </summary>
        public static void SetDomainList()
        {
            DomainNameList.Clear();
            DomainSubjectDict.Clear();

            // 從學生修課動態取得科目領域名稱
            if (_SelClassIDList.Count > 0 && _SelSchoolYear != "" && _SelSemester != "" && _SelExamID != "")
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "SELECT DISTINCT " +
                    "domain" +
                    ",subject " +
                    "FROM " +
                    "sc_attend " +
                    "INNER JOIN " +
                    "course " +
                    "ON sc_attend.ref_course_id=course.id " +
                    "INNER JOIN te_include " +
                    "ON course.ref_exam_template_id = te_include.ref_exam_template_id" +
                    " INNER JOIN student" +
                    " ON sc_attend.ref_student_id = student.id " +
                    "WHERE student.ref_class_id IN(" + string.Join(",", _SelClassIDList.ToArray()) + ") " +
                    "AND student.status = 1" +
                    "AND course.school_year=" + _SelSchoolYear + " " +
                    "AND course.semester=" + _SelSemester + " " +
                    "AND te_include.ref_exam_id = " + _SelExamID + " AND domain <>'';";
                DataTable dt = qh.Select(strSQL);

                foreach (DataRow dr in dt.Rows)
                {
                    string domain = dr["domain"].ToString();
                    string subject = dr["subject"].ToString();
                    if (!DomainNameList.Contains(domain))
                        DomainNameList.Add(domain);

                    if (!DomainSubjectDict.ContainsKey(domain))
                    {
                        DomainSubjectDict.Add(domain, new List<string>());
                    }

                    if (!DomainSubjectDict[domain].Contains(subject))
                        DomainSubjectDict[domain].Add(subject);
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


        public static void ExportMappingFieldWord()
        {
            string inputReportName = "新竹班級評量成績單合併欄位總表";
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

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.新竹班級評量成績單合併欄位總表));

            try
            {
                #region 動態產生合併欄位
                // 讀取總表檔案並動態加入合併欄位

                DocumentBuilder builder = new DocumentBuilder(tempDoc);
                builder.Write("=== 新竹評量成績單合併欄位總表 ===");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("學校名稱");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學校名稱" + " \\* MERGEFORMAT ", "學校名稱" + "");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("學年度");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學年度" + " \\* MERGEFORMAT ", "學年度" + "");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("學期");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學期" + " \\* MERGEFORMAT ", "學期" + "");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("試別名稱");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 試別名稱" + " \\* MERGEFORMAT ", "試別名稱" + "");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("班級");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 班級" + " \\* MERGEFORMAT ", "班級" + "");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("班導師");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 班導師" + " \\* MERGEFORMAT ", "班導師" + "");
                builder.EndRow();


                builder.EndTable();

                builder.Writeln();
                builder.Writeln();

                builder.Writeln("領域學分");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                foreach (string domainName in Global.DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + domainName + "領域學分" + " \\* MERGEFORMAT ", "DC" + "");
                    builder.EndRow();
                }

                builder.EndTable();

                builder.Writeln();
                builder.Writeln();
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                foreach (string domainName in Global.DomainNameList)
                {

                    builder.Writeln(domainName + "領域科目名稱");
                    for (int subj = 1; subj <= 12; subj++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domainName + "領域_科目" + subj + "名稱" + " \\* MERGEFORMAT ", "SN" + "");
                    }
                    builder.EndRow();
                }
                builder.Writeln();
                foreach (string domainName in Global.DomainNameList)
                {

                    builder.Writeln(domainName + "領域科目學分");
                    for (int subj = 1; subj <= 12; subj++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domainName + "領域_科目" + subj + "學分" + " \\* MERGEFORMAT ", "SC" + "");
                    }
                    builder.EndRow();
                }
                builder.EndTable();

                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("分數與排名");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("姓名");
                builder.InsertCell();
                builder.Write("座號");
                builder.InsertCell();
                builder.Write("總分");
                builder.InsertCell();
                builder.Write("加權總分");
                builder.InsertCell();
                builder.Write("平均");
                builder.InsertCell();
                builder.Write("加權平均");
                builder.InsertCell();
                builder.Write("總分班排名");
                builder.InsertCell();
                builder.Write("加權總分班排名");
                builder.InsertCell();
                builder.Write("平均班排名");
                builder.InsertCell();
                builder.Write("加權平均班排名");
                builder.InsertCell();
                builder.Write("總分年排名");
                builder.InsertCell();
                builder.Write("加權總分年排名");
                builder.InsertCell();
                builder.Write("平均年排名");
                builder.InsertCell();
                builder.Write("加權平均年排名");
                builder.EndRow();


                for (int studCot = 1; studCot <= 50; studCot++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 姓名" + studCot + " \\* MERGEFORMAT ", "姓" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 座號" + studCot + " \\* MERGEFORMAT ", "座" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 總分" + studCot + " \\* MERGEFORMAT ", "S" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權總分" + studCot + " \\* MERGEFORMAT ", "SA" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均" + studCot + " \\* MERGEFORMAT ", "A" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均" + studCot + " \\* MERGEFORMAT ", "AA" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 總分班排名" + studCot + " \\* MERGEFORMAT ", "SCR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權總分班排名" + studCot + " \\* MERGEFORMAT ", "SACR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均班排名" + studCot + " \\* MERGEFORMAT ", "ACR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均班排名" + studCot + " \\* MERGEFORMAT ", "AACR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 總分年排名" + studCot + " \\* MERGEFORMAT ", "SYR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權總分年排名" + studCot + " \\* MERGEFORMAT ", "SAYR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均年排名" + studCot + " \\* MERGEFORMAT ", "AYR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均年排名" + studCot + " \\* MERGEFORMAT ", "AAYR" + studCot + "");

                    builder.EndRow();
                }
                builder.EndTable();

                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("分數與排名(定期)");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("姓名");
                builder.InsertCell();
                builder.Write("座號");
                builder.InsertCell();
                builder.Write("總分(定期)");
                builder.InsertCell();
                builder.Write("加權總分(定期)");
                builder.InsertCell();
                builder.Write("平均(定期)");
                builder.InsertCell();
                builder.Write("加權平均(定期)");
                builder.InsertCell();
                builder.Write("總分(定期)班排名");
                builder.InsertCell();
                builder.Write("加權總分(定期)班排名");
                builder.InsertCell();
                builder.Write("平均(定期)班排名");
                builder.InsertCell();
                builder.Write("加權平均(定期)班排名");
                builder.InsertCell();
                builder.Write("總分(定期)年排名");
                builder.InsertCell();
                builder.Write("加權總分(定期)年排名");
                builder.InsertCell();
                builder.Write("平均(定期)年排名");
                builder.InsertCell();
                builder.Write("加權平均(定期)年排名");
                builder.EndRow();


                for (int studCot = 1; studCot <= 50; studCot++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 姓名" + studCot + " \\* MERGEFORMAT ", "姓" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 座號" + studCot + " \\* MERGEFORMAT ", "座" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 總分_定期" + studCot + " \\* MERGEFORMAT ", "S" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權總分_定期" + studCot + " \\* MERGEFORMAT ", "SA" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均_定期" + studCot + " \\* MERGEFORMAT ", "A" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均_定期" + studCot + " \\* MERGEFORMAT ", "AA" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 總分_定期班排名" + studCot + " \\* MERGEFORMAT ", "SCR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權總分_定期班排名" + studCot + " \\* MERGEFORMAT ", "SACR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均_定期班排名" + studCot + " \\* MERGEFORMAT ", "ACR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均_定期班排名" + studCot + " \\* MERGEFORMAT ", "AACR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 總分_定期年排名" + studCot + " \\* MERGEFORMAT ", "SYR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權總分_定期年排名" + studCot + " \\* MERGEFORMAT ", "SAYR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均_定期年排名" + studCot + " \\* MERGEFORMAT ", "AYR" + studCot + "");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均_定期年排名" + studCot + " \\* MERGEFORMAT ", "AAYR" + studCot + "");

                    builder.EndRow();
                }
                builder.EndTable();


                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("領域成績與學分");

                foreach (string domainName in DomainNameList)
                {
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;

                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write("==" + domainName + "領域 ==");
                    builder.InsertCell();
                    builder.Write("領域成績");
                    builder.InsertCell();
                    builder.Write("領域定期成績");
                    builder.InsertCell();
                    builder.Write("領域學分");
                    builder.EndRow();
                    for (int studCot = 1; studCot <= 50; studCot++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domainName + "領域成績" + studCot + " \\* MERGEFORMAT ", "DS" + studCot + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domainName + "領域_定期成績" + studCot + " \\* MERGEFORMAT ", "DS" + studCot + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domainName + "領域學分" + studCot + " \\* MERGEFORMAT ", "DC" + studCot + "");
                        builder.EndRow();
                    }

                    builder.EndTable();
                }

                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("領域科目成績與學分");
                foreach (string domainName in DomainNameList)
                {
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;

                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write("==" + domainName + "領域科目名稱、成績、學分 ==");

                    for (int cot = 1; cot <= 12; cot++)
                    {
                        builder.InsertCell();
                        builder.Write("科目" + cot);
                        builder.InsertCell();
                        builder.Write("成績" + cot);
                        builder.InsertCell();
                        builder.Write("定期成績" + cot);
                        builder.InsertCell();
                        builder.Write("學分" + cot);
                    }
                    builder.EndRow();
                    for (int studCot = 1; studCot <= 50; studCot++)
                    {
                        // 產生12科目可以使用
                        for (int cot = 1; cot <= 12; cot++)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domainName + "領域_科目名稱" + studCot + "_" + cot + " \\* MERGEFORMAT ", "SN" + studCot + "");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domainName + "領域_科目成績" + studCot + "_" + cot + " \\* MERGEFORMAT ", "SS" + studCot + "");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domainName + "領域_科目_定期成績" + studCot + "_" + cot + " \\* MERGEFORMAT ", "SS" + studCot + "");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domainName + "領域_科目學分" + studCot + "_" + cot + " \\* MERGEFORMAT ", "SC" + studCot + "");
                        }
                        builder.EndRow();

                    }

                    builder.EndTable();
                }
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級領域平均");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                foreach (string domainName in Global.DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + domainName + "班級領域平均" + " \\* MERGEFORMAT ", "CDC" + "");
                    builder.EndRow();
                }

                builder.EndTable();

                builder.Writeln();
                builder.Writeln();
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                foreach (string domainName in Global.DomainNameList)
                {

                    builder.Writeln(domainName + "班級領域科目平均");
                    for (int subj = 1; subj <= 12; subj++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domainName + "班級領域_科目平均" + subj + " \\* MERGEFORMAT ", "CSC" + "");
                    }
                    builder.EndRow();
                }
                builder.EndTable();

                builder.Writeln("班級領域及格人數");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                foreach (string domainName in Global.DomainNameList)
                {
                    builder.InsertCell();
                    builder.Write(domainName);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + domainName + "班級領域及格人數" + " \\* MERGEFORMAT ", "CDC" + "");
                    builder.EndRow();
                }

                builder.EndTable();

                builder.Writeln("");
                builder.Writeln("");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;
                foreach (string domainName in Global.DomainNameList)
                {

                    builder.Writeln(domainName + "班級領域科目及格人數");
                    for (int subj = 1; subj <= 12; subj++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domainName + "班級領域_科目及格人數" + subj + " \\* MERGEFORMAT ", "CSC" + "");
                    }
                    builder.EndRow();
                }
                builder.EndTable();


                List<string> rankType1List = new List<string>();
                rankType1List.Add("頂標");
                rankType1List.Add("高標");
                rankType1List.Add("均標");
                rankType1List.Add("低標");
                rankType1List.Add("底標");

                List<string> rankType1VList = new List<string>();
                rankType1VList.Add("avg_top_25");
                rankType1VList.Add("avg_top_50");
                rankType1VList.Add("avg");
                rankType1VList.Add("avg_bottom_50");
                rankType1VList.Add("avg_bottom_25");

                List<string> rankType2List = new List<string>();
                rankType2List.Add("總人數");
                rankType2List.Add("10以下");
                rankType2List.Add("10-19");
                rankType2List.Add("20-29");
                rankType2List.Add("30-39");
                rankType2List.Add("40-49");
                rankType2List.Add("50-59");
                rankType2List.Add("60-69");
                rankType2List.Add("70-79");
                rankType2List.Add("80-89");
                rankType2List.Add("90-99");
                rankType2List.Add("100以上");

                List<string> rankType2VList = new List<string>();
                rankType2VList.Add("matrix_count");
                rankType2VList.Add("level_lt10");
                rankType2VList.Add("level_10");
                rankType2VList.Add("level_20");
                rankType2VList.Add("level_30");
                rankType2VList.Add("level_40");
                rankType2VList.Add("level_50");
                rankType2VList.Add("level_60");
                rankType2VList.Add("level_70");
                rankType2VList.Add("level_80");
                rankType2VList.Add("level_90");
                rankType2VList.Add("level_gte100");

                // 班級領域成績五標
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級領域成績五標");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 班級領域成績五標");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();

                    // 班級
                    for (int cla = 1; cla <= 20; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_" + dName + "_領域成績_" + rk1 + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 班級領域定期成績五標
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級領域定期成績五標");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 班級領域定期成績五標");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();

                    // 班級
                    for (int cla = 1; cla <= 20; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_" + dName + "_領域定期成績_" + rk1 + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }

                // 年級領域成績五標              
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級領域成績五標");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 年級領域成績五標");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();

                    // 年級
                    for (int cla = 1; cla <= 12; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_" + dName + "_領域成績_" + rk1 + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 年級領域定期成績五標
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級領域定期成績五標");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 年級領域定期成績五標");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();

                    // 年級
                    for (int cla = 1; cla <= 12; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_" + dName + "_領域定期成績_" + rk1 + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 班級領域成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級領域成績組距");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 班級領域成績組距");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();

                    // 班級
                    for (int cla = 1; cla <= 20; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_" + dName + "_領域成績_" + rk2 + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 班級領域定期成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級領域定期成績組距");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 班級領域定期成績組距");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();

                    // 班級
                    for (int cla = 1; cla <= 20; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_" + dName + "_領域定期成績_" + rk2 + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }

                // 年級領域成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級領域成績組距");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 年級領域成績組距");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();

                    // 年級
                    for (int cla = 1; cla <= 12; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_" + dName + "_領域成績_" + rk2 + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }

                // 年級領域定期成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級領域定期成績組距");

                foreach (string dName in Global.DomainNameList)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.Write(dName + " 年級領域定期成績組距");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();

                    // 年級
                    for (int cla = 1; cla <= 12; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_" + dName + "_領域定期成績_" + rk2 + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 班級科目成績五標
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級科目成績五標");
                for (int subjCot = 1; subjCot <= 7; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();
                    // 班級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_科目成績_" + rk1 + subjCot + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }



                // 班級科目定期成績五標     
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級科目定期成績五標");
                for (int subjCot = 1; subjCot <= 10; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();
                    // 班級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_科目定期成績_" + rk1 + subjCot + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 年級科目成績五標
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級科目成績五標");
                for (int subjCot = 1; subjCot <= 7; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();
                    // 年級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_科目成績_" + rk1 + subjCot + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 年級科目定期成績五標
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級科目定期成績五標");
                for (int subjCot = 1; subjCot <= 7; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk1 in rankType1List)
                    {
                        builder.InsertCell();
                        builder.Write(rk1);
                    }
                    builder.EndRow();
                    // 年級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk1 in rankType1VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_科目定期成績_" + rk1 + subjCot + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }

                // 班級科目成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級科目成績組距");
                for (int subjCot = 1; subjCot <= 7; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();
                    // 班級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_科目成績_" + rk2 + subjCot + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }

                // 班級科目定期成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("班級科目定期成績組距");
                for (int subjCot = 1; subjCot <= 7; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("班級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();
                    // 班級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 班級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 班級" + cla + "_科目定期成績_" + rk2 + subjCot + " \\* MERGEFORMAT ", "CR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 年級科目成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級科目成績組距");
                for (int subjCot = 1; subjCot <= 7; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();
                    // 班級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_科目成績_" + rk2 + subjCot + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }


                // 年級科目定期成績組距
                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("年級科目定期成績組距");
                for (int subjCot = 1; subjCot <= 7; subjCot++)
                {
                    builder.Writeln("");
                    builder.Writeln("");
                    builder.StartTable();
                    builder.CellFormat.Borders.LineStyle = LineStyle.None;
                    builder.InsertCell();
                    builder.Write("年級");
                    builder.InsertCell();
                    builder.Write("科目" + subjCot);
                    foreach (string rk2 in rankType2List)
                    {
                        builder.InsertCell();
                        builder.Write(rk2);
                    }
                    builder.EndRow();
                    // 班級
                    for (int cla = 1; cla <= 10; cla++)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_名稱" + " \\* MERGEFORMAT ", "N" + "");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 年級" + cla + "_科目名稱_" + subjCot + " \\* MERGEFORMAT ", "N" + "");
                        foreach (string rk2 in rankType2VList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD 年級" + cla + "_科目定期成績_" + rk2 + subjCot + " \\* MERGEFORMAT ", "GR" + "");
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }

                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("領域名稱與科目名稱");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                builder.InsertCell();
                builder.Write("學生領域名稱");
                builder.InsertCell();
                builder.Write("學生科目名稱");
                builder.EndRow();

                for (int r = 1; r <= 30; r++)
                {
                    builder.InsertCell();
                    builder.Write("" + r);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學生_領域名稱" + r + " \\* MERGEFORMAT ", "領");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學生_科目名稱" + r + " \\* MERGEFORMAT ", "科");
                    builder.EndRow();
                }
                builder.EndTable();


                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("學生領域分開成績");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                for (int studCot = 1; studCot <= 50; studCot++)
                {
                    builder.InsertCell();
                    builder.Write("" + studCot);
                    for (int ss = 1; ss <= 10; ss++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_領域名稱" + studCot + "_" + ss + " \\* MERGEFORMAT ", "N" + ss);
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_領域成績" + studCot + "_" + ss + " \\* MERGEFORMAT ", "S" + ss);
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_領域_定期成績" + studCot + "_" + ss + " \\* MERGEFORMAT ", "S" + ss);
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_領域學分" + studCot + "_" + ss + " \\* MERGEFORMAT ", "C" + ss);

                    }
                    builder.EndRow();
                }
                builder.EndTable();

                builder.Writeln("");
                builder.Writeln("");
                builder.Writeln("學生科目分開成績");
                builder.StartTable();
                builder.CellFormat.Borders.LineStyle = LineStyle.None;

                for (int studCot = 1; studCot <= 50; studCot++)
                {
                    builder.InsertCell();
                    builder.Write("" + studCot);
                    for (int ss = 1; ss <= 10; ss++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_科目名稱" + studCot + "_" + ss + " \\* MERGEFORMAT ", "N" + ss);
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_科目成績" + studCot + "_" + ss + " \\* MERGEFORMAT ", "S" + ss);
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_科目_定期成績" + studCot + "_" + ss + " \\* MERGEFORMAT ", "S" + ss);
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD 學生_科目學分" + studCot + "_" + ss + " \\* MERGEFORMAT ", "C" + ss);

                    }
                    builder.EndRow();
                }
                builder.EndTable();

                #endregion

                tempDoc.Save(path, SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);

            }
            catch (Exception ex)
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
        }
    }
}
