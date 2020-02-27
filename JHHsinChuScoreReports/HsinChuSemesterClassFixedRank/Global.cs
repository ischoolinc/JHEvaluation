using Aspose.Words;
using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace HsinChuSemesterClassFixedRank
{
    public class Global
    {
        public const string _UDTTableName = "ischool.HsinChuSemesterScoreClassFixedRank.configure";

        public static string _ProjectName = "國中新竹班級學期成績單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static string _UserConfTypeName = "使用者選擇設定檔";

        public static string _SelSchoolYear;
        public static string _SelSemester;
        public static string _SelExamID = "";
        public static List<string> _SelStudentIDList = new List<string>();
        public static List<string> _SelClassIDList = new List<string>();
        public static Dictionary<string, List<string>> DomainSubjectDict = new Dictionary<string, List<string>>();

        public static List<string> domainNameList = new List<string>();

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

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.新竹學期成績單合併欄位總表));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(tempDoc);
            builder.MoveToDocumentEnd();
            builder.Writeln();
            builder.Writeln();


            // 2020/2/21 與宏安討論，先將學生各別五標、組距先移除，因為在實作上沒有使用到，節省空間。

            // 產生動態合併欄位總表
            List<string> r2aList = new List<string>();
            //List<string> r2bList = new List<string>();
            r2aList.Add("rank");
            r2aList.Add("matrix_count");
            r2aList.Add("pr");
            r2aList.Add("percentile");
            //r2aList.Add("avg_top_25");
            //r2aList.Add("avg_top_50");
            //r2aList.Add("avg");
            //r2aList.Add("avg_bottom_50");
            //r2aList.Add("avg_bottom_25");

            //r2bList.Add("level_gte100");
            //r2bList.Add("level_90");
            //r2bList.Add("level_80");
            //r2bList.Add("level_70");
            //r2bList.Add("level_60");
            //r2bList.Add("level_50");
            //r2bList.Add("level_40");
            //r2bList.Add("level_30");
            //r2bList.Add("level_20");
            //r2bList.Add("level_10");
            //r2bList.Add("level_lt10");

            // 班級排名
            List<string> cr2aList = new List<string>();
            List<string> cr2bList = new List<string>();
            cr2aList.Add("matrix_count");
            cr2aList.Add("avg_top_25");
            cr2aList.Add("avg_top_50");
            cr2aList.Add("avg");
            cr2aList.Add("avg_bottom_50");
            cr2aList.Add("avg_bottom_25");

            cr2bList.Add("level_gte100");
            cr2bList.Add("level_90");
            cr2bList.Add("level_80");
            cr2bList.Add("level_70");
            cr2bList.Add("level_60");
            cr2bList.Add("level_50");
            cr2bList.Add("level_40");
            cr2bList.Add("level_30");
            cr2bList.Add("level_20");
            cr2bList.Add("level_10");
            cr2bList.Add("level_lt10");


            #region 學生-總計成績-排名、PR、五標

            List<string> r1List = new List<string>();
            r1List.Add("課程學習總成績");
            r1List.Add("學習領域總成績");
            r1List.Add("課程學習總成績(原始)");
            r1List.Add("學習領域總成績(原始)");

            List<string> rkList = new List<string>();
            rkList.Add("班排名");
            rkList.Add("年排名");
            rkList.Add("類別1排名");
            rkList.Add("類別2排名");

            List<string> cr2aNameList = new List<string>();
            cr2aNameList.Add("總人數");
            cr2aNameList.Add("頂標");
            cr2aNameList.Add("高標");
            cr2aNameList.Add("均標");
            cr2aNameList.Add("低標");
            cr2aNameList.Add("底標");


            List<string> cr2bNameList = new List<string>();
            cr2bNameList.Add("100以上");
            cr2bNameList.Add("90以上小於100");
            cr2bNameList.Add("80以上小於90");
            cr2bNameList.Add("70以上小於80");
            cr2bNameList.Add("60以上小於70");
            cr2bNameList.Add("50以上小於60");
            cr2bNameList.Add("40以上小於50");
            cr2bNameList.Add("30以上小於40");
            cr2bNameList.Add("20以上小於30");
            cr2bNameList.Add("10以上小於20");
            cr2bNameList.Add("10以下");

            #region 學生總計成績 排名、五標
            foreach (string r1 in r1List)
            {

                // 排名類別
                foreach (string rk in rkList)
                {
                    // 學生 課程學習總成績
                    builder.Writeln("");
                    builder.Writeln("學生-" + r1 + " " + rk + "  排名、PR、百分比");
                    builder.StartTable();
                    builder.InsertCell(); builder.Write("座號");
                    builder.InsertCell(); builder.Write("姓名");
                    builder.InsertCell(); builder.Write("分數");
                    builder.InsertCell();
                    builder.Write("排名");
                    builder.InsertCell();
                    builder.Write("總人數");
                    builder.InsertCell();
                    builder.Write("PR");
                    builder.InsertCell();
                    builder.Write("百分比");
                    //builder.InsertCell();
                    //builder.Write("頂標");
                    //builder.InsertCell();
                    //builder.Write("高標");
                    //builder.InsertCell();
                    //builder.Write("均標");
                    //builder.InsertCell();
                    //builder.Write("低標");
                    //builder.InsertCell();
                    //builder.Write("底標");
                    builder.EndRow();

                    for (int si = 1; si <= 50; si++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "座號" + si + " \\* MERGEFORMAT ", "«" + "座" + si + "»");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "姓名" + si + " \\* MERGEFORMAT ", "«" + "姓" + si + "»");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "學生" + si + "_" + r1 + "_成績" + " \\* MERGEFORMAT ", "«" + "S" + si + "»");

                        foreach (string ra in r2aList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + "學生" + si + "_" + r1 + "_" + rk + "_" + ra + " \\* MERGEFORMAT ", "«" + "R" + si + "»");
                        }
                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln("");
                }
            }
            #endregion

            //#region 學生總計成績 組距
            //foreach (string r1 in r1List)
            //{

            //    // 排名類別
            //    foreach (string rk in rkList)
            //    {
            //        // 學生 課程學習總成績
            //        builder.Writeln("");
            //        builder.Writeln("學生-" + r1 + " " + rk + "  組距");
            //        builder.StartTable();
            //        builder.InsertCell(); builder.Write("座號");
            //        builder.InsertCell(); builder.Write("姓名");
            //        builder.InsertCell(); builder.Write("分數");
            //        builder.InsertCell(); builder.Write("100以上");
            //        builder.InsertCell(); builder.Write("90以上小於100");
            //        builder.InsertCell(); builder.Write("80以上小於90");
            //        builder.InsertCell(); builder.Write("70以上小於80");
            //        builder.InsertCell(); builder.Write("60以上小於70");
            //        builder.InsertCell(); builder.Write("50以上小於60");
            //        builder.InsertCell(); builder.Write("40以上小於50");
            //        builder.InsertCell(); builder.Write("30以上小於40");
            //        builder.InsertCell(); builder.Write("20以上小於30");
            //        builder.InsertCell(); builder.Write("10以上小於20");
            //        builder.InsertCell(); builder.Write("10以下");
            //        builder.EndRow();

            //        for (int si = 1; si <= 50; si++)
            //        {
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "座號" + si + " \\* MERGEFORMAT ", "«" + "座" + si + "»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "姓名" + si + " \\* MERGEFORMAT ", "«" + "姓" + si + "»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + si + "_" + r1 + "_成績" + " \\* MERGEFORMAT ", "«" + "S" + si + "»");

            //            foreach (string ra in r2bList)
            //            {
            //                builder.InsertCell();
            //                builder.InsertField("MERGEFIELD " + "學生" + si + "_" + r1 + "_" + rk + "_" + ra + " \\* MERGEFORMAT ", "«" + "R" + si + "»");
            //            }
            //            builder.EndRow();
            //        }

            //        builder.EndTable();
            //        builder.Writeln("");
            //    }
            //}
            //#endregion

            #endregion

            //#region 學生-領域成績 領域1,2,3 成績、排名、百分比、
            //for (int dd = 1; dd <= 12; dd++)
            //{
            //    builder.Writeln("學生-領域" + dd + "成績相關");
            //    builder.StartTable();
            //    builder.InsertCell(); builder.Write("座號");
            //    builder.InsertCell(); builder.Write("姓名");
            //    builder.InsertCell(); builder.Write("名稱");
            //    builder.InsertCell(); builder.Write("學分");
            //    builder.InsertCell(); builder.Write("成績");
            //    builder.InsertCell(); builder.Write("成績(原始)");
            //    builder.InsertCell(); builder.Write("班排名");
            //    builder.InsertCell(); builder.Write("年排名");
            //    builder.InsertCell(); builder.Write("類別1排名");
            //    builder.InsertCell(); builder.Write("類別2排名");
            //    builder.InsertCell(); builder.Write("班排名(原始)");
            //    builder.InsertCell(); builder.Write("年排名(原始)");
            //    builder.InsertCell(); builder.Write("類別1排名(原始)");
            //    builder.InsertCell(); builder.Write("類別2排名(原始)");
            //    builder.InsertCell(); builder.Write("班排名百分比");
            //    builder.InsertCell(); builder.Write("年排名百分比");
            //    builder.InsertCell(); builder.Write("類別1排名百分比");
            //    builder.InsertCell(); builder.Write("類別2排名百分比");
            //    builder.InsertCell(); builder.Write("班排名百分比(原始)");
            //    builder.InsertCell(); builder.Write("年排名百分比(原始)");
            //    builder.InsertCell(); builder.Write("類別1排名百分比(原始)");
            //    builder.InsertCell(); builder.Write("類別2排名百分比(原始)");
            //    builder.EndRow();

            //    for (int studIdx = 1; studIdx <= 50; studIdx++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_名稱" + " \\* MERGEFORMAT ", "«N»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_學分" + " \\* MERGEFORMAT ", "«C»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_成績" + " \\* MERGEFORMAT ", "«S»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_成績" + " \\* MERGEFORMAT ", "«S»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_班排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_年排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_類別1排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_類別2排名_rank" + " \\* MERGEFORMAT ", "«R»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_班排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_年排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_類別1排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_類別2排名_rank" + " \\* MERGEFORMAT ", "«R»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_班排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_年排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_類別1排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_類別2排名_percentile" + " \\* MERGEFORMAT ", "«P»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_班排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_年排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_類別1排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)" + dd + "_類別2排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.EndRow();
            //    }

            //    builder.EndTable();
            //    builder.Writeln();
            //}

            //#endregion

            //#region 學生-科目成績 科目1,2,3 成績、排名、百分比
            //for (int dd = 1; dd <= 16; dd++)
            //{
            //    builder.Writeln("學生-科目" + dd + "成績相關");
            //    builder.StartTable();
            //    builder.InsertCell(); builder.Write("座號");
            //    builder.InsertCell(); builder.Write("姓名");
            //    builder.InsertCell(); builder.Write("名稱");
            //    builder.InsertCell(); builder.Write("學分");
            //    builder.InsertCell(); builder.Write("成績");
            //    builder.InsertCell(); builder.Write("成績(原始)");
            //    builder.InsertCell(); builder.Write("班排名");
            //    builder.InsertCell(); builder.Write("年排名");
            //    builder.InsertCell(); builder.Write("類別1排名");
            //    builder.InsertCell(); builder.Write("類別2排名");
            //    builder.InsertCell(); builder.Write("班排名(原始)");
            //    builder.InsertCell(); builder.Write("年排名(原始)");
            //    builder.InsertCell(); builder.Write("類別1排名(原始)");
            //    builder.InsertCell(); builder.Write("類別2排名(原始)");
            //    builder.InsertCell(); builder.Write("班排名百分比");
            //    builder.InsertCell(); builder.Write("年排名百分比");
            //    builder.InsertCell(); builder.Write("類別1排名百分比");
            //    builder.InsertCell(); builder.Write("類別2排名百分比");
            //    builder.InsertCell(); builder.Write("班排名百分比(原始)");
            //    builder.InsertCell(); builder.Write("年排名百分比(原始)");
            //    builder.InsertCell(); builder.Write("類別1排名百分比(原始)");
            //    builder.InsertCell(); builder.Write("類別2排名百分比(原始)");
            //    builder.EndRow();

            //    for (int studIdx = 1; studIdx <= 50; studIdx++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");
            //        builder.InsertCell();
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«N»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_學分" + " \\* MERGEFORMAT ", "«C»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_成績" + " \\* MERGEFORMAT ", "«S»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_成績" + " \\* MERGEFORMAT ", "«S»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_班排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_年排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_類別1排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_類別2排名_rank" + " \\* MERGEFORMAT ", "«R»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_班排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_年排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_類別1排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_類別2排名_rank" + " \\* MERGEFORMAT ", "«R»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_班排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_年排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_類別1排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_類別2排名_percentile" + " \\* MERGEFORMAT ", "«P»");

            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_班排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_年排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_類別1排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)" + dd + "_類別2排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //        builder.EndRow();
            //    }

            //    builder.EndTable();
            //    builder.Writeln();
            //}

            //#endregion

            //#region 學生-領域科目成績 領域-科目  成績、排名、百分比
            //foreach (string dName in domainNameList)
            //{
            //    for (int dd = 1; dd <= 5; dd++)
            //    {
            //        builder.Writeln("學生" + dName + "領域-科目" + dd + "成績相關");
            //        builder.StartTable();
            //        builder.InsertCell(); builder.Write("座號");
            //        builder.InsertCell(); builder.Write("姓名");
            //        builder.InsertCell(); builder.Write("名稱");
            //        builder.InsertCell(); builder.Write("學分");
            //        builder.InsertCell(); builder.Write("成績");
            //        builder.InsertCell(); builder.Write("成績(原始)");
            //        builder.InsertCell(); builder.Write("班排名");
            //        builder.InsertCell(); builder.Write("年排名");
            //        builder.InsertCell(); builder.Write("類別1排名");
            //        builder.InsertCell(); builder.Write("類別2排名");
            //        builder.InsertCell(); builder.Write("班排名(原始)");
            //        builder.InsertCell(); builder.Write("年排名(原始)");
            //        builder.InsertCell(); builder.Write("類別1排名(原始)");
            //        builder.InsertCell(); builder.Write("類別2排名(原始)");
            //        builder.InsertCell(); builder.Write("班排名百分比");
            //        builder.InsertCell(); builder.Write("年排名百分比");
            //        builder.InsertCell(); builder.Write("類別1排名百分比");
            //        builder.InsertCell(); builder.Write("類別2排名百分比");
            //        builder.InsertCell(); builder.Write("班排名百分比(原始)");
            //        builder.InsertCell(); builder.Write("年排名百分比(原始)");
            //        builder.InsertCell(); builder.Write("類別1排名百分比(原始)");
            //        builder.InsertCell(); builder.Write("類別2排名百分比(原始)");
            //        builder.EndRow();

            //        for (int studIdx = 1; studIdx <= 50; studIdx++)
            //        {
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«N»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_學分" + " \\* MERGEFORMAT ", "«C»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_成績" + " \\* MERGEFORMAT ", "«S»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_成績" + " \\* MERGEFORMAT ", "«S»");

            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_班排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_年排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_類別1排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_類別2排名_rank" + " \\* MERGEFORMAT ", "«R»");

            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_班排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_年排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_類別1排名_rank" + " \\* MERGEFORMAT ", "«R»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_類別2排名_rank" + " \\* MERGEFORMAT ", "«R»");

            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_班排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_年排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_類別1排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域_科目" + dd + "_類別2排名_percentile" + " \\* MERGEFORMAT ", "«P»");

            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_班排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_年排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_類別1排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_" + dName + "領域(原始)_科目" + dd + "_類別2排名_percentile" + " \\* MERGEFORMAT ", "«P»");
            //            builder.EndRow();
            //        }

            //        builder.EndTable();
            //        builder.Writeln();
            //    }
            //}
            //#endregion



            #region 班級-總計成績 五標、組距
            #region 班級總計成績 五標
            foreach (string r1 in r1List)
            {
                // 排名類別
                foreach (string rk in rkList)
                {
                    // 班級 課程學習總成績
                    builder.Writeln("");
                    builder.Writeln("班級-" + r1 + " " + rk + "五標");
                    builder.StartTable();

                    builder.InsertCell();
                    builder.Write("總人數");

                    builder.InsertCell();
                    builder.Write("頂標");
                    builder.InsertCell();
                    builder.Write("高標");
                    builder.InsertCell();
                    builder.Write("均標");
                    builder.InsertCell();
                    builder.Write("低標");
                    builder.InsertCell();
                    builder.Write("底標");
                    builder.EndRow();

                    foreach (string ra in cr2aList)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_" + r1 + "_" + rk + "_" + ra + " \\* MERGEFORMAT ", "«" + "R" + "»");
                    }
                    builder.EndRow();


                    builder.EndTable();

                }
            }
            #endregion

            #region 班級總計成績 組距
            foreach (string r1 in r1List)
            {

                // 排名類別
                foreach (string rk in rkList)
                {
                    // 班級 課程學習總成績
                    builder.Writeln("");
                    builder.Writeln("班級-" + r1 + " " + rk + "  組距");
                    builder.StartTable();

                    builder.InsertCell(); builder.Write("100以上");
                    builder.InsertCell(); builder.Write("90以上小於100");
                    builder.InsertCell(); builder.Write("80以上小於90");
                    builder.InsertCell(); builder.Write("70以上小於80");
                    builder.InsertCell(); builder.Write("60以上小於70");
                    builder.InsertCell(); builder.Write("50以上小於60");
                    builder.InsertCell(); builder.Write("40以上小於50");
                    builder.InsertCell(); builder.Write("30以上小於40");
                    builder.InsertCell(); builder.Write("20以上小於30");
                    builder.InsertCell(); builder.Write("10以上小於20");
                    builder.InsertCell(); builder.Write("10以下");
                    builder.EndRow();


                    foreach (string ra in cr2bList)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_" + r1 + "_" + rk + "_" + ra + " \\* MERGEFORMAT ", "«" + "R" + "»");
                    }
                    builder.EndRow();

                    builder.EndTable();

                }
            }
            #endregion
            #endregion

            #region 班級-領域成績 領域1,2,3 五標、組距
            #region 班級領域成績1,2,3 五標
            // 排名類別
            foreach (string rk in rkList)
            {
                // 班級 各領域成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各領域成績 " + rk + "五標");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");
                foreach (string na in cr2aNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();


                for (int dd = 1; dd <= 15; dd++)
                {
                    // 班級_領域1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_領域" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                    int colIdx = 0;
                    foreach (string na in cr2aNameList)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_領域" + dd + "_" + rk + "_" + cr2aList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                        colIdx++;
                    }
                    builder.EndRow();
                }
                builder.EndTable();
            }

            foreach (string rk in rkList)
            {
                // 班級 各領域成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各領域成績 (原始)" + rk + "五標");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");
                foreach (string na in cr2aNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();

                for (int dd = 1; dd <= 15; dd++)
                {
                    // 班級_領域(原始)1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_領域" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                    int colIdx = 0;
                    foreach (string na in cr2aNameList)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_領域(原始)" + dd + "_" + rk + "_" + cr2aList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                        colIdx++;
                    }

                    builder.EndRow();
                }
                builder.EndTable();
            }

            #endregion


            #region 班級領域成績1,2,3 組距
            // 排名類別
            foreach (string rk in rkList)
            {
                // 班級 課程學習總成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各領域成績 " + rk + "組距");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");

                foreach (string na in cr2bNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();

                for (int dd = 1; dd <= 15; dd++)
                {
                    // 班級_領域1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_領域" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                    int colIdx = 0;
                    foreach (string na in cr2bNameList)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_領域" + dd + "_" + rk + "_" + cr2bList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");

                        colIdx++;
                    }
                    builder.EndRow();
                }
                builder.EndTable();
            }

            foreach (string rk in rkList)
            {
                // 班級 各領域成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各領域成績 (原始) " + rk + "組距");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");
                foreach (string na in cr2bNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();

                for (int dd = 1; dd <= 15; dd++)
                {
                    // 班級_領域(原始)1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_領域" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                    int colIdx = 0;
                    foreach (string na in cr2bNameList)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_領域(原始)" + dd + "_" + rk + "_" + cr2bList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                        colIdx++;
                    }
                    builder.EndRow();
                }
                builder.EndTable();
            }

            #endregion
            #endregion

            #region 班級-科目成績 科目1,2,3 五標、組距

            #region 班級科目成績1,2,3 五標
            // 排名類別
            foreach (string rk in rkList)
            {
                // 班級 科目成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各科目成績 " + rk + "五標");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");
                foreach (string na in cr2aNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();
                for (int dd = 1; dd <= 20; dd++)
                {
                    // 班級_科目1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");
                    int colIdx = 0;
                    foreach (string na in cr2aNameList)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_科目" + dd + "_" + rk + "_" + cr2aList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                        colIdx++;
                    }
                    builder.EndRow();

                }
                builder.EndTable();
            }

            // 排名類別
            foreach (string rk in rkList)
            {
                // 班級 科目成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各科目成績(原始) " + rk + "五標");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");

                foreach (string na in cr2aNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();

                for (int dd = 1; dd <= 20; dd++)
                {
                    // 班級_科目(原始)1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                    int colIdx = 0;
                    foreach (string na in cr2aNameList)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_科目(原始)" + dd + "_" + rk + "_" + cr2aList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                        colIdx++;
                    }
                    builder.EndRow();
                }
                builder.EndTable();
            }


            #endregion

            #region 班級科目成績1,2,3 組距
            // 排名類別
            foreach (string rk in rkList)
            {
                // 班級 科目成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各科目成績 " + rk + "組距");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");
                foreach (string na in cr2bNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();

                for (int dd = 1; dd <= 20; dd++)
                {

                    // 班級_科目1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                    int colIdx = 0;
                    foreach (string na in cr2bNameList)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_科目" + dd + "_" + rk + "_" + cr2bList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                        colIdx++;
                    }
                    builder.EndRow();
                }
                builder.EndTable();
            }

            foreach (string rk in rkList)
            {
                // 班級 科目成績
                builder.Writeln("");
                builder.Writeln("班級-" + "各科目成績 (原始) " + rk + "組距");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("項目名稱");
                foreach (string na in cr2bNameList)
                {
                    builder.InsertCell();
                    builder.Write(na);
                }
                builder.EndRow();

                for (int dd = 1; dd <= 20; dd++)
                {
                    // 班級_科目(原始)1_名稱
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級" + "_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                    int colIdx = 0;
                    foreach (string na in cr2bNameList)
                    {

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級" + "_科目(原始)" + dd + "_" + rk + "_" + cr2bList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                    }
                    builder.EndRow();
                    colIdx++;
                }
                builder.EndTable();
            }
            #endregion

            #endregion


            #region 班級-領域科目成績 領域-科目 1,2,3 五標、組距

            #region 班級科目成績1,2,3 五標
            // 排名類別
            foreach (string rk in rkList)
            {
                foreach (string dName in domainNameList)
                {
                    // 班級 科目成績
                    builder.Writeln("");
                    builder.Writeln("班級-" + dName + "領域各科目成績 " + rk + "五標");
                    builder.StartTable();

                    builder.InsertCell();
                    builder.Write("項目名稱");
                    foreach (string na in cr2aNameList)
                    {
                        builder.InsertCell();
                        builder.Write(na);
                    }
                    builder.EndRow();

                    for (int dd = 1; dd <= 7; dd++)
                    {
                        // 班級_科目1_名稱
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                        int colIdx = 0;
                        foreach (string na in cr2aNameList)
                        {

                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + dd + "_" + rk + "_" + cr2aList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                            colIdx++;
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }
            }


            foreach (string rk in rkList)
            {
                foreach (string dName in domainNameList)
                {
                    // 班級 科目成績
                    builder.Writeln("");
                    builder.Writeln("班級-" + dName + "領域各科目成績(原始) " + rk + "五標");
                    builder.StartTable();

                    builder.InsertCell();
                    builder.Write("項目名稱");

                    foreach (string na in cr2aNameList)
                    {
                        builder.InsertCell();
                        builder.Write(na);
                    }
                    builder.EndRow();


                    for (int dd = 1; dd <= 7; dd++)
                    {
                        // 班級_科目1_名稱
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                        int colIdx = 0;
                        foreach (string na in cr2aNameList)
                        {

                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域(原始)_科目" + dd + "_" + rk + "_" + cr2aList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                            colIdx++;
                        }
                        builder.EndRow();

                    }
                    builder.EndTable();
                }
            }
            #endregion


            #region 班級科目成績1,2,3 組距
            // 排名類別
            foreach (string rk in rkList)
            {
                foreach (string dName in domainNameList)
                {
                    // 班級 科目成績
                    builder.Writeln("");
                    builder.Writeln("班級-" + dName + "領域各科目成績 " + rk + "組距");
                    builder.StartTable();

                    builder.InsertCell();
                    builder.Write("項目名稱");
                    foreach (string na in cr2bNameList)
                    {
                        builder.InsertCell();
                        builder.Write(na);
                    }
                    builder.EndRow();

                    for (int dd = 1; dd <= 7; dd++)
                    {
                        // 班級_科目1_名稱
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                        int colIdx = 0;
                        foreach (string na in cr2bNameList)
                        {

                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + dd + "_" + rk + "_" + cr2bList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                            colIdx++;
                        }
                        builder.EndRow();

                    }
                    builder.EndTable();
                }
            }



            foreach (string rk in rkList)
            {
                foreach (string dName in domainNameList)
                {
                    // 班級 科目成績
                    builder.Writeln("");
                    builder.Writeln("班級-" + dName + "領域各科目成績(原始) " + rk + "組距");
                    builder.StartTable();

                    builder.InsertCell();
                    builder.Write("項目名稱");
                    foreach (string na in cr2bNameList)
                    {
                        builder.InsertCell();
                        builder.Write(na);
                    }
                    builder.EndRow();

                    for (int dd = 1; dd <= 7; dd++)
                    {
                        // 班級_科目1_名稱
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域(原始)_科目" + dd + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + dd + "»");

                        int colIdx = 0;
                        foreach (string na in cr2bNameList)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域(原始)_科目" + dd + "_" + rk + "_" + cr2bList[colIdx] + " \\* MERGEFORMAT ", "«" + "R" + dd + "»");
                            colIdx++;
                        }
                        builder.EndRow();
                    }
                    builder.EndTable();
                }
            }

            #endregion

            #endregion


            #region 領域成績標示
            builder.Writeln("");
            builder.Writeln("領域成績標示：");

            for (int i = 1; i <= 13; i += 3)
            {
                builder.StartTable();
                builder.InsertCell(); builder.Write("座號");
                builder.InsertCell(); builder.Write("姓名");
                for (int dd = i; dd <= (i + 2); dd++)
                {
                    builder.InsertCell();
                    builder.Write("領域" + dd + "_需補考標示");
                    builder.InsertCell();
                    builder.Write("領域" + dd + "_補考成績標示");
                    builder.InsertCell();
                    builder.Write("領域" + dd + "_不及格標示");
                }
                builder.EndRow();
                for (int studIdx = 1; studIdx <= 50; studIdx++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");
                    for (int dd = i; dd <= (i + 2); dd++)
                    {
                        // 學生1_領域1_需補考標示
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_需補考標示" + " \\* MERGEFORMAT ", "«C" + studIdx + "»");

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_補考成績標示" + " \\* MERGEFORMAT ", "«C" + studIdx + "»");


                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域" + dd + "_不及格標示" + " \\* MERGEFORMAT ", "«C" + studIdx + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
            }

            #endregion

            #region 科目成績標示
            builder.Writeln("科目成績標示：");

            for (int i = 1; i <= 19; i += 3)
            {
                builder.StartTable();
                builder.InsertCell(); builder.Write("座號");
                builder.InsertCell(); builder.Write("姓名");
                for (int dd = i; dd <= (i + 2); dd++)
                {
                    builder.InsertCell();
                    builder.Write("科目" + dd + "_需補考標示");
                    builder.InsertCell();
                    builder.Write("科目" + dd + "_補考成績標示");
                    builder.InsertCell();
                    builder.Write("科目" + dd + "_不及格標示");
                }
                builder.EndRow();
                for (int studIdx = 1; studIdx <= 50; studIdx++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");
                    for (int dd = i; dd <= (i + 2); dd++)
                    {
                        // 學生1_科目1_需補考標示
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_需補考標示" + " \\* MERGEFORMAT ", "«C" + studIdx + "»");

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_補考成績標示" + " \\* MERGEFORMAT ", "«C" + studIdx + "»");


                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目" + dd + "_不及格標示" + " \\* MERGEFORMAT ", "«C" + studIdx + "»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();
            }

            #endregion


            #region 班級領域1,2,3 平均、及格人數
            builder.Writeln();
            builder.Writeln();
            builder.Writeln("班級領域平均、及格人數");
            builder.StartTable();
            builder.InsertCell(); builder.Write("名稱");
            builder.InsertCell(); builder.Write("學分");
            builder.InsertCell(); builder.Write("平均");
            builder.InsertCell(); builder.Write("平均(原始)");
            builder.InsertCell(); builder.Write("及格人數");
            builder.InsertCell(); builder.Write("及格人數(原始)");
            builder.EndRow();

            for (int rowIdx = 1; rowIdx <= 15; rowIdx++)
            {
                // 班級_領域1_名稱                
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_領域" + rowIdx + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_領域" + rowIdx + "_學分" + " \\* MERGEFORMAT ", "«" + "C" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_領域" + rowIdx + "_平均" + " \\* MERGEFORMAT ", "«" + "S" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_領域(原始)" + rowIdx + "_平均" + " \\* MERGEFORMAT ", "«" + "S" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_領域" + rowIdx + "_及格人數" + " \\* MERGEFORMAT ", "«" + "P" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_領域(原始)" + rowIdx + "_及格人數" + " \\* MERGEFORMAT ", "«" + "P" + rowIdx + "»");
                builder.EndRow();
            }

            builder.EndTable();
            #endregion

            #region 班級科目1,2,3 平均、及格人數
            builder.Writeln();
            builder.Writeln();
            builder.Writeln("班級科目平均、及格人數");
            builder.StartTable();
            builder.InsertCell(); builder.Write("名稱");
            builder.InsertCell(); builder.Write("學分");
            builder.InsertCell(); builder.Write("平均");
            builder.InsertCell(); builder.Write("平均(原始)");
            builder.InsertCell(); builder.Write("及格人數");
            builder.InsertCell(); builder.Write("及格人數(原始)");
            builder.EndRow();

            for (int rowIdx = 1; rowIdx <= 20; rowIdx++)
            {
                // 班級_科目1_名稱                
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_科目" + rowIdx + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_科目" + rowIdx + "_學分" + " \\* MERGEFORMAT ", "«" + "C" + rowIdx + "»");

                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_科目" + rowIdx + "_平均" + " \\* MERGEFORMAT ", "«" + "S" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_科目(原始)" + rowIdx + "_平均" + " \\* MERGEFORMAT ", "«" + "S" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_科目" + rowIdx + "_及格人數" + " \\* MERGEFORMAT ", "«" + "P" + rowIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "班級" + "_科目(原始)" + rowIdx + "_及格人數" + " \\* MERGEFORMAT ", "«" + "P" + rowIdx + "»");
                builder.EndRow();

            }

            builder.EndTable();

            #endregion

            #region 班級領域_科目 1,2,3 平均、及格人數

            builder.Writeln();
            builder.Writeln();
            foreach (string dName in domainNameList)
            {
                builder.Writeln();
                builder.Writeln("班級 " + dName + " 領域_科目平均、及格人數");
                builder.StartTable();
                builder.InsertCell(); builder.Write("名稱");
                builder.InsertCell(); builder.Write("學分");
                builder.InsertCell(); builder.Write("平均");
                builder.InsertCell(); builder.Write("平均(原始)");
                builder.InsertCell(); builder.Write("及格人數");
                builder.InsertCell(); builder.Write("及格人數(原始)");
                builder.EndRow();

                for (int rowIdx = 1; rowIdx <= 12; rowIdx++)
                {
                    // 班級_領域_科目1_名稱                
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + rowIdx + "_名稱" + " \\* MERGEFORMAT ", "«" + "N" + rowIdx + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + rowIdx + "_學分" + " \\* MERGEFORMAT ", "«" + "C" + rowIdx + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + rowIdx + "_平均" + " \\* MERGEFORMAT ", "«" + "S" + rowIdx + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域(原始)_科目" + rowIdx + "_平均" + " \\* MERGEFORMAT ", "«" + "S" + rowIdx + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域_科目" + rowIdx + "_及格人數" + " \\* MERGEFORMAT ", "«" + "P" + rowIdx + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "班級_" + dName + "領域(原始)_科目" + rowIdx + "_及格人數" + " \\* MERGEFORMAT ", "«" + "P" + rowIdx + "»");
                    builder.EndRow();

                }
                builder.EndTable();
            }


            #endregion

            #region 學生領域相關總計成績
            builder.Writeln();
            builder.Writeln();
            builder.Writeln("學生領域相關總計成績(即時運算：加權總分、總分、加權平均、平均)");
            builder.StartTable();
            builder.InsertCell(); builder.Write("座號");
            builder.InsertCell(); builder.Write("姓名");
            builder.InsertCell(); builder.Write("加權總分");
            builder.InsertCell(); builder.Write("總分");
            builder.InsertCell(); builder.Write("加權平均");
            builder.InsertCell(); builder.Write("平均");
            builder.InsertCell(); builder.Write("加權總分(原始)");
            builder.InsertCell(); builder.Write("總分(原始)");
            builder.InsertCell(); builder.Write("加權平均(原始)");
            builder.InsertCell(); builder.Write("平均(原始)");
            builder.EndRow();

            for (int studIdx = 1; studIdx <= 50; studIdx++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");

                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域總計_加權總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域總計_總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域總計_加權平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域總計_平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)總計_加權總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)總計_總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)總計_加權平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_領域(原始)總計_平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.EndRow();
            }
            builder.EndTable();

            #endregion

            #region 學生科目相關總計成績
            builder.Writeln();
            builder.Writeln();
            builder.Writeln("學生科目相關總計成績(即時運算：加權總分、總分、加權平均、平均)");
            builder.StartTable();
            builder.InsertCell(); builder.Write("座號");
            builder.InsertCell(); builder.Write("姓名");
            builder.InsertCell(); builder.Write("加權總分");
            builder.InsertCell(); builder.Write("總分");
            builder.InsertCell(); builder.Write("加權平均");
            builder.InsertCell(); builder.Write("平均");
            builder.InsertCell(); builder.Write("加權總分(原始)");
            builder.InsertCell(); builder.Write("總分(原始)");
            builder.InsertCell(); builder.Write("加權平均(原始)");
            builder.InsertCell(); builder.Write("平均(原始)");
            builder.EndRow();

            for (int studIdx = 1; studIdx <= 50; studIdx++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");

                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目總計_加權總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目總計_總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目總計_加權平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目總計_平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)總計_加權總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)總計_總分" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)總計_加權平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "學生" + studIdx + "_科目(原始)總計_平均" + " \\* MERGEFORMAT ", "«" + "S" + studIdx + "»");
                builder.EndRow();
            }
            builder.EndTable();

            #endregion




            try
            {

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
