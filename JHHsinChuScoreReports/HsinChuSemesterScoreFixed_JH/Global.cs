using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using FISCA.Data;
using System.Data;
using Aspose.Words;
using JHSchool.Data;
using K12.Data;
using System.Xml.Linq;


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

        public const int SupportSubjectCount = 40, SupportDomainCount = 30, SupportAbsentCount = 20;

        public static Dictionary<string, string> DLBehaviorRef = new Dictionary<string, string>()
        {
            {"日常生活表現程度","DailyBehavior/Item"},
            //{"團體活動表現","GroupActivity/Item"},
            //{"公共服務表現","PublicService/Item"},
            //{"校內外特殊表現","SchoolSpecial/Item"},
            {"日常生活表現具體建議","DailyLifeRecommend"},
            {"團體活動表現","OtherRecommend"},
            //{"綜合評語","DailyLifeRecommend"}
        };

        public static string GetKey(params string[] list)
        {
            return string.Join("_", list);
        }


        /// <summary>
        /// XML 內解析子項目名稱
        /// </summary>
        /// <param name="elm"></param>
        /// <returns></returns>
        private static List<string> ParseItems(XElement elm)
        {
            List<string> retVal = new List<string>();

            foreach (XElement subElm in elm.Elements("Item"))
            {
                // 因為社團功能，所以要將"社團活動" 字不放入
                string name = subElm.Attribute("Name").Value;
                if (name != "社團活動")
                    retVal.Add(name);
            }
            return retVal;
        }


        /// <summary>
        /// 設定檔預設名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> DefaultConfigNameList()
        {
            List<string> retVal = new List<string>();
            //retVal.Add("領域成績單");
            //retVal.Add("科目成績單");
            //retVal.Add("科目及領域成績單_領域組距");
            //retVal.Add("科目及領域成績單_科目組距");
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

            #region 儲存檔案
            string inputReportName = "學期成績單合併欄位總表";
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

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.學期成績單合併欄位總表));

            try
            {
                #region 動態產生合併欄位
                // 讀取總表檔案並動態加入合併欄位
                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(tempDoc);
                builder.MoveToDocumentEnd();

                List<string> plist = K12.Data.PeriodMapping.SelectAll().Select(x => x.Type).Distinct().ToList();
                List<string> alist = K12.Data.AbsenceMapping.SelectAll().Select(x => x.Name).ToList();
                builder.Writeln();
                builder.Writeln();
                builder.Writeln("缺曠動態產生合併欄位");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("缺曠名稱與合併欄位");
                builder.EndRow();

                foreach (string pp in plist)
                {
                    foreach (string aa in alist)
                    {

                        string key = pp.Replace(" ", "_") + "_" + aa.Replace(" ", "_");

                        builder.InsertCell();
                        builder.Write(key);
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + key + " \\* MERGEFORMAT ", "«" + key + "»");
                        builder.EndRow();
                    }
                }

                builder.EndTable();

                builder.Writeln();
                builder.Writeln("缺曠總計(不分節次類型)合併欄位");
                builder.StartTable();

                foreach (string aa in alist)
                {
                    string key = aa.Replace(" ","_") + "總計";
                    builder.InsertCell();
                    builder.Write(key);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + key + " \\* MERGEFORMAT ", "«" + key + "»");
                    builder.EndRow();
                }
                builder.EndTable();


                // 日常生活表現
                builder.Writeln();
                builder.Writeln();
                builder.Writeln("日常生活表現評量");
                builder.StartTable();
                builder.InsertCell();
                builder.Write("分類");
                builder.InsertCell();
                builder.Write("名稱");
                builder.InsertCell();
                builder.Write("建議內容");
                builder.EndRow();

                foreach (string key in DLBehaviorRef.Keys)
                {
                    builder.InsertCell();
                    builder.Write(key);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + key + "_Name" + " \\* MERGEFORMAT ", "«" + key + "名稱»");
                    builder.InsertCell();
                    // 新竹版沒有
                    if (key != "日常生活表現程度")
                        builder.InsertField("MERGEFIELD " + key + "_Description" + " \\* MERGEFORMAT ", "«" + key + "建議內容»");

                    builder.EndRow();

                }
                builder.EndTable();

                // 日常生活表現
                builder.Writeln();
                builder.Writeln();
                builder.Writeln("日常生活表現評量子項目");
                builder.StartTable();
                builder.InsertCell();
                builder.Write("項目");
                builder.InsertCell();
                builder.Write("指標");
                builder.InsertCell();
                builder.Write("表現程度");
                builder.EndRow();

                for (int i = 1; i <= 7; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Name" + i + " \\* MERGEFORMAT ", "«項目" + i + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Index" + i + " \\* MERGEFORMAT ", "«指標" + i + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Degree" + i + " \\* MERGEFORMAT ", "«表現" + i + "»");
                    builder.EndRow();
                }

                builder.EndTable();





                // 動態計算領域
                List<JHSemesterScoreRecord> SemesterScoreRecordList = JHSemesterScore.SelectBySchoolYearAndSemester(_SelStudentIDList, _SelSchoolYear, _SelSemester);

                // 領域名稱
                List<string> DomainNameList = new List<string>();

                foreach (JHSemesterScoreRecord SemsScore in SemesterScoreRecordList)
                {
                    foreach (string dn in SemsScore.Domains.Keys)
                    {
                        if (!DomainNameList.Contains(dn))
                            DomainNameList.Add(dn);
                    }
                }
                DomainNameList.Sort(new StringComparer("語文", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"));
                DomainNameList.Add("彈性課程");

                List<string> m1 = new List<string>();
                List<string> m2a = new List<string>();
                List<string> m2b = new List<string>();

                m1.Add("班排名");
                m1.Add("年排名");
                m1.Add("類別1排名");
                m1.Add("類別2排名");

                m2a.Add("rank");
                m2a.Add("matrix_count");
                m2a.Add("pr");
                m2a.Add("percentile");
                m2a.Add("avg_top_25");
                m2a.Add("avg_top_50");
                m2a.Add("avg");
                m2a.Add("avg_bottom_50");
                m2a.Add("avg_bottom_25");
                m2a.Add("pr_88");
                m2a.Add("pr_75");
                m2a.Add("pr_50");
                m2a.Add("pr_25");
                m2a.Add("pr_12");
                m2a.Add("std_dev_pop");
                m2b.Add("level_gte100");
                m2b.Add("level_90");
                m2b.Add("level_80");
                m2b.Add("level_70");
                m2b.Add("level_60");
                m2b.Add("level_50");
                m2b.Add("level_40");
                m2b.Add("level_30");
                m2b.Add("level_20");
                m2b.Add("level_10");
                m2b.Add("level_lt10");


                // 領域排名 排名、母數、五標
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域成績排名 排名、母數、五標";
                    builder.Writeln(dn);
                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("名稱");
                    builder.InsertCell();
                    builder.Write("排名");
                    builder.InsertCell();
                    builder.Write("排名母數");
                    builder.InsertCell();
                    builder.Write("PR");
                    builder.InsertCell();
                    builder.Write("百分比");
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
                    builder.InsertCell();
                    builder.Write("新頂標");
                    builder.InsertCell();
                    builder.Write("新前標");
                    builder.InsertCell();
                    builder.Write("新均標");
                    builder.InsertCell();
                    builder.Write("新後標");
                    builder.InsertCell();
                    builder.Write("新底標");
                    builder.InsertCell();
                    builder.Write("標準差");
                    builder.EndRow();
                    foreach (string m in m1)
                    {
                        builder.InsertCell();
                        builder.Write(m);
                        foreach (string nn in m2a)
                        {
                            string dd = dName + "領域成績_" + m + "_" + nn;
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R»");
                        }
                        builder.EndRow();
                    }

                    builder.EndTable();
                }


                // 領域(原始)排名 排名、母數、五標
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域成績(原始)排名 排名、母數、五標";
                    builder.Writeln(dn);
                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("名稱");
                    builder.InsertCell();
                    builder.Write("排名");
                    builder.InsertCell();
                    builder.Write("排名母數");
                    builder.InsertCell();
                    builder.Write("PR");
                    builder.InsertCell();
                    builder.Write("百分比");
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
                    builder.InsertCell();
                    builder.Write("新頂標");
                    builder.InsertCell();
                    builder.Write("新前標");
                    builder.InsertCell();
                    builder.Write("新均標");
                    builder.InsertCell();
                    builder.Write("新後標");
                    builder.InsertCell();
                    builder.Write("新底標");
                    builder.InsertCell();
                    builder.Write("標準差");
                    builder.EndRow();
                    foreach (string m in m1)
                    {
                        builder.InsertCell();
                        builder.Write(m);
                        foreach (string nn in m2a)
                        {
                            string dd = dName + "領域成績(原始)_" + m + "_" + nn;
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R»");
                        }
                        builder.EndRow();
                    }

                    builder.EndTable();
                }


                // 領域排名 組距
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域成績排名 組距";
                    builder.Writeln(dn);
                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("名稱");
                    builder.InsertCell();
                    builder.Write("100以上");
                    builder.InsertCell();
                    builder.Write("90以上小於100");
                    builder.InsertCell();
                    builder.Write("80以上小於90");
                    builder.InsertCell();
                    builder.Write("70以上小於80");
                    builder.InsertCell();
                    builder.Write("60以上小於70");
                    builder.InsertCell();
                    builder.Write("50以上小於60");
                    builder.InsertCell();
                    builder.Write("40以上小於50");
                    builder.InsertCell();
                    builder.Write("30以上小於40");
                    builder.InsertCell();
                    builder.Write("20以上小於30");
                    builder.InsertCell();
                    builder.Write("10以上小於20");
                    builder.InsertCell();
                    builder.Write("10以下");
                    builder.EndRow();
                    foreach (string m in m1)
                    {
                        builder.InsertCell();
                        builder.Write(m);
                        foreach (string nn in m2b)
                        {
                            string dd = dName + "領域成績_" + m + "_" + nn;
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R»");
                        }
                        builder.EndRow();
                    }

                    builder.EndTable();
                }

                // 領域排名(原始) 組距
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域成績(原始)排名 組距";
                    builder.Writeln(dn);
                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("名稱");
                    builder.InsertCell();
                    builder.Write("100以上");
                    builder.InsertCell();
                    builder.Write("90以上小於100");
                    builder.InsertCell();
                    builder.Write("80以上小於90");
                    builder.InsertCell();
                    builder.Write("70以上小於80");
                    builder.InsertCell();
                    builder.Write("60以上小於70");
                    builder.InsertCell();
                    builder.Write("50以上小於60");
                    builder.InsertCell();
                    builder.Write("40以上小於50");
                    builder.InsertCell();
                    builder.Write("30以上小於40");
                    builder.InsertCell();
                    builder.Write("20以上小於30");
                    builder.InsertCell();
                    builder.Write("10以上小於20");
                    builder.InsertCell();
                    builder.Write("10以下");
                    builder.EndRow();
                    foreach (string m in m1)
                    {
                        builder.InsertCell();
                        builder.Write(m);
                        foreach (string nn in m2b)
                        {
                            string dd = dName + "領域成績(原始)_" + m + "_" + nn;
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R»");
                        }
                        builder.EndRow();
                    }

                    builder.EndTable();
                }

                // 領域-科目排名 排名、母數、五標
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域-科目成績排名 排名、母數、五標";
                    builder.Writeln(dn);
                    builder.StartTable();
                    foreach (string m in m1)
                    {
                        builder.Writeln(m);
                        builder.InsertCell();
                        builder.Write("名稱");
                        builder.InsertCell();
                        builder.Write("排名");
                        builder.InsertCell();
                        builder.Write("排名母數");
                        builder.InsertCell();
                        builder.Write("PR");
                        builder.InsertCell();
                        builder.Write("百分比");
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
                        builder.InsertCell();
                        builder.Write("新頂標");
                        builder.InsertCell();
                        builder.Write("新前標");
                        builder.InsertCell();
                        builder.Write("新均標");
                        builder.InsertCell();
                        builder.Write("新後標");
                        builder.InsertCell();
                        builder.Write("新底標");
                        builder.InsertCell();
                        builder.Write("標準差");
                        builder.EndRow();

                        for (int i = 1; i <= 12; i++)
                        {
                            builder.InsertCell();
                            string dsn = dName + "_科目排名名稱" + i;
                            builder.InsertField("MERGEFIELD " + dsn + " \\* MERGEFORMAT ", "«N" + i + "»");
                            foreach (string nn in m2a)
                            {
                                string dd = dName + "_科目成績" + i + "_" + m + "_" + nn;
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R" + i + "»");
                            }
                            builder.EndRow();
                        }

                    }

                    builder.EndTable();
                }

                // 領域-科目排名 排名、母數、五標
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域-科目成績(原始)排名 排名、母數、五標";
                    builder.Writeln(dn);
                    builder.StartTable();
                    foreach (string m in m1)
                    {
                        builder.Writeln(m);
                        builder.InsertCell();
                        builder.Write("名稱");
                        builder.InsertCell();
                        builder.Write("排名");
                        builder.InsertCell();
                        builder.Write("排名母數");
                        builder.InsertCell();
                        builder.Write("PR");
                        builder.InsertCell();
                        builder.Write("百分比");
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
                        builder.InsertCell();
                        builder.Write("新頂標");
                        builder.InsertCell();
                        builder.Write("新前標");
                        builder.InsertCell();
                        builder.Write("新均標");
                        builder.InsertCell();
                        builder.Write("新後標");
                        builder.InsertCell();
                        builder.Write("新底標");
                        builder.InsertCell();
                        builder.Write("標準差");
                        builder.EndRow();

                        for (int i = 1; i <= 12; i++)
                        {
                            builder.InsertCell();
                            string dsn = dName + "_科目排名名稱" + i;
                            builder.InsertField("MERGEFIELD " + dsn + " \\* MERGEFORMAT ", "«N" + i + "»");
                            foreach (string nn in m2a)
                            {
                                string dd = dName + "_科目成績(原始)" + i + "_" + m + "_" + nn;
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R" + i + "»");
                            }
                            builder.EndRow();
                        }

                    }

                    builder.EndTable();
                }

                // 領域-科目排名 組距
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域-科目成績排名 組距";
                    builder.Writeln(dn);
                    builder.StartTable();
                    foreach (string m in m1)
                    {
                        builder.Writeln(m);
                        builder.InsertCell();
                        builder.Write("名稱");
                        builder.InsertCell();
                        builder.Write("100以上");
                        builder.InsertCell();
                        builder.Write("90以上小於100");
                        builder.InsertCell();
                        builder.Write("80以上小於90");
                        builder.InsertCell();
                        builder.Write("70以上小於80");
                        builder.InsertCell();
                        builder.Write("60以上小於70");
                        builder.InsertCell();
                        builder.Write("50以上小於60");
                        builder.InsertCell();
                        builder.Write("40以上小於50");
                        builder.InsertCell();
                        builder.Write("30以上小於40");
                        builder.InsertCell();
                        builder.Write("20以上小於30");
                        builder.InsertCell();
                        builder.Write("10以上小於20");
                        builder.InsertCell();
                        builder.Write("10以下");
                        builder.EndRow();

                        for (int i = 1; i <= 12; i++)
                        {
                            builder.InsertCell();
                            string dsn = dName + "_科目排名名稱" + i;
                            builder.InsertField("MERGEFIELD " + dsn + " \\* MERGEFORMAT ", "«N" + i + "»");
                            foreach (string nn in m2b)
                            {
                                string dd = dName + "_科目成績" + i + "_" + m + "_" + nn;
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R" + i + "»");
                            }
                            builder.EndRow();
                        }


                    }

                    builder.EndTable();
                }

                // 領域-科目(原始)排名 組距
                foreach (string dName in DomainNameList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    string dn = dName + "領域-科目成績(原始)排名 組距";
                    builder.Writeln(dn);
                    builder.StartTable();
                    foreach (string m in m1)
                    {
                        builder.Writeln(m);
                        builder.InsertCell();
                        builder.Write("名稱");
                        builder.InsertCell();
                        builder.Write("100以上");
                        builder.InsertCell();
                        builder.Write("90以上小於100");
                        builder.InsertCell();
                        builder.Write("80以上小於90");
                        builder.InsertCell();
                        builder.Write("70以上小於80");
                        builder.InsertCell();
                        builder.Write("60以上小於70");
                        builder.InsertCell();
                        builder.Write("50以上小於60");
                        builder.InsertCell();
                        builder.Write("40以上小於50");
                        builder.InsertCell();
                        builder.Write("30以上小於40");
                        builder.InsertCell();
                        builder.Write("20以上小於30");
                        builder.InsertCell();
                        builder.Write("10以上小於20");
                        builder.InsertCell();
                        builder.Write("10以下");
                        builder.EndRow();

                        for (int i = 1; i <= 12; i++)
                        {
                            builder.InsertCell();
                            string dsn = dName + "_科目排名名稱" + i;
                            builder.InsertField("MERGEFIELD " + dsn + " \\* MERGEFORMAT ", "«N" + i + "»");
                            foreach (string nn in m2b)
                            {
                                string dd = dName + "_科目成績(原始)" + i + "_" + m + "_" + nn;
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + dd + " \\* MERGEFORMAT ", "«R" + i + "»");
                            }
                            builder.EndRow();
                        }


                    }

                    builder.EndTable();
                }

                #endregion
                tempDoc.Save(path, SaveFormat.Doc);

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
