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
            {"日常行為表現","DailyBehavior/Item"},
            {"團體活動表現","GroupActivity/Item"},
            {"公共服務表現","PublicService/Item"},
            {"校內外特殊表現","SchoolSpecial/Item"},
            {"具體建議","DailyLifeRecommend"},
            {"其他表現","OtherRecommend"},
            {"綜合評語","DailyLifeRecommend"}
        };

        public static string GetKey(params string[] list)
        {
            return string.Join("_", list);
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
