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
