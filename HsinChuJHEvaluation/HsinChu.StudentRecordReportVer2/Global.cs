using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

namespace HsinChu.StudentRecordReportVer2
{
    public class Global
    {
        /// <summary>
        /// 匯出合併欄位總表Word
        /// </summary>
        public static void ExportMappingFieldWord()
        {
            #region 儲存檔案
            string inputReportName = "國中學籍表(十二年國教)合併欄位總表";
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

            Document tempDoc = new Document(new MemoryStream(Resources.新竹國中學籍表功能變數));


            try
            {
                #region 動態產生合併欄位
                // 讀取總表檔案並動態加入合併欄位
                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(tempDoc);
                builder.MoveToDocumentEnd();

                #region 日常生活表現
                for (int a = 1; a <= 6; a++)
                {
                    builder.Writeln();
                    builder.Writeln("日常生活表現評量子項目_第" + a + "學期");
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
                        builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Name" + i + "_" + a + " \\* MERGEFORMAT ", "«項目" + i + "»");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Index" + i + "_" + a + " \\* MERGEFORMAT ", "«指標" + i + "»");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Degree" + i + "_" + a + " \\* MERGEFORMAT ", "«表現" + i + "»");
                        builder.EndRow();
                    }
                    builder.EndTable();
                }
                #endregion

                builder.Writeln();
                builder.Font.Size = 20;
                builder.Font.Bold = true;
                builder.Writeln("畢業相關");

                builder.Font.Size = 12;
                builder.Font.Bold = false;
                builder.InsertField("MERGEFIELD " + "畢業總成績_平均" + " \\* MERGEFORMAT ", "«畢業總成績_平均»");
                builder.Writeln();
                builder.InsertField("MERGEFIELD " + "畢業總成績_等第" + " \\* MERGEFORMAT ", "«畢業總成績_等第»");
                builder.Writeln();
                builder.InsertField("MERGEFIELD " + "准予畢業" + " \\* MERGEFORMAT ", "«准予畢業»");
                builder.Writeln();
                builder.InsertField("MERGEFIELD " + "發給修業證書" + " \\* MERGEFORMAT ", "«發給修業證書»");

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


        public static Dictionary<string, string> DLBehaviorRef = new Dictionary<string, string>()
        {
            {"日常生活表現程度","DailyBehavior/Item"},
            {"日常生活表現具體建議","DailyLifeRecommend"},
            {"團體活動表現","OtherRecommend"},
        };


    }
}
