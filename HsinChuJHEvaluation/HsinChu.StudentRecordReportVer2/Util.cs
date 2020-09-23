using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using K12.Data;
using Aspose.Words;
using Aspose.Words.Tables;
using FISCA.Presentation.Controls;
using Campus.Rating;
using System.Globalization;
using FISCA.Data;
using System.Data;
using System.IO;

namespace HsinChu.StudentRecordReportVer2
{
    internal static class Util
    {
        /// <summary>
        /// 英文日期格式。
        /// </summary>
        public const string EnglishFormat = "MMMM dd, yyyy";

        public static CultureInfo USCulture = new CultureInfo("en-us");

        public static void DisableControls(Control topControl)
        {
            ChangeControlsStatus(topControl, false);
        }

        public static void EnableControls(Control topControl)
        {
            ChangeControlsStatus(topControl, true);
        }

        private static void ChangeControlsStatus(Control topControl, bool status)
        {
            foreach (Control each in topControl.Controls)
            {
                string tag = each.Tag + "";
                if (tag.ToUpper() == "StatusVarying".ToUpper())
                {
                    each.Enabled = status;
                }

                if (each.Controls.Count > 0)
                    ChangeControlsStatus(each, status);
            }
        }

        public static void Save(Document doc, string fileName, bool convertToPDF)
        {
            try
            {
                if (doc != null)
                {
                    string path = "";
                    if (convertToPDF)
                    {
                        path = $"{Application.StartupPath}\\Reports\\{fileName}.pdf";
                    }
                    else
                    {
                        path = $"{Application.StartupPath}\\Reports\\{fileName}.docx";
                    }

                    int i = 1;
                    while (File.Exists(path))
                    {
                        string newPath = $"{Path.GetDirectoryName(path)}\\{fileName}{i++}{Path.GetExtension(path)}";
                        path = newPath;
                    }

                    doc.Save(path, convertToPDF ? SaveFormat.Pdf : SaveFormat.Docx);

                    DialogResult dialogResult = MessageBox.Show($"{path}\n{fileName}產生完成，是否立即開啟？", "訊息", MessageBoxButtons.YesNo);

                    if (DialogResult.Yes == dialogResult)
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存失敗。" + ex.Message);
                return;
            }
        }

        public static string GetGradeyearString(string gradeYear)
        {
            switch (gradeYear)
            {
                case "1":
                    return "一";
                case "2":
                    return "二";
                case "3":
                    return "三";
                case "4":
                    return "四";
                case "5":
                    return "五";
                case "6":
                    return "六";
                case "7":
                    return "七";
                case "8":
                    return "八";
                case "9":
                    return "九";
                case "10":
                    return "十";
                case "11":
                    return "十一";
                case "12":
                    return "十二";
                default:
                    return gradeYear;
            }
        }

        /// <summary>
        /// 取得下一個 Cell 的 Paragraph。
        /// </summary>
        public static Paragraph NextCell(Paragraph para)
        {
            if (para.ParentNode is Cell)
            {
                Cell cell = para.ParentNode.NextSibling as Cell;

                if (cell == null) return null;

                if (cell.Paragraphs.Count <= 0)
                    cell.Paragraphs.Add(new Paragraph(para.Document));

                return cell.Paragraphs[0];
            }
            else
                return null;
        }

        /// <summary>
        /// 取得前一個 Cell 的 Paragraph。
        /// </summary>
        /// <param name="para"></param>
        /// <returns></returns>
        public static Paragraph PreviousCell(Paragraph para)
        {
            if (para.ParentNode is Cell)
            {
                Cell cell = para.ParentNode.PreviousSibling as Cell;

                if (cell == null) return null;

                if (cell.Paragraphs.Count <= 0)
                    cell.Paragraphs.Add(new Paragraph(para.Document));

                return cell.Paragraphs[0];
            }
            else
                return null;
        }

        public static void Write(this Cell cell, DocumentBuilder builder, string text)
        {
            if (cell.Paragraphs.Count <= 0)
                cell.Paragraphs.Add(new Paragraph(cell.Document));

            builder.MoveTo(cell.Paragraphs[0]);
            builder.Write(text);
        }

        public static string GetDegree(decimal score)
        {
            if (score >= 90) return "優";
            else if (score >= 80) return "甲";
            else if (score >= 70) return "乙";
            else if (score >= 60) return "丙";
            else return "丁";
        }

        public static string GetDegreeEnglish(decimal score)
        {
            if (score >= 90) return "A";
            else if (score >= 80) return "B";
            else if (score >= 70) return "C";
            else if (score >= 60) return "D";
            else return "E";
        }

        /// <summary>
        ///  例：一般:曠課,事假,病假;集合:曠課,事假,公假
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public static string PeriodOptionsToString(this Dictionary<string, List<string>> setting)
        {
            StringBuilder builder = new StringBuilder();
            foreach (KeyValuePair<string, List<string>> eachType in setting)
            {
                builder.Append(eachType.Key + ":");

                foreach (string each in eachType.Value)
                    builder.Append(each + ",");

                builder.Append(";");
            }
            return builder.ToString();
        }

        /// <summary>
        /// 例：一般:曠課,事假,病假;集合:曠課,事假,公假
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public static Dictionary<string, List<string>> PeriodOptionsFromString(this string setting)
        {
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>();
            //以「;」分割每一個節次類別。
            foreach (string eachType in setting.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                //以「:」分割類別名稱與資料。
                string[] arrTypeData = eachType.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                string typeName, typeData;
                if (arrTypeData.Length >= 2)
                {
                    typeName = arrTypeData[0];
                    typeData = arrTypeData[1];
                }
                else
                    continue;

                result.Add(typeName, new List<string>());
                //以「,」分割每個資料項。
                foreach (string eachEntry in typeData.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    result[typeName].Add(eachEntry);
            }
            return result;
        }

        /// <summary>
        /// 透過學生編號,取得特定學年度學期服務學習時數
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// /// <returns></returns>
        public static Dictionary<string, Dictionary<string, string>> GetServiceLearningDetail(List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, string>> retVal = new Dictionary<string, Dictionary<string, string>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,school_year,semester,sum(hours) as hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') group by ref_student_id,school_year,semester order by school_year,semester;";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["ref_student_id"].ToString();
                    string key1 = dr["school_year"].ToString() + "_" + dr["semester"].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, new Dictionary<string, string>());

                    if (!retVal[sid].ContainsKey(key1))
                        retVal[sid].Add(key1, "0");

                    retVal[sid][key1] = dr["hours"].ToString();
                }
            }
            return retVal;
        }

        /// <summary>
        /// 服務學時數暫存使用
        /// </summary>
        public static Dictionary<string, Dictionary<string, string>> _SLRDict = new Dictionary<string, Dictionary<string, string>>();
    }
}
