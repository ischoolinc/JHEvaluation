﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using JHEvaluation.ScoreCalculation;
using JHSchool.Data;
using Aspose.Words;
using FISCA.Presentation.Controls;
using Campus.Rating;
using System.Globalization;
using K12.Data;
using FISCA.UDT;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;

namespace KH_StudentScoreSummaryReport
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

        /// <summary>
        /// 將學生編號轉換成 SCStudent 物件。
        /// </summary>
        /// <remarks>使用指定的學生編號，向 DAL 取得 VO 後轉換成 SCStudent 物件。</remarks>
        public static List<ReportStudent> ToReportStudent(this IEnumerable<string> studentIDs)
        {
            List<string> StudentIDList = new List<string>();
            List<ReportStudent> students = new List<ReportStudent>();
            foreach (JHStudentRecord each in JHStudent.SelectByIDs(studentIDs))
            {
                StudentIDList.Add(each.ID);
                students.Add(new ReportStudent(each));
            }
            // 取得學生類別List
            List<StudentTagRecord> StudTagRecList = StudentTag.SelectByStudentIDs(StudentIDList);

            // 取得特種身分學生,加分比
            Dictionary<string, decimal> StudAddWeightDict = DALTransfer.GetStudentAddWeightFormUDTByStudentTag(StudTagRecList, DALTransfer.SchoolType.高中);
            // 取得特種身分學生名稱
            Dictionary<string, string> StudSpecTypeDict = DALTransfer.GetStudentSpcTypeFormUDTByStudentTag(StudTagRecList, DALTransfer.SchoolType.高中);
            Dictionary<string, K12.Data.UpdateRecordRecord> studUpdateRec = DALTransfer.GetStudentUpdareDate3ByStudentID(StudentIDList);

            foreach (ReportStudent rs in students)
            {
                // 加分比
                if (StudAddWeightDict.ContainsKey(rs.StudentID))
                    rs.AddWeight = StudAddWeightDict[rs.StudentID];

                // 學生身分
                if (StudSpecTypeDict.ContainsKey(rs.StudentID))
                    rs.SpcStudTypeName = StudSpecTypeDict[rs.StudentID];

                // 轉入學生異動
                if (studUpdateRec.ContainsKey(rs.StudentID))
                {
                    if (studUpdateRec[rs.StudentID] != null)
                    {
                        DateTime dt;

                        if (DateTime.TryParse(studUpdateRec[rs.StudentID].UpdateDate, out dt))
                        {
                            rs.TransUpdateDateStr = (dt.Year - 1911).ToString() + "/" + dt.Month + "/" + dt.Day;
                        }
                        else
                            rs.TransUpdateDateStr = "";

                        rs.LastEnterSchoolyear = studUpdateRec[rs.StudentID].SchoolYear;
                        rs.LastEnterSemester = studUpdateRec[rs.StudentID].Semester;
                        int gr;
                        if (int.TryParse(studUpdateRec[rs.StudentID].GradeYear, out gr))
                            rs.LastEnterGradeYear = gr;
                    }
                }

            }

            return students;
        }

        /// <summary>
        /// 取得全部學生的資料(只包含一般、輟學)。
        /// </summary>
        /// <returns></returns>
        public static List<ReportStudent> GetAllStudents()
        {
            List<ReportStudent> students = new List<ReportStudent>();
            foreach (JHStudentRecord each in JHStudent.SelectAll())
            {
                if (each.Status == K12.Data.StudentRecord.StudentStatus.一般 ||
                    each.Status == K12.Data.StudentRecord.StudentStatus.輟學)
                    students.Add(new ReportStudent(each));
            }
            return students;
        }

        /// <summary>
        /// 傳入不包含學生的學生IDList，取得一般或輟學學生
        /// </summary>
        /// <param name="NotInStudentIDList"></param>
        /// <returns></returns>
        public static List<ReportStudent> GetStudentsDef(List<string> NotInStudentIDList)
        {
            List<string> StudentIDList = new List<string>();

            List<ReportStudent> students = new List<ReportStudent>();
            foreach (JHStudentRecord each in JHStudent.SelectAll())
            {
                // 不包含
                if (NotInStudentIDList.Contains(each.ID))
                    continue;

                if (each.Status == K12.Data.StudentRecord.StudentStatus.一般 ||
                    each.Status == K12.Data.StudentRecord.StudentStatus.輟學)
                {
                    ReportStudent rs = new ReportStudent(each);
                    students.Add(rs);
                    StudentIDList.Add(each.ID);
                }
            }

            // 取得學生類別List
            List<StudentTagRecord> StudTagRecList = StudentTag.SelectByStudentIDs(StudentIDList);


            // 取得特種身分學生,加分比
            Dictionary<string, decimal> StudAddWeightDict = DALTransfer.GetStudentAddWeightFormUDTByStudentTag(StudTagRecList, DALTransfer.SchoolType.高中);
            // 取得特種身分學生名稱
            Dictionary<string, string> StudSpecTypeDict = DALTransfer.GetStudentSpcTypeFormUDTByStudentTag(StudTagRecList, DALTransfer.SchoolType.高中);
            Dictionary<string, K12.Data.UpdateRecordRecord> studUpdateRec = DALTransfer.GetStudentUpdareDate3ByStudentID(StudentIDList);

            foreach (ReportStudent rs in students)
            {
                // 加分比
                if (StudAddWeightDict.ContainsKey(rs.StudentID))
                    rs.AddWeight = StudAddWeightDict[rs.StudentID];

                // 學生身分
                if (StudSpecTypeDict.ContainsKey(rs.StudentID))
                    rs.SpcStudTypeName = StudSpecTypeDict[rs.StudentID];

                // 轉入學生異動
                if (studUpdateRec.ContainsKey(rs.StudentID))
                {
                    if (studUpdateRec[rs.StudentID] != null)
                    {
                        DateTime dt;

                        if (DateTime.TryParse(studUpdateRec[rs.StudentID].UpdateDate, out dt))
                        {
                            rs.TransUpdateDateStr = (dt.Year - 1911).ToString() + "/" + dt.Month + "/" + dt.Day;
                        }
                        else
                            rs.TransUpdateDateStr = "";

                        rs.LastEnterSchoolyear = studUpdateRec[rs.StudentID].SchoolYear;
                        rs.LastEnterSemester = studUpdateRec[rs.StudentID].Semester;
                        int gr;
                        if (int.TryParse(studUpdateRec[rs.StudentID].GradeYear, out gr))
                            rs.LastEnterGradeYear = gr;
                    }
                }

            }
            return students;
        }

        /// <summary>
        /// 轉型成 StudentScore 集合。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static List<StudentScore> ToSC(this IEnumerable<ReportStudent> students)
        {
            List<StudentScore> stus = new List<StudentScore>();
            foreach (ReportStudent each in students)
                stus.Add(each);
            return stus;
        }

        /// <summary>
        /// 轉型成 StudentScore。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static List<ReportStudent> ToSS(this IEnumerable<StudentScore> students)
        {
            List<ReportStudent> stus = new List<ReportStudent>();
            foreach (StudentScore each in students)
                stus.Add(each as ReportStudent);
            return stus;
        }

        /// <summary>
        /// 將指定的 SCStudent 集合轉換成 ID->SCStudent 對照。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static Dictionary<string, StudentScore> ToDictionary(this IEnumerable<StudentScore> students)
        {
            Dictionary<string, StudentScore> dicstuds = new Dictionary<string, StudentScore>();
            foreach (StudentScore each in students)
                dicstuds.Add(each.Id, each);
            return dicstuds;
        }

        public static Dictionary<string, ReportStudent> ToDictionary(this IEnumerable<ReportStudent> students)
        {
            Dictionary<string, ReportStudent> dicstuds = new Dictionary<string, ReportStudent>();
            foreach (ReportStudent each in students)
                dicstuds.Add(each.Id, each);
            return dicstuds;
        }

        /// <summary>
        /// 將 SCStudent 集合轉換成編號的集合。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static List<string> ToKeys(this IEnumerable<StudentScore> students)
        {
            List<string> keys = new List<string>();
            foreach (StudentScore each in students)
                keys.Add(each.Id);
            return keys;
        }

        public static void Save(Document doc, string fileName, bool convertToPDF)
        {
            //SaveFileDialog sdf = new SaveFileDialog();

            //if (convertToPDF)
            //{
            //    sdf.Filter = "PDF 檔案(*.pdf)|*.pdf";
            //    sdf.FileName = fileName + ".pdf";
            //}
            //else
            //{
            //    sdf.Filter = "Word 檔案(*.doc)|*.doc";
            //    sdf.FileName = fileName + ".doc";
            //}

            //if (sdf.ShowDialog() == DialogResult.OK)
            //{
            try
            {
                //doc.Save(fileName, SaveFormat.Doc);
                Campus.Report.ReportSaver.SaveDocument(doc, fileName,
                    convertToPDF ? Campus.Report.ReportSaver.OutputType.PDF : Campus.Report.ReportSaver.OutputType.Word);
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存失敗。" + ex.Message);
                return;
            }

            //try
            //{
            //    //if (MsgBox.Show("排名完成，是否立刻開啟？", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //    //{
            //    //System.Diagnostics.Process.Start(sdf.FileName);
            //    //}
            //}
            //catch (Exception ex)
            //{
            //    MsgBox.Show("開啟失敗。" + ex.Message);
            //}
            //}
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

        // 放入年排名的學生
        public static List<RatingScope<ReportStudent>> ToGradeYearScopes(this IEnumerable<ReportStudent> students)
        {
            Dictionary<string, RatingScope<ReportStudent>> scopes = new Dictionary<string, RatingScope<ReportStudent>>();

            foreach (ReportStudent each in students)
            {
                string gradeYear = string.Empty;

                if (!string.IsNullOrEmpty(each.RefClassID))
                {
                    int? gy = JHClass.SelectByID(each.RefClassID).GradeYear;
                    if (gy.HasValue) gradeYear = gy.Value.ToString();
                }

                if (!scopes.ContainsKey(gradeYear))
                    scopes.Add(gradeYear, new RatingScope<ReportStudent>(gradeYear, "年排名"));

                scopes[gradeYear].Add(each);
            }

            return new List<RatingScope<ReportStudent>>(scopes.Values);
        }

        //private static List<string> subjOrder = new List<string>(new string[] { "國語文", "語文", "國文", "英文", "英語", "數學", "社會", "歷史", "公民", "地理", "藝術與人文", "自然與生活科技", "理化", "生物", "健康與體育", "綜合活動", "學習領域", "彈性課程" });
        //public static int SortSubject(RowHeader x, RowHeader y)
        //{
        //    int ix = subjOrder.IndexOf(x.Subject);
        //    int iy = subjOrder.IndexOf(y.Subject);

        //    if (ix >= 0 && iy >= 0) //如果都有找到位置。
        //        return ix.CompareTo(iy);
        //    else if (ix >= 0)
        //        return -1;
        //    else if (iy >= 0)
        //        return 1;
        //    else
        //        return x.Subject.CompareTo(y.Subject);
        //}

        //public static int SortDomain(RowHeader x, RowHeader y)
        //{
        //    int ix = subjOrder.IndexOf(x.Domain);
        //    int iy = subjOrder.IndexOf(y.Domain);

        //    if (ix >= 0 && iy >= 0) //如果都有找到位置。
        //        return ix.CompareTo(iy);
        //    else if (ix >= 0)
        //        return -1;
        //    else if (iy >= 0)
        //        return 1;
        //    else
        //        return x.Domain.CompareTo(y.Domain);
        //}

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
        /// 取得學生服務學習時數
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<SLRecord>> GetStudentSLRecordDictByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, List<SLRecord>> retVal = new Dictionary<string, List<SLRecord>>();
            if (StudentIDList.Count > 0)
            {
                AccessHelper accessHelper = new AccessHelper();
                string query = "ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "')";
                List<SLRecord> SLRecordList = accessHelper.Select<SLRecord>(query);
                foreach (SLRecord rec in SLRecordList)
                {
                    if (!retVal.ContainsKey(rec.RefStudentID))
                        retVal.Add(rec.RefStudentID, new List<SLRecord>());

                    retVal[rec.RefStudentID].Add(rec);
                }
            }
            return retVal;
        }

        /// <summary>
        /// 取得學生擔任幹部次數 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<itemCount>> GetStudentCount1ByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, List<itemCount>> retVal = new Dictionary<string, List<itemCount>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,school_year,semester,count(id) as count from discipline where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and reason like '%[幹部]%' group by ref_student_id,school_year,semester";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string key = dr["ref_student_id"].ToString();

                    if (!retVal.ContainsKey(key))
                        retVal.Add(key, new List<itemCount>());

                    itemCount ic = new itemCount();
                    ic.SchoolYear = int.Parse(dr["school_year"].ToString());
                    ic.Semester = int.Parse(dr["semester"].ToString());
                    ic.Count = int.Parse(dr["count"].ToString());

                    retVal[key].Add(ic);
                }
            }
            return retVal;
        }

        /// <summary>
        /// 取得獎勵紀錄
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<JHMeritRecord>> GetStudentMeritRecordDict(List<string> StudentIDList)
        {
            Dictionary<string, List<JHMeritRecord>> retVal = new Dictionary<string, List<JHMeritRecord>>();
            if (StudentIDList.Count > 0)
            {
                foreach (JHMeritRecord rec in JHMerit.SelectByStudentIDs(StudentIDList))
                {
                    if (!retVal.ContainsKey(rec.RefStudentID))
                        retVal.Add(rec.RefStudentID, new List<JHMeritRecord>());

                    retVal[rec.RefStudentID].Add(rec);
                }
            }
            return retVal;
        }

        /// <summary>
        /// 透過學生ID取得學生體適能資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<StudentFitnessRecord>> GetStudentFitnessRecordDictByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, List<StudentFitnessRecord>> retVal = new Dictionary<string, List<StudentFitnessRecord>>();
            if (StudentIDList.Count > 0)
            {
                AccessHelper _accessHelper = new AccessHelper();
                string query = "ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "')";
                List<StudentFitnessRecord> FitnessRecordList = _accessHelper.Select<StudentFitnessRecord>(query);
                foreach (StudentFitnessRecord rec in FitnessRecordList)
                {
                    if (!retVal.ContainsKey(rec.StudentID))
                        retVal.Add(rec.StudentID, new List<StudentFitnessRecord>());

                    retVal[rec.StudentID].Add(rec);
                }
            }
            return retVal;


        }

        /// <summary>
        /// 將特殊字用空白表示
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string SurrogatePairString(string input)
        {
            string value = "";
            int idx = 0;
            foreach (char c in input)
            {
                if (char.IsSurrogatePair(input, idx) || char.IsSurrogate(c))
                {
                    value += " ";
                }
                else
                {
                    value += c;
                }
                idx++;
            }
            return value;
        }

    }
}
