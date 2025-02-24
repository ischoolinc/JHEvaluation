﻿using System;
using System.Collections.Generic;
using System.Text;
using JHSchool.Data;
using DevComponents.DotNetBar.Controls;
using K12.Data;
using FISCA.Presentation.Controls;
using DevComponents.Editors;
using Aspose.Cells;
using System.Windows.Forms;
using FISCA.Data;
using System.Data;


namespace JHEvaluation.ScoreCalculation
{
    internal static class Util
    {
        /// <summary>
        /// 最大的執行緒數量。
        /// </summary>
        public const int MaxThread = 5;

        public static void SetSemesterDefaultItems(IntegerInput intSchoolYear, IntegerInput intSemester)
        {
            try
            {
                int schoolyear = int.Parse(School.DefaultSchoolYear);
                int semester = int.Parse(School.DefaultSemester);

                intSchoolYear.MaxValue = schoolyear + 20;
                intSchoolYear.MinValue = schoolyear - 20;
                intSchoolYear.Value = schoolyear;

                intSemester.MaxValue = 2;
                intSemester.MinValue = 1;
                intSemester.Value = semester;
            }
            catch (Exception)
            {
                MsgBox.Show("讀取學期資訊錯誤，請確定網路連線正常後再試一次。");
            }
        }

        public static List<string> GetGradeyearStudents(int gradeYear)
        {
            List<JHStudentRecord> allstudent = JHStudent.SelectAll();

            List<string> gyStudent = new List<string>();

            foreach (JHStudentRecord each in allstudent)
            {
                if (each.Status != StudentRecord.StudentStatus.一般) continue;

                JHClassRecord cr = each.Class;

                if (cr == null) continue;
                if (!cr.GradeYear.HasValue) continue;
                if (cr.GradeYear.Value != gradeYear) continue;

                gyStudent.Add(each.ID);
            }

            return gyStudent;
        }

        /// <summary>
        /// 將學生編號轉換成 SCStudent 物件。
        /// </summary>
        /// <remarks>使用指定的學生編號，向 DAL 取得 VO 後轉換成 SCStudent 物件。</remarks>
        public static List<StudentScore> ToStudentScore(this IEnumerable<string> studentIDs)
        {
            List<StudentScore> students = new List<StudentScore>();
            foreach (JHStudentRecord each in JHStudent.SelectByIDs(studentIDs))
                students.Add(new StudentScore(each));
            return students;
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

        public static decimal Round(decimal value, int decimals)
        {
            return Math.Round(value, decimals);
        }

        public static int CalculatePercentage(int totalCount, int currentCount)
        {
            decimal dTotalCount = totalCount;
            decimal dCurrentCount = currentCount;

            if (dCurrentCount <= 0) return 100;

            return (int)Math.Round((dCurrentCount / dTotalCount) * 100m, 0m);
        }

        public static List<string> SortSubjectDomain(IEnumerable<string> names)
        {
            List<string> list = new List<string>(new string[] { "國語文", "語文", "國文", "英文", "英語", "語文", "數學", "歷史", "公民", "地理", "社會", "藝術與人文", "理化", "生物", "自然與生活科技", "健康與體育", "綜合活動" });

            List<string> orderName = new List<string>(names);
            orderName.Sort(delegate (string x, string y)
            {
                int ix = list.IndexOf(x);
                int iy = list.IndexOf(y);

                if (ix >= 0 && iy >= 0) //如果都有找到位置。
                    return ix.CompareTo(iy);
                else if (ix >= 0)
                    return -1;
                else if (iy >= 0)
                    return 1;
                else
                    return x.CompareTo(y);
            });

            return orderName;
        }

        public static void Save(Workbook book, string fileName)
        {
            SaveFileDialog sdf = new SaveFileDialog();
            sdf.FileName = fileName;
            sdf.Filter = "Excel檔案(*.xls)|*.xls";

            if (sdf.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    book.Save(sdf.FileName, FileFormatType.Excel2003);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("儲存失敗。" + ex.Message);
                    return;
                }

                try
                {
                    //if (MsgBox.Show("排名完成，是否立刻開啟？", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    //{
                    System.Diagnostics.Process.Start(sdf.FileName);
                    //}
                }
                catch (Exception ex)
                {
                    MsgBox.Show("開啟失敗。" + ex.Message);
                }
            }
        }

        #region 彈性課程處理
        public static UniqueSet<string> VariableDomains = new UniqueSet<string>(new string[] { "__彈性課程", "彈性課程", "" });

        /// <summary>
        /// 判斷是否屬於彈性課程的領域。
        /// </summary>
        /// <param name="domainName"></param>
        /// <returns></returns>
        public static bool IsVariableDomain(string domainName)
        {
            return VariableDomains.Contains(domainName);
        }

        /// <summary>
        /// 將領域名稱轉換成一致的名稱。
        /// </summary>
        /// <param name="domainName"></param>
        /// <returns></returns>
        public static string Injection(this string domainName)
        {
            string dn = domainName.Trim();

            if (IsVariableDomain(dn))
                return "__彈性課程";
            else
                return dn;
        }

        public static void ClearSubjectScore(this IEnumerable<StudentScore> students, SemesterData semester)
        {
            foreach (StudentScore each in students)
            {
                if (each.SemestersScore.Contains(semester))
                    each.SemestersScore[semester].Subject.Clear();
            }
        }

        public static void ClearDomainScore(this IEnumerable<StudentScore> students, SemesterData semester)
        {
            foreach (StudentScore each in students)
            {
                if (each.SemestersScore.Contains(semester))
                    each.SemestersScore[semester].Domain.Clear();
            }
        }


        public static void ClearLearningDomainScore(this IEnumerable<StudentScore> students, SemesterData semester)
        {
            foreach (StudentScore each in students)
            {
                if (each.SemestersScore.Contains(semester))
                    each.SemestersScore[semester].LearnDomainScore = null;
            }
        }

        #endregion


        /// <summary>
        /// 取得評量比例設定
        /// </summary>
        public static Dictionary<string, decimal> GetScorePercentageHS()
        {
            Dictionary<string, decimal> returnData = new Dictionary<string, decimal>();
            FISCA.Data.QueryHelper qh1 = new FISCA.Data.QueryHelper();
            string query1 = @"select id,CAST(regexp_replace( xpath_string(exam_template.extension,'/Extension/ScorePercentage'), '^$', '0') as integer) as ScorePercentage  from exam_template";
            System.Data.DataTable dt1 = qh1.Select(query1);

            foreach (System.Data.DataRow dr in dt1.Rows)
            {
                string id = dr["id"].ToString();
                decimal sp = 50;
                if (decimal.TryParse(dr["ScorePercentage"].ToString(), out sp))
                {
                    returnData.Add(id, sp);
                }
                else
                    returnData.Add(id, 50);
            }
            return returnData;
        }

        /// <summary>
        /// 評量比例設定
        /// </summary>
        public static Dictionary<string, decimal> ScorePercentageHSDict = new Dictionary<string, decimal>();

        /// <summary>
        /// 載入 108 課綱領域對照
        /// </summary>
        public static void LoadDomainMap108()
        {
            DomainMap108List.Clear();
            DomainMap108List = new List<string>(new string[] { "語文", "數學", "自然科學", "科技", "社會", "藝術", "健康與體育", "綜合活動", "藝術與人文", "自然與生活科技" });

            KHSpecialDomainMapList.Clear();
            KHSpecialDomainMapList = new List<string>(new string[] { "國語文", "英語", "本土語文" });

        }

        /// <summary>
        ///  108 課綱領域對照
        /// </summary>
        public static List<string> DomainMap108List = new List<string>();

        /// <summary>
        ///  高雄國中特殊領域對照
        /// </summary>
        public static List<string> KHSpecialDomainMapList = new List<string>();

        /// <summary>
        /// 使用SQL傳入學年度、學期、學生系統編號，取得學生修課紀錄，目的在處理 K12 無法透過這條件，需要SelectAll學生修課造成記憶體不足。
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="StudentIDs"></param>
        /// <returns></returns>
        public static List<JHSCAttendRecord> GetSCAttendRecordListBySchoolYearSemsStudentIDs(int SchoolYear,int Semester,List<string> StudentIDs)
        {
            List<JHSCAttendRecord> value = new List<JHSCAttendRecord>();

            string qry = "SELECT " +
            "sc_attend.id" +
            " FROM" +
            " sc_attend" +
            " INNER JOIN" +
            " course" +
            " ON sc_attend.ref_course_id = course.id" +
            " WHERE sc_attend.ref_student_id IN(" + string.Join(",", StudentIDs.ToArray()) + ") AND course.school_year = " + SchoolYear + " AND course.semester = " + Semester;

            QueryHelper qh1 = new QueryHelper();
            DataTable dt = qh1.Select(qry);
            List<string> scIDList = new List<string>();
            foreach (DataRow dr in dt.Rows)
            {
                scIDList.Add(dr["id"] + "");
            }
            
            value = JHSCAttend.SelectByIDs(scIDList);

            // 節省記憶體
            dt = null;
            GC.Collect();

            return value;
        }


    }
}
