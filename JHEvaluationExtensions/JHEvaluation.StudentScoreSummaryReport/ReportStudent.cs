﻿using System;
using System.Collections.Generic;
using System.Text;
using JHSchool.Data;
using System.Xml;
using JHEvaluation.ScoreCalculation;
using System.Globalization;
using K12.Data;
using System.Drawing;
using System.IO;
using System.Data;

namespace JHEvaluation.StudentScoreSummaryReport
{
    internal class ReportStudent : StudentScore, Campus.Rating.IStudent
    {
        public ReportStudent(JHStudentRecord student)
            : base(student)
        {
            Places = new Campus.Rating.PlaceCollection();
            HeaderList = new ReportHeaderList();
            Gender = student.Gender;
            Birthday = student.Birthday.HasValue ? student.Birthday.Value.ToString("yyyy/MM/dd") : "";
            Summaries = new Dictionary<SemesterData, XmlElement>();
            EntranceDate = string.Empty;
            GraduateDate = string.Empty;
            EnglishName = student.EnglishName;
            IDNumber = student.IDNumber;
            StudentStatus = student.Status;
            StudentID = student.ID;
            nationality1 = string.Empty;
            passport_name1 = string.Empty;
            nationality2 = string.Empty;
            passport_name2 = string.Empty;
            Enationality1 = string.Empty;
            Enationality2 = string.Empty;
        }

        public StudentRecord.StudentStatus StudentStatus { get; private set; }

        public string Gender { get; private set; }

        public string Birthday { get; private set; }

        public string EnglishName { get; private set; }

        public string IDNumber { get; private set; }

        public string StudentID { get; private set; }

        public Image GraduatePhoto { get; set; }

        /// <summary>
        /// 學期的 Header Index。
        /// </summary>
        public ReportHeaderList HeaderList { get; private set; }

        #region IStudent 成員

        public Campus.Rating.PlaceCollection Places { get; private set; }

        #endregion

        public Dictionary<SemesterData, XmlElement> Summaries { get; private set; }

        /// <summary>
        /// 學生入學日期(民國格式)。
        /// </summary>
        public string EntranceDate { get; set; }

        /// <summary>
        /// 學生入學日期(西元格式)。
        /// </summary>
        public string EngEntranceDate { get; set; }

        /// <summary>
        /// 學生入學日期(民國格式)。
        /// </summary>
        public string GraduateDate { get; set; }

        /// <summary>
        /// 學生入學日期(西元格式)。
        /// </summary>
        public string EngGraduateDate { get; set; }

        public string nationality1 { get; set; }
        public string passport_name1 { get; set; }
        public string nationality2 { get; set; }
        public string passport_name2 { get; set; }
        public string Enationality1 { get; set; }
        public string Enationality2 { get; set; }

        public string PrintOrder
        {
            // 原本OrderString的順序有問題，故改用下列方式設定列印順序
            //年級//班級排列序號/班級名稱//座號
            get
            {
                if (Class.DisplayOrder == null || Class.DisplayOrder == "")
                    return GradeYear.ToString().PadLeft(3, '0') + Class.DisplayOrder.PadLeft(3, 'Z') + ClassName + SeatNo.PadLeft(3, '0');
                else
                    return GradeYear.ToString().PadLeft(3, '0') + Class.DisplayOrder.PadLeft(3, '0') + ClassName + SeatNo.PadLeft(3, '0');
            }
        }
    }

    internal class ReportHeaderList : Dictionary<SemesterData, int>
    {
        private Dictionary<SemesterData, SemesterData> RawList = new Dictionary<SemesterData, SemesterData>(10);

        public ReportHeaderList()
        {
        }

        /// <summary>
        /// 包含學年度學期年級的 SemesterData。
        /// </summary>
        /// <param name="semester"></param>
        /// <param name="count"></param>
        public void AddRaw(SemesterData semester, int count)
        {
            SemesterData sd = new SemesterData(0, semester.SchoolYear, semester.Semester);

            RawList.Add(sd, semester); //原始的 SemesterData。
            Add(sd, count);
        }

        /// <summary>
        /// 用「學年度、學期」取得原始的 SemesterData。
        /// </summary>
        /// <param name="semester"></param>
        /// <returns></returns>
        public SemesterData GetSRaw(SemesterData semester)
        {
            SemesterData sd;

            if (RawList.TryGetValue(semester, out sd))
                return sd;
            else
                return SemesterData.Empty;
        }
    }

    internal static class StudentScore_Extens
    {
        #region ReadNationalityData
        public static void ReadNationalityData(this IEnumerable<ReportStudent> students, IStatusReporter reporter)
        {
            //加入護照資料
            foreach (ReportStudent student in students)
            {
                FISCA.Data.QueryHelper qh1 = new FISCA.Data.QueryHelper();
                string strSQL1 = "select nationality1, passport_name1, nat1.eng_name as nat_eng1, nationality2, passport_name2, nat2.eng_name as nat_eng2, nationality3, passport_name3, nat3.eng_name as nat_eng3 from student_info_ext  as stud_info left outer join $ischool.mapping.nationality as nat1 on nat1.name = stud_info.nationality1 left outer join $ischool.mapping.nationality as nat2 on nat2.name = stud_info.nationality2 left outer join $ischool.mapping.nationality as nat3 on nat3.name = stud_info.nationality3 WHERE ref_student_id=" + student.StudentID;
                DataTable student_info_ext = qh1.Select(strSQL1);
                if (student_info_ext.Rows.Count > 0)
                {
                    student.nationality1 = student_info_ext.Rows[0]["nationality1"].ToString();
                    student.passport_name1 = student_info_ext.Rows[0]["passport_name1"].ToString();
                    student.nationality2 = student_info_ext.Rows[0]["nationality2"].ToString();
                    student.passport_name2 = student_info_ext.Rows[0]["passport_name2"].ToString();
                    student.Enationality1 = student_info_ext.Rows[0]["nat_eng1"].ToString();
                    student.Enationality2 = student_info_ext.Rows[0]["nat_eng2"].ToString();
                }
                else
                {
                    student.nationality1 = "";
                    student.passport_name1 = "";
                    student.nationality2 = "";
                    student.passport_name2 = "";
                    student.Enationality1 = "";
                    student.Enationality2 = "";

                }
            }
        }
        #endregion

        #region ReadUpdateRecordDate
        public static void ReadUpdateRecordDate(this IEnumerable<ReportStudent> students, IStatusReporter reporter)
        {
            int t1 = Environment.TickCount;
            Dictionary<string, ReportStudent> dicstudents = students.ToDictionary();
            List<string> keys = students.ToSC().ToKeys();

            Campus.FunctionSpliter<string, JHUpdateRecordRecord> selectData = new Campus.FunctionSpliter<string, JHUpdateRecordRecord>(500, 5);
            selectData.Function = delegate (List<string> ps)
            {
                return JHUpdateRecord.SelectByStudentIDs(ps);
            };
            List<JHUpdateRecordRecord> updaterecords = selectData.Execute(keys);

            Dictionary<string, List<JHUpdateRecordRecord>> dicupdaterecords = new Dictionary<string, List<JHUpdateRecordRecord>>();
            string ValidCodes = "1:2";  //新生:1 轉入:3 復學:6，畢業:2
            foreach (JHUpdateRecordRecord each in updaterecords)
            {
                //不是要處理的代碼，就跳過。
                if (ValidCodes.IndexOf(each.UpdateCode) < 0) continue;

                if (!dicupdaterecords.ContainsKey(each.StudentID))
                    dicupdaterecords.Add(each.StudentID, new List<JHUpdateRecordRecord>());
                dicupdaterecords[each.StudentID].Add(each);
            }

            foreach (KeyValuePair<string, List<JHUpdateRecordRecord>> each in dicupdaterecords)
            {
                each.Value.Sort(delegate (JHUpdateRecordRecord x, JHUpdateRecordRecord y)
                {
                    DateTime xx, yy;

                    if (!DateTime.TryParse(x.UpdateDate, out xx))
                        xx = DateTime.MinValue;

                    if (!DateTime.TryParse(y.UpdateDate, out yy))
                        yy = DateTime.MinValue;

                    return xx.CompareTo(yy);
                });
            }

            string ECodes = "1"; //入學
            string GCodes = "2";    //畢業
            foreach (KeyValuePair<string, List<JHUpdateRecordRecord>> each in dicupdaterecords)
            {
                if (!dicstudents.ContainsKey(each.Key)) continue;
                ReportStudent student = dicstudents[each.Key];

                JHUpdateRecordRecord e = each.Value[0];
                JHUpdateRecordRecord g = each.Value[each.Value.Count - 1];

                if (ECodes.IndexOf(e.UpdateCode) >= 0)
                {
                    DateTime dt;

                    if (DateTime.TryParse(e.UpdateDate, out dt))
                    {
                        student.EntranceDate = string.Format("{0}/{1}/{2}", dt.Year - 1911, dt.Month, dt.Day);
                        student.EngEntranceDate = dt.ToString(Util.EnglishFormat, Util.USCulture);
                    }
                }

                if (GCodes.IndexOf(g.UpdateCode) >= 0)
                {
                    DateTime dt;

                    if (DateTime.TryParse(g.UpdateDate, out dt))
                    {
                        student.GraduateDate = string.Format("{0}/{1}/{2}", dt.Year - 1911, dt.Month, dt.Day);
                        student.EngGraduateDate = dt.ToString(Util.EnglishFormat, Util.USCulture);
                    }
                }
            }

            Console.WriteLine(Environment.TickCount - t1);
        }
        #endregion

        public static void ReadGraduatePhoto(this IEnumerable<ReportStudent> students, IStatusReporter reporter)
        {
            Dictionary<string, ReportStudent> dicstudents = students.ToDictionary();
            List<string> keys = students.ToSC().ToKeys();

            Campus.FunctionSpliter<string, PhotoRecord> selectData = new Campus.FunctionSpliter<string, PhotoRecord>(300, 5);
            selectData.Function = delegate (List<string> ps)
            {
                List<PhotoRecord> photos = new List<PhotoRecord>();

                foreach (KeyValuePair<string, string> each in Photo.SelectGraduatePhoto(ps))
                {
                    if (each.Value == "")
                        photos.Add(new PhotoRecord(each.Key, Photo.SelectFreshmanPhoto(each.Key)));
                    else
                        photos.Add(new PhotoRecord(each.Key, each.Value));


                }
                return photos;
            };
            List<PhotoRecord> updaterecords = selectData.Execute(keys);

            foreach (PhotoRecord each in updaterecords)
            {
                if (dicstudents.ContainsKey(each.ID))
                    dicstudents[each.ID].GraduatePhoto = each.Photo;
            }
        }

        private class PhotoRecord
        {
            public PhotoRecord(string id, string photo)
            {
                ID = id;

                if (!string.IsNullOrEmpty(photo))
                {
                    Stream imgstream = new MemoryStream(Convert.FromBase64String(photo));
                    imgstream.Seek(0, SeekOrigin.Begin);

                    Photo = Image.FromStream(imgstream);
                }
                else
                    Photo = null;
            }

            public string ID { get; private set; }

            public Image Photo { get; private set; }
        }
    }
}
