﻿using System;
using System.Collections.Generic;
using System.Text;
using JHSchool.Data;
using JHEvaluation.ScoreCalculation.ScoreStruct;
using AEIncludeData = JHEvaluation.ScoreCalculation.ScoreStruct.TakeScore.AEIncludeData;
using System.Windows.Forms;

namespace JHEvaluation.ScoreCalculation.BigFunction
{
    class AttendScoreReader
    {
        private IStatusReporter Reporter { get; set; }

        private Dictionary<string, StudentScore> Students { get; set; }

        private int SchoolYear { get; set; }

        private int Semester { get; set; }

        private UniqueSet<string> FilterSubject { get; set; }

        private Dictionary<string, JHCourseRecord> Courses { get; set; }

        //AssessmentID,ExamID,YouKnow...
        private Dictionary<string, Dictionary<string, AEIncludeData>> AEIncludes { get; set; }

        private Dictionary<string, JHSCAttendRecord> Attends { get; set; }

        public AttendScoreReader(List<StudentScore> students, int schoolYear, int semester, IEnumerable<string> filterSubject, IStatusReporter reporter)
        {
            Students = students.ToDictionary();
            SchoolYear = schoolYear;
            Semester = semester;
            FilterSubject = new UniqueSet<string>(filterSubject);
            Reporter = reporter;
        }

        internal void Read()
        {
            Courses = ReadCourse();
            AEIncludes = ReadAEInclude();
            Attends = ReadCurrentSemesterAttend();
            List<JHSCETakeRecord> SCETakes = ReadSCETake();

            // 取得新竹評量樣版設定
            Util.ScorePercentageHSDict.Clear();
            Util.ScorePercentageHSDict = Util.GetScorePercentageHS();

            //清除原有的資料。
            foreach (StudentScore each in Students.Values)
                each.AttendScore.Clear();

            Reporter.Feedback("計算成績...", 0);
            int t1 = Environment.TickCount;

            foreach (JHSCAttendRecord Attend in Attends.Values)
            {
                if (!Students.ContainsKey(Attend.RefStudentID)) continue;
                StudentScore Student = Students[Attend.RefStudentID];

                if (!Courses.ContainsKey(Attend.RefCourseID)) continue;
                JHCourseRecord Course = Courses[Attend.RefCourseID];

                bool toSems = (Course.CalculationFlag == "1") ? true : false;

                string subjName = Course.Subject.Trim();

                // 如果科目名稱空白就不處理，因為無法計算到科目成績
                if (string.IsNullOrWhiteSpace(subjName))
                    continue;

                //傳入學年度,判斷103以前的學年度5:5而以後6:4比例計算
                if (!Student.AttendScore.Contains(subjName))
                    Student.AttendScore.Add(subjName, new AttendScore(Attend, Course.Credit, Course.Period, toSems, Course.Domain, Course.SchoolYear));
            }

            try
            {
                foreach (JHSCETakeRecord take in SCETakes)
                {
                    if (!Attends.ContainsKey(take.RefSCAttendID)) continue;
                    JHSCAttendRecord Attend = Attends[take.RefSCAttendID];

                    if (!Students.ContainsKey(Attend.RefStudentID)) continue;
                    StudentScore Student = Students[Attend.RefStudentID];

                    if (!Courses.ContainsKey(Attend.RefCourseID)) continue;
                    JHCourseRecord Course = Courses[Attend.RefCourseID];

                    if (!AEIncludes.ContainsKey(Course.RefAssessmentSetupID)) continue;
                    Dictionary<string, AEIncludeData> Include = AEIncludes[Course.RefAssessmentSetupID];

                    //沒有該次考試就不處理。
                    if (!Include.ContainsKey(take.RefExamID)) continue;

                    string subjName = Course.Subject.Trim();

                    // 如果科目名稱空白就不處理，因為無法計算到科目成績
                    if (string.IsNullOrWhiteSpace(subjName))
                        continue;

                    try
                    {
                        Student.AttendScore[subjName].Subscores.Add(take.RefExamID, new TakeScore(take, Include[take.RefExamID]));
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message);
                    }

                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }



            Reporter.Feedback("時間：" + (Environment.TickCount - t1), 0);
        }

        private List<JHSCETakeRecord> ReadSCETake()
        {
            Reporter.Feedback("讀取評量成績資料...", 0);

            List<string> attendids = new List<string>();
            foreach (JHSCAttendRecord each in Attends.Values)
                attendids.Add(each.ID);

            FunctionSpliter<string, JHSCETakeRecord> selectData = new FunctionSpliter<string, JHSCETakeRecord>(300 * 10, Util.MaxThread);
            selectData.Function = delegate (List<string> attendIdsPart)
            {
                return JHSCETake.Select(new string[] { }, new string[] { }, new string[] { }, new string[] { }, attendIdsPart);
            };
            selectData.ProgressChange = delegate (int progress)
            {
                Reporter.Feedback("讀取評量成績資料...", Util.CalculatePercentage(attendids.Count, progress));
            };
            return selectData.Execute(attendids);
        }

        private Dictionary<string, JHSCAttendRecord> ReadCurrentSemesterAttend()
        {
            Reporter.Feedback("讀取修課資料...", 0);

            List<string> studKeys = Students.Values.ToKeys();
            Dictionary<string, JHSCAttendRecord> currentAttends = new Dictionary<string, JHSCAttendRecord>();

            // CT 2020/7/6 因 Select All 效能問題，將它改寫
            ////分批取得資料。
            //FunctionSpliter<string, JHSCAttendRecord> spliter = new FunctionSpliter<string, JHSCAttendRecord>(300, Util.MaxThread);
            //spliter.Function = delegate(List<string> studKeysPart)
            //{
            //    return JHSCAttend.SelectByStudentIDs(studKeysPart);
            //};
            //spliter.ProgressChange = delegate(int progress)
            //{
            //    Reporter.Feedback("讀取修課資料...", Util.CalculatePercentage(studKeys.Count, progress));
            //};
            //List<JHSCAttendRecord> allAttends = spliter.Execute(studKeys);

            ////用於檢查重覆修習科目。
            ////Dictionary<string, Dictionary<string, string>> duplicate = new Dictionary<string, Dictionary<string, string>>();



            ////過濾修課記錄的學年度、學期，只要本學期的成績。
            //foreach (JHSCAttendRecord each in allAttends)
            //{
            //    if (!Courses.ContainsKey(each.RefCourseID)) continue;
            //    JHCourseRecord course = Courses[each.RefCourseID];

            //    if (!course.SchoolYear.HasValue) continue;
            //    if (!course.Semester.HasValue) continue;

            //    if (course.SchoolYear.Value != SchoolYear) continue;
            //    if (course.Semester.Value != Semester) continue;

            //    currentAttends.Add(each.ID, each);         


            //    //// 2016/5/25 穎驊新增 ，只有在該科目需要列入計算才更進一步加到currentAttends，如此一來可以濾掉科目名稱為空的社團課(不列入計算的課)
            //    //if (each.Course.CalculationFlag == "1")
            //    //{

            //    //    #region 檢查重覆修習科目。
            //    //    if (!duplicate.ContainsKey(each.RefStudentID))
            //    //        duplicate.Add(each.RefStudentID, new Dictionary<string, string>());

            //    //    if (duplicate[each.RefStudentID].ContainsKey(course.Subject.Trim()))
            //    //    {
            //    //        if (!Students.ContainsKey(each.RefStudentID)) continue;
            //    //        StudentScore student = Students[each.RefStudentID];


            //    //        //  2016/5/25 穎驊新增，針對如果同一學生同一學期有修習同一門科目名稱的課，會跳出視窗提醒。
            //    //        MessageBox.Show(string.Format("學生「{0}」在「{1}」學年「{2}」學期，重覆修習科目「{3}」，將使學期科目成績計算遺漏誤植，請確認並修正該科目。", student.Name, course.SchoolYear, course.Semester, course.Subject));

            //    //        break;

            //    //        //throw new ArgumentException(string.Format("學生「{0}」重覆修習科目「{1}」。", student.Name, course.Subject));
            //    //        //continue; //先略過不管。
            //    //    }
            //    //    duplicate[each.RefStudentID].Add(course.Subject.Trim(), null);
            //    //    #endregion;


            //    //    currentAttends.Add(each.ID, each);

            //    //}

            //}

            List<JHSCAttendRecord> currentAttendsList = Util.GetSCAttendRecordListBySchoolYearSemsStudentIDs(SchoolYear, Semester, studKeys);

            foreach (JHSCAttendRecord rec in currentAttendsList)
            {
                if (!currentAttends.ContainsKey(rec.ID))
                    currentAttends.Add(rec.ID, rec);
            }

            return currentAttends;
        }

        private Dictionary<string, Dictionary<string, AEIncludeData>> ReadAEInclude()
        {
            Reporter.Feedback("讀取評量設定資料...", 0);

            Dictionary<string, Dictionary<string, AEIncludeData>> includes = new Dictionary<string, Dictionary<string, AEIncludeData>>();
            List<JHAEIncludeRecord> lstIncludes = JHAEInclude.SelectAll();

            foreach (JHAEIncludeRecord each in lstIncludes)
            {
                if (!includes.ContainsKey(each.RefAssessmentSetupID))
                    includes.Add(each.RefAssessmentSetupID, new Dictionary<string, AEIncludeData>());

                includes[each.RefAssessmentSetupID].Add(each.RefExamID, new AEIncludeData(each));
            }

            return includes;
        }

        private Dictionary<string, JHCourseRecord> ReadCourse()
        {
            Reporter.Feedback("讀取課程資料...", 0);

            Dictionary<string, JHCourseRecord> courses = new Dictionary<string, JHCourseRecord>();
            JHCourse.RemoveAll();
            List<JHCourseRecord> lstCourses = JHCourse.SelectBySchoolYearAndSemester(SchoolYear, Semester);

            foreach (JHCourseRecord each in lstCourses)
            {
                //過慮掉不計算的科目。
                if (!IsValidSubject(each))
                    continue;

                if (!courses.ContainsKey(each.ID))
                    courses.Add(each.ID, each);
            }

            return courses;
        }

        private bool IsValidSubject(JHCourseRecord each)
        {
            if (FilterSubject.Count > 0) //大於零才過慮。
                return FilterSubject.Contains(each.Subject.Trim());
            else //不然就通通都要算。
                return true;
        }

        private List<string> GetCourseKeys()
        {
            List<string> keys = new List<string>();

            foreach (JHCourseRecord each in Courses.Values)
                keys.Add(each.ID);

            return keys;
        }
    }
}
