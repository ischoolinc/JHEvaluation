using System.Collections.Generic;
using System.Xml;
using JHSchool.Data;

namespace JHSchool.Evaluation.Calculation.GraduationConditions
{
    /// <summary>
    /// 學習領域成績是否符合最後學期條件
    /// </summary>
    internal class LearnDomainLastEval : IEvaluative
    {
        private EvaluationResult _result;
        private int _domain_count = 0;
        private decimal _score = 0;

        /// <summary>
        /// XML參數建構式
        /// <![CDATA[
        /// <條件 Checked="True" Type="LearnDomainLast" 學習領域="3" 等第="丙"/>
        /// ]]>
        /// </summary>
        /// <param name="element"></param>
        public LearnDomainLastEval(XmlElement element)
        {
            _result = new EvaluationResult();

            _domain_count = int.Parse(element.GetAttribute("學習領域"));
            string degree = element.GetAttribute("等第");

            //ConfigData cd = School.Configuration["等第對照表"];
            //if (!string.IsNullOrEmpty(cd["xml"]))
            //{
            //    XmlElement xml = XmlHelper.LoadXml(cd["xml"]);
            //    XmlElement scoreMapping = (XmlElement)xml.SelectSingleNode("ScoreMapping[@Name=\"" + degree + "\"]");
            //    decimal d;
            //    if (scoreMapping != null && decimal.TryParse(scoreMapping.GetAttribute("Score"), out d))
            //        _score = d;
            //}

            JHSchool.Evaluation.Mapping.DegreeMapper mapper = new JHSchool.Evaluation.Mapping.DegreeMapper();
            decimal? d = mapper.GetScoreByDegree(degree);
            if (d.HasValue) _score = d.Value;            
        }

        #region IEvaluative 成員

        public Dictionary<string, bool> Evaluate(IEnumerable<StudentRecord> list)
        {
            _result.Clear();

            Dictionary<string, bool> passList = new Dictionary<string, bool>();

            //Dictionary<string, SemesterHistoryUtility> shList = new Dictionary<string, SemesterHistoryUtility>();
            //foreach (Data.JHSemesterHistoryRecord shRecord in Data.JHSemesterHistory.SelectByStudentIDs(list.AsKeyList().ToArray()))
            //{
            //    if (!shList.ContainsKey(shRecord.RefStudentID))
            //        shList.Add(shRecord.RefStudentID, new SemesterHistoryUtility(shRecord));
            //}

            //list.SyncSemesterScoreCache();
            Dictionary<string, List<Data.JHSemesterScoreRecord>> studentSemesterScoreCache = new Dictionary<string, List<JHSchool.Data.JHSemesterScoreRecord>>();
            foreach (Data.JHSemesterScoreRecord record in Data.JHSemesterScore.SelectByStudentIDs(list.AsKeyList()))
            {
                if (!studentSemesterScoreCache.ContainsKey(record.RefStudentID))
                    studentSemesterScoreCache.Add(record.RefStudentID, new List<JHSchool.Data.JHSemesterScoreRecord>());
                studentSemesterScoreCache[record.RefStudentID].Add(record);
            }

            foreach (StudentRecord each in list)
            {
                List<ResultDetail> resultList = new List<ResultDetail>();
                JHSemesterHistoryRecord shRec= new JHSemesterHistoryRecord ();
                if(UIConfig._StudentSHistoryRecDict.ContainsKey(each.ID ))
                    shRec = UIConfig._StudentSHistoryRecDict[each.ID];
                // 有成績學年度學期
                List<string> hasSemsScoreSchoolYearSemester = new List<string>();
                if (studentSemesterScoreCache.ContainsKey(each.ID))
                {
                    foreach (Data.JHSemesterScoreRecord record in studentSemesterScoreCache[each.ID])
                    {
                        hasSemsScoreSchoolYearSemester.Add(record.SchoolYear.ToString() + record.Semester.ToString());
                        // 只檢查三下，以學期歷程為主                       
                        int gradeYear = 0;
                        foreach (K12.Data.SemesterHistoryItem shi in shRec.SemesterHistoryItems)
                            if (shi.SchoolYear == record.SchoolYear && shi.Semester == record.Semester)
                                gradeYear = shi.GradeYear;

                        if (!((gradeYear == 3 || gradeYear == 9) && record.Semester == 2)) continue;

                        // 領域有及格的數量
                        int count = 0;

                        foreach (K12.Data.DomainScore domain in record.Domains.Values)
                        {
                            if (domain.Score.HasValue && domain.Score.Value >= _score)
                                count++;
                        }

                        if (count < _domain_count)
                        {
                            ResultDetail rd = new ResultDetail(each.ID, "" + gradeYear, "" + record.Semester);
                            rd.AddMessage("學期領域成績不符合畢業規範");
                            rd.AddDetail("學期領域成績不符合畢業規範");
                            resultList.Add(rd);
                        }
                    }
                }

                    // 檢查有學期歷程沒有成績
                    foreach (K12.Data.SemesterHistoryItem shi in shRec.SemesterHistoryItems)
                    {
                        if (shi.GradeYear == 3 || shi.GradeYear == 9)
                        {
                            if(!hasSemsScoreSchoolYearSemester.Contains(shi.SchoolYear.ToString ()+ shi.Semester.ToString ()))
                            {
                                ResultDetail rd = new ResultDetail(each.ID, shi.GradeYear.ToString (), shi.Semester.ToString ());
                                rd.AddMessage("學期領域成績資料缺漏");
                                rd.AddDetail("學期領域成績資料缺漏");
                                resultList.Add(rd);
                            }
                        }
                    }

                if (resultList.Count > 0)
                {
                    _result.Add(each.ID, resultList);
                    passList.Add(each.ID, false);
                }
                else
                    passList.Add(each.ID, true);
            }

            return passList;
        }

        public EvaluationResult Result
        {
            get { return _result; }
        }

        #endregion
    }
}