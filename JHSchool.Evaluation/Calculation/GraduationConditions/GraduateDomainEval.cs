using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using System.Xml;
using K12.Data;
using JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReportControls;

namespace JHSchool.Evaluation.Calculation.GraduationConditions
{
    /// <summary>
    /// 所有學習領域符合畢業總平均成績規範。
    /// </summary>
    internal class GraduateDomainEval : IEvaluative
    {
        private EvaluationResult _result;
        private int _domain_count = 0;
        private decimal _score = 0;

        /// <summary>
        /// XML參數建構式
        /// <![CDATA[ 
        /// <條件 Checked="True" Type="GraduateDomain" 學習領域="3" 等第="甲"/>
        /// ]]>
        /// </summary>
        /// <param name="element"></param>
        public GraduateDomainEval(XmlElement element)
        {
            _result = new EvaluationResult();

            _domain_count = int.Parse(element.GetAttribute("學習領域"));
            string degree = element.GetAttribute("等第");

            JHSchool.Evaluation.Mapping.DegreeMapper mapper = new JHSchool.Evaluation.Mapping.DegreeMapper();
            decimal? d = mapper.GetScoreByDegree(degree);
            if (d.HasValue) _score = d.Value;
        }

        public Dictionary<string, bool> Evaluate(IEnumerable<StudentRecord> list)
        {
            _result.Clear();

            Dictionary<string, bool> passList = new Dictionary<string, bool>();
            
            Dictionary<string, List<JHSemesterScoreRecord>> studentSemesterScoreCache = new Dictionary<string, List<JHSemesterScoreRecord>>();
            foreach (JHSemesterScoreRecord record in JHSemesterScore.SelectByStudentIDs(list.AsKeyList()))
            {
                if (!studentSemesterScoreCache.ContainsKey(record.RefStudentID))
                    studentSemesterScoreCache.Add(record.RefStudentID, new List<JHSemesterScoreRecord>());
                studentSemesterScoreCache[record.RefStudentID].Add(record);
            }

            //學生能被承認的學年度學期對照
            Dictionary<string, List<string>> studentSYSM = new Dictionary<string, List<string>>();

            #region 學期歷程
            foreach (K12.Data.SemesterHistoryRecord shr in K12.Data.SemesterHistory.SelectByStudentIDs(list.Select(x => x.ID)))
            {
                if (!studentSYSM.ContainsKey(shr.RefStudentID))
                    studentSYSM.Add(shr.RefStudentID, new List<string>());

                Dictionary<string, K12.Data.SemesterHistoryItem> check = new Dictionary<string, K12.Data.SemesterHistoryItem>() { 
                {"1a",null },
                {"1b",null },
                {"2a",null },
                {"2b",null },
                {"3a",null },
                {"3b",null }
                };

                foreach (K12.Data.SemesterHistoryItem item in shr.SemesterHistoryItems)
                {
                    string grade = item.GradeYear + "";

                    if (grade == "7") grade = "1";
                    if (grade == "8") grade = "2";
                    if (grade == "9") grade = "3";

                    if (grade == "1" || grade == "2" || grade == "3")
                    {
                        string key = "";
                        if (item.Semester == 1)
                            key = grade + "a";
                        else if (item.Semester == 2)
                            key = grade + "b";
                        else
                            continue;

                        //相同年級取較新的學年度
                        if (check[key] == null)
                            check[key] = item;
                        else if (item.SchoolYear > check[key].SchoolYear)
                            check[key] = item;
                    }
                }

                foreach (string key in check.Keys)
                {
                    if (check[key] == null)
                        continue;

                    K12.Data.SemesterHistoryItem item = check[key];

                    studentSYSM[shr.RefStudentID].Add(item.SchoolYear + "_" + item.Semester);
                }
            }

            #endregion

            foreach (StudentRecord each in list)
            {
                List<ResultDetail> resultList = new List<ResultDetail>();
           
                // 有學期成績
                if (studentSemesterScoreCache.ContainsKey(each.ID))
                {
                    // 存放符合標準畢業領域成績
                    List<decimal> passScoreList = new List<decimal>();

                    //每個學期整理後的成績
                    List<List<K12.Data.DomainScore>> GradeScoreList = new List<List<K12.Data.DomainScore>>();

                    // 取得學生學生領域成績填入計算畢業成績用
                    foreach (JHSemesterScoreRecord record in studentSemesterScoreCache[each.ID])
                    {
                        string key = record.SchoolYear + "_" + record.Semester;

                        //只處理承認的學年度學期
                        if (!studentSYSM.ContainsKey(each.ID) || !studentSYSM[each.ID].Contains(key))
                            continue;

                        //整理後的領域成績
                        List<K12.Data.DomainScore> domainScoreList = new List<K12.Data.DomainScore>();

                        K12.Data.DomainScore 語文 = new K12.Data.DomainScore();
                        語文.Domain = "語文";
                        
                        decimal sum = 0;
                        decimal credit = 0;

                        //跑一遍領域成績
                        foreach (K12.Data.DomainScore domain in record.Domains.Values)
                        {
                            //這三種挑出來處理
                            if (domain.Domain == "國語文" || domain.Domain == "英語")
                            {
                                if (domain.Score.HasValue && domain.Credit.HasValue)
                                {
                                    sum += domain.Score.Value * domain.Credit.Value;
                                    credit += domain.Credit.Value;

                                    //處理高雄語文顯示
                                    //  加權總分
                                    if (!TempData.tmpStudDomainScoreDict.ContainsKey(record.RefStudentID))
                                        TempData.tmpStudDomainScoreDict.Add(record.RefStudentID, new Dictionary<string, decimal>());

                                    if (!TempData.tmpStudDomainCreditDict.ContainsKey(record.RefStudentID))
                                        TempData.tmpStudDomainCreditDict.Add(record.RefStudentID, new Dictionary<string, decimal>());

                                    if (!TempData.tmpStudDomainScoreDict[record.RefStudentID].ContainsKey(domain.Domain))
                                        TempData.tmpStudDomainScoreDict[record.RefStudentID].Add(domain.Domain, 0);

                                    // 學分數
                                    if (!TempData.tmpStudDomainCreditDict[record.RefStudentID].ContainsKey(domain.Domain))
                                        TempData.tmpStudDomainCreditDict[record.RefStudentID].Add(domain.Domain, 0);

                                    TempData.tmpStudDomainScoreDict[record.RefStudentID][domain.Domain] += (domain.Score.Value * domain.Credit.Value);
                                    TempData.tmpStudDomainCreditDict[record.RefStudentID][domain.Domain] += domain.Credit.Value;
                                }
                            }
                            else
                                domainScoreList.Add(domain);
                        }

                        if (credit > 0)
                        {
                            語文.Score = Math.Round(sum / credit, 2, MidpointRounding.AwayFromZero);
                            語文.Credit = credit;

                            domainScoreList.Add(語文);
                        }

                        //會被加入就代表承認了
                        GradeScoreList.Add(domainScoreList);
                    }

                    Dictionary<string, decimal> domainScoreSum = new Dictionary<string, decimal>();
                    Dictionary<string, decimal> domainScoreCount = new Dictionary<string, decimal>();

                    foreach (List<K12.Data.DomainScore> scoreList in GradeScoreList)
                    {
                        foreach (K12.Data.DomainScore ds in scoreList)
                        {
                            string domainName = ds.Domain;

                            if (!domainScoreSum.ContainsKey(domainName))
                                domainScoreSum.Add(domainName, 0);
                            if (!domainScoreCount.ContainsKey(domainName))
                                domainScoreCount.Add(domainName, 0);

                            if(ds.Score.HasValue)
                            {
                                domainScoreSum[domainName] += ds.Score.Value;
                                //同一學期不會有相同領域名稱,可直接作++
                                domainScoreCount[domainName]++;
                            }
                        }
                    }

                    foreach (string domainName in domainScoreCount.Keys)
                    {
                        if (domainScoreCount[domainName] > 0)
                        {
                            decimal grScore = Math.Round(domainScoreSum[domainName] / domainScoreCount[domainName], 2, MidpointRounding.AwayFromZero);

                            if(grScore >= _score)
                                passScoreList.Add(grScore);

                            StudentDomainResult.AddDomain(each.ID, domainName, grScore, grScore >= _score);
                        }
                    }

                    // 當及格數小於標準數，標示不符格畢業規範
                    if (passScoreList.Count < _domain_count)
                    {
                        ResultDetail rd = new ResultDetail(each.ID, "0", "0");
                        rd.AddMessage("領域畢業加權總平均成績不符合畢業規範");
                        rd.AddDetail("領域畢業加權總平均成績不符合畢業規範");
                        resultList.Add(rd);
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
    }
}
