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

            foreach (StudentRecord each in list)
            {
                List<ResultDetail> resultList = new List<ResultDetail>();
           

                // 有學期成績
                if (studentSemesterScoreCache.ContainsKey(each.ID))
                {
                    // 存放畢業領域成績
                    Dictionary<string, decimal> graduateDomainScoreDict = new Dictionary<string, decimal>();
                    Dictionary<string, decimal> graduateDomainCreditDict = new Dictionary<string, decimal>();
                    Dictionary<string, int> graduateDomainCountDict = new Dictionary<string, int>();
                    // 存放符合標準畢業領域成績
                    List<decimal> passScoreList = new List<decimal>();
                    // 取得學生學生領域成績填入計算畢業成績用
                    foreach (JHSemesterScoreRecord record in studentSemesterScoreCache[each.ID])
                    {
                        foreach (K12.Data.DomainScore domain in record.Domains.Values)
                        {
                            // 國語文、英語，轉成語文領域對應
                            string domainName = domain.Domain;

                            if (domain.Domain == "國語文" || domain.Domain == "英語")
                                domainName = "語文";

                            if (!graduateDomainScoreDict.ContainsKey(domainName))
                            {
                                graduateDomainScoreDict.Add(domainName, 0);
                                graduateDomainCountDict.Add(domainName, 0);
                                graduateDomainCreditDict.Add(domainName, 0);
                            }
                            if (domain.Score.HasValue && domain.Credit.HasValue)
                            {
                                graduateDomainScoreDict[domainName] += (domain.Score.Value*domain.Credit.Value);
                                graduateDomainCreditDict[domainName] += domain.Credit.Value;
                                graduateDomainCountDict[domainName]++;
                            }
                        }
                    }

                    // 即時計算畢業成績並判斷是否符合
                    foreach (string name in graduateDomainScoreDict.Keys)
                    {
                        decimal grScore = 0;
                        // if (graduateDomainCountDict[name] > 0)
                        if (graduateDomainCreditDict[name] > 0) // 小郭, 2013/12/30
                        {
                            //// 算術平均
                            //grScore = graduateDomainScoreDict[name] / graduateDomainCountDict[name];

                            // 加權平均,加權總分,加權學分
                            grScore = graduateDomainScoreDict[name] / graduateDomainCreditDict[name];

                            if (grScore >= _score)
                                passScoreList.Add(grScore);
                            
                            // 小郭, 2013/12/30
                            StudentDomainResult.AddDomain(each.ID, name, grScore, grScore >= _score);
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
