using System;
using System.Collections.Generic;
using System.Text;
using JHEvaluation.ScoreCalculation.ScoreStruct;
using SCSemsScore = JHEvaluation.ScoreCalculation.ScoreStruct.SemesterScore;
using System.Linq;

namespace JHEvaluation.ScoreCalculation.BigFunction
{
    class SemesterScoreCalculator
    {
        private List<StudentScore> Students { get; set; }

        private UniqueSet<string> Filter { get; set; }

        public SemesterScoreCalculator(List<StudentScore> students, IEnumerable<string> filter)
        {
            Students = students;
            Filter = new UniqueSet<string>(filter);
        }

        /// <summary>
        /// 計算科目成績
        /// </summary>
        public void CalculateSubjectScore()
        {
            foreach (StudentScore student in Students)
            {
                if (student.CalculationRule == null) continue; //沒有成績計算規則就不計算。

                foreach (string subject in student.AttendScore)
                {
                    if (!IsValidItem(subject)) continue; //慮掉不算的科目。

                    AttendScore attend = student.AttendScore[subject];
                    SemesterSubjectScore SemsSubj = null;
                    SCSemsScore Sems = student.SemestersScore[SemesterData.Empty];

                    if (Sems.Subject.Contains(subject))
                        SemsSubj = Sems.Subject[subject];
                    else
                        SemsSubj = new SemesterSubjectScore();

                    if (attend.ToSemester) //列入學期成績。
                    {
                        //分數、權數都有資料才進行計算科目成績。
                        if (attend.Value.HasValue && attend.Weight.HasValue && attend.Period.HasValue)
                        {
                            decimal score = student.CalculationRule.ParseSubjectScore(attend.Value.Value); //進位處理。
                            SemsSubj.Value = score;
                            SemsSubj.Weight = attend.Weight.Value;
                            SemsSubj.Period = attend.Period.Value;
                            SemsSubj.Effort = attend.Effort;
                            SemsSubj.Text = attend.Text;
                            SemsSubj.Domain = attend.Domain;
                        }
                        else //資料不合理，保持原來的分數狀態。
                            continue;

                        if (!Sems.Subject.Contains(subject))
                            Sems.Subject.Add(subject, SemsSubj);
                    }
                    else
                    {
                        //不計入學期成績，就將其刪除掉。
                        if (Sems.Subject.Contains(subject))
                            Sems.Subject.Remove(subject);
                    }
                }
            }
        }

        /// <summary>
        /// 計算領域成績
        /// </summary>
        /// <param name="defaultRule"></param>
        public void CalculateDomainScore(ScoreCalculator defaultRule)
        {
            EffortMap effortmap = new EffortMap(); //努力程度對照表。

            foreach (StudentScore student in Students)
            {
                SemesterScore semsscore = student.SemestersScore[SemesterData.Empty];
                SemesterDomainScoreCollection dscores = semsscore.Domain;
                SemesterSubjectScoreCollection jscores = semsscore.Subject;
                ScoreCalculator rule = student.CalculationRule;

                if (rule == null)
                    rule = defaultRule;

                //各領域分數加總。
                Dictionary<string, decimal> domainTotal = new Dictionary<string, decimal>();
                //各領域權重加總。
                Dictionary<string, decimal> domainWeight = new Dictionary<string, decimal>();
                //各領域節數加總。
                Dictionary<string, decimal> domainPeriod = new Dictionary<string, decimal>();
                //文字評量串接。
                Dictionary<string, string> domainText = new Dictionary<string, string>();

                // 只計算領域成績，不計算科目
                bool OnlyCalcDomainScore = false;

                // 檢查科目成績是否有成績，當有成績且成績的科目領域名稱非空白才計算
                int CheckSubjDomainisNotNullCot = 0;

                // 檢查是否有非科目領域是空白的科目
                foreach (string str in jscores)
                {
                    SemesterSubjectScore objSubj = jscores[str];
                    string strDomainName = objSubj.Domain.Trim();
                    if (!string.IsNullOrEmpty(strDomainName))
                        CheckSubjDomainisNotNullCot++;
                }

                // 當沒有科目成績或科目成績內領域沒有非空白，只計算領域成績。
                if (jscores.Count == 0 || CheckSubjDomainisNotNullCot == 0)
                    OnlyCalcDomainScore = true;
                else
                    OnlyCalcDomainScore = false;


                if (OnlyCalcDomainScore == false)
                {
                    // 從科目計算到領域

                    #region 總計各領域的總分、權重、節數。
                    foreach (string strSubj in jscores)
                    {
                        SemesterSubjectScore objSubj = jscores[strSubj];
                        string strDomain = objSubj.Domain.Injection();

                        //不計算的領域就不算。
                        if (!IsValidItem(strDomain)) continue;

                        if (objSubj.Value.HasValue && objSubj.Weight.HasValue && objSubj.Period.HasValue)
                        {
                            if (!domainTotal.ContainsKey(strDomain))
                            {
                                domainTotal.Add(strDomain, 0);
                                domainWeight.Add(strDomain, 0);
                                domainPeriod.Add(strDomain, 0);
                                domainText.Add(strDomain, string.Empty);
                            }

                            domainTotal[strDomain] += objSubj.Value.Value * objSubj.Weight.Value;
                            domainWeight[strDomain] += objSubj.Weight.Value;
                            domainPeriod[strDomain] += objSubj.Period.Value;
                            domainText[strDomain] += GetDomainSubjectText(strSubj, objSubj.Text);
                        }
                    }
                    #endregion

                    #region 計算各領域加權平均。

                    /* (2014/11/18 補考調整)調整為保留領域成績中的資訊，但是移除多的領域成績項目。 */
                    // 從科目算過來清空原來領域成績，以科目成績的領域為主
                    //dscores.Clear();

                    foreach (string strDomain in domainTotal.Keys)
                    {
                        decimal total = domainTotal[strDomain];
                        decimal weight = domainWeight[strDomain];
                        decimal period = domainPeriod[strDomain];
                        string text = string.Join(";", domainText[strDomain].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries));

                        if (weight <= 0) continue; //沒有權重就不計算，保留原來的成績。

                        decimal weightAvg = rule.ParseDomainScore(total / weight);

                        //將成績更新回學生。
                        SemesterDomainScore dscore = null;
                        if (dscores.Contains(strDomain))
                            dscore = dscores[strDomain];
                        else
                        {
                            dscore = new SemesterDomainScore();
                            dscores.Add(strDomain, dscore);
                        }

                        dscore.ScoreOrigin = weightAvg;
                        dscore.Weight = weight;
                        dscore.Period = period;
                        dscore.Text = text;
                        dscore.Effort = effortmap.GetCodeByScore(weightAvg);

                        dscore.BetterScoreSelection(); //進行成績擇優。
                    }

                    foreach (var domainName in dscores.ToArray())
                    {
                        //如果新計算的領域成績中不包含在原領域清單中，就移除他。
                        if (!domainTotal.ContainsKey(domainName))
                            dscores.Remove(domainName);
                    }
                }
                else
                {
                    // 因為原本就會計算學期領域成績，所以不處理。
                }

                //計算課程學習成績。
                ScoreResult result = CalcDomainWeightAvgScore(dscores, new UniqueSet<string>());
                if (result.Score.HasValue)
                    semsscore.CourseLearnScore = rule.ParseLearnDomainScore(result.Score.Value);
                if (result.ScoreOrigin.HasValue)
                    semsscore.CourseLearnScoreOrigin = rule.ParseLearnDomainScore(result.ScoreOrigin.Value);

                //計算學習領域成績。
                result = CalcDomainWeightAvgScore(dscores, Util.VariableDomains);
                if (result.Score.HasValue)
                    semsscore.LearnDomainScore = rule.ParseLearnDomainScore(result.Score.Value);
                if (result.ScoreOrigin.HasValue)
                    semsscore.LearnDomainScoreOrigin = rule.ParseLearnDomainScore(result.ScoreOrigin.Value);
                    #endregion
            }
        }

        /// <summary>
        /// 單純從學期科目成績的文字描述加總到學期領域成績的文字描述
        /// </summary>
        /// <param name="defaultRule"></param>
        public void SumDomainTextScore()
        {
            foreach (StudentScore student in Students)
            {
                SemesterScore semsscore = student.SemestersScore[SemesterData.Empty];
                SemesterDomainScoreCollection dscores = semsscore.Domain;
                SemesterSubjectScoreCollection jscores = semsscore.Subject;

                Dictionary<string, string> domainText = new Dictionary<string, string>();

                #region 總計各領域的總分、權重、節數。
                foreach (string strSubj in jscores)
                {
                    SemesterSubjectScore objSubj = jscores[strSubj];
                    string strDomain = objSubj.Domain.Injection();

                    //不計算的領域就不算。
                    if (!IsValidItem(strDomain)) continue;

                    if (objSubj.Value.HasValue && objSubj.Weight.HasValue && objSubj.Period.HasValue)
                    {
                        if (!domainText.ContainsKey(strDomain))
                        {
                            domainText.Add(strDomain, string.Empty);
                        }

                        domainText[strDomain] += GetDomainSubjectText(strSubj, objSubj.Text);
                    }
                }
                #endregion

                #region 計算各領域加權平均。
                foreach (string strDomain in domainText.Keys)
                {
                    string text = string.Join(";", domainText[strDomain].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries));

                    //將成績更新回學生。
                    SemesterDomainScore dscore = null;
                    if (dscores.Contains(strDomain))
                        dscore = dscores[strDomain];
                    else
                    {
                        dscore = new SemesterDomainScore();
                        dscores.Add(strDomain, dscore);
                    }

                    dscore.Text = text;
                }
                #endregion
            }
        }

        /// <summary>
        /// 計算學習領域成績
        /// </summary>
        /// <param name="defaultRule"></param>
        public void CalculateLearningDomainScore(ScoreCalculator defaultRule)
        {
            foreach (StudentScore student in Students)
            {
                SemesterScore semsscore = student.SemestersScore[SemesterData.Empty];
                SemesterDomainScoreCollection dscores = semsscore.Domain;
                ScoreCalculator rule = student.CalculationRule;

                if (rule == null)
                    rule = defaultRule;

                #region 計算各領域加權平均。
                //計算學習領域成績。
                //decimal? result = CalcDomainWeightAvgScore(dscores, Util.VariableDomains);
                ScoreResult result = CalcDomainWeightAvgScore(dscores, Util.VariableDomains); ;
                if (result.Score.HasValue)
                    semsscore.LearnDomainScore = rule.ParseLearnDomainScore(result.Score.Value);

                if (result.ScoreOrigin.HasValue)
                    semsscore.LearnDomainScoreOrigin = rule.ParseLearnDomainScore(result.ScoreOrigin.Value);
                #endregion

            }
        }

        private static ScoreResult CalcDomainWeightAvgScore(SemesterDomainScoreCollection dscores, UniqueSet<string> excludeItem)
        {
            decimal TotalScore = 0, TotalScoreOrigin = 0, TotalWeight = 0;
            foreach (string strDomain in dscores)
            {
                if (excludeItem.Contains(strDomain.Trim())) continue;

                SemesterDomainScore dscore = dscores[strDomain];
                if (dscore.Value.HasValue && dscore.Weight.HasValue) //dscore.Value 是原來的結構。
                {
                    TotalScore += (dscore.Value.Value * dscore.Weight.Value); //擇優成績。

                    if (dscore.ScoreOrigin.HasValue)
                        TotalScoreOrigin += (dscore.ScoreOrigin.Value * dscore.Weight.Value); //原始成績。
                    else
                        TotalScoreOrigin += (dscore.Value.Value * dscore.Weight.Value); //將擇優當成原始。

                    TotalWeight += dscore.Weight.Value; //比重不會因為哪種成績而不同。
                }
            }

            if (TotalWeight <= 0) return new ScoreResult(); //沒有成績。

            ScoreResult sr = new ScoreResult();

            sr.Score = TotalScore / TotalWeight;
            sr.ScoreOrigin = TotalScoreOrigin / TotalWeight;

            return sr;
        }

        private class ScoreResult
        {
            /// <summary>
            /// 成績，一般是擇優後的成績，大部份的報表都會使用此成績。
            /// </summary>
            public decimal? Score { get; set; }

            /// <summary>
            /// 原始成績，最原始的成績。
            /// </summary>
            public decimal? ScoreOrigin { get; set; }
        }

        private string GetDomainSubjectText(string subjName, string subjText)
        {
            if (!string.IsNullOrEmpty(subjText))
            {
                return string.Format("({0}){1};", subjName, subjText);
            }
            else
                return string.Empty;
        }

        private bool IsValidItem(string item)
        {
            if (Filter.Count <= 0) return true;

            return (Filter.Contains(item));
        }
    }
}
