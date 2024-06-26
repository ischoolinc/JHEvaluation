﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Evaluation.Calculation;

namespace HsinChuExamScore_JH.DAO
{
    /// <summary>
    /// 學生評量成績
    /// </summary>
    public class StudExamScore
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public int GradeYear { get; set; }

        /// <summary>
        /// 班級ID
        /// </summary>
        public string ClassID { get; set; }

        /// <summary>
        /// 試別名稱
        /// </summary>
        public string ExamName { get; set; }

        /// <summary>
        /// 領域成績
        /// </summary>
        public Dictionary<string, IScore> _ExamDomainScoreDict = new Dictionary<string, IScore>();

        /// <summary>
        /// 科目成績
        /// </summary>
        public Dictionary<string, IScore> _ExamSubjectScoreDict = new Dictionary<string, IScore>();

        /// <summary>
        /// 參考科目成績
        /// </summary>
        public Dictionary<string, IScore> _RefExamSubjectScoreDict = new Dictionary<string, IScore>();



        /// <summary>
        /// 成績計算規則
        /// </summary>
        ScoreCalculator _Calculator;

        public StudExamScore(ScoreCalculator studentCalculator)
        {
            _Calculator = studentCalculator;
        }

        public void CalcSubjectToDomain(string ifRefScore)
        {

            Dictionary<string, IScore> scoreDictionary = new Dictionary<string, IScore>();

            Dictionary<string, List<decimal>> scDict = new Dictionary<string, List<decimal>>();

            List<string> DomainNameList = new List<string>();
            string ss = "_加權", sc = "_學分", sc_f = "_定期學分", sc_a = "_平時學分", sa = "_平時加權", sf = "_定期加權", scs = "_學分值";

            foreach (ExamSubjectScore ess in this.GetSubjScoreDictionary(ifRefScore).Values)
            {
                string keys = ess.DomainName + ss;
                string keyc = ess.DomainName + sc;
                string keyc_f = ess.DomainName + sc_f;
                string keyc_a = ess.DomainName + sc_a;
                string keya = ess.DomainName + sa;
                string keyf = ess.DomainName + sf;
                string keycs = ess.DomainName + scs;

                if (!DomainNameList.Contains(ess.DomainName))
                    DomainNameList.Add(ess.DomainName);

                if (!scDict.ContainsKey(keys)) scDict.Add(keys, new List<decimal>());
                if (!scDict.ContainsKey(keyc)) scDict.Add(keyc, new List<decimal>());
                if (!scDict.ContainsKey(keyc_f)) scDict.Add(keyc_f, new List<decimal>());
                if (!scDict.ContainsKey(keyc_a)) scDict.Add(keyc_a, new List<decimal>());
                if (!scDict.ContainsKey(keya)) scDict.Add(keya, new List<decimal>());
                if (!scDict.ContainsKey(keyf)) scDict.Add(keyf, new List<decimal>());
                if (!scDict.ContainsKey(keycs)) scDict.Add(keycs, new List<decimal>());

                // 總分加權平均
                if (ess.ScoreT.HasValue && ess.Credit.HasValue)
                    scDict[keys].Add(_Calculator.ParseSubjectScore(ess.ScoreT.Value * ess.Credit.Value)); // 成績計算規則 的處理

                // 定期加權平均
                if (ess.ScoreF.HasValue && ess.Credit.HasValue)
                {
                    if (Program.ScoreValueMap.ContainsKey(ess.ScoreF.Value))
                    {
                        if (Program.ScoreValueMap[ess.ScoreF.Value].AllowCalculation)
                        {
                            scDict[keyf].Add(_Calculator.ParseSubjectScore(Program.ScoreValueMap[ess.ScoreF.Value].Score.Value * ess.Credit.Value)); // 抓成績計算規則
                            scDict[keyc_f].Add(ess.Credit.Value);
                        }
                    } 
                    else
                    {
                        scDict[keyf].Add(_Calculator.ParseSubjectScore(ess.ScoreF.Value * ess.Credit.Value)); // 抓成績計算規則
                        scDict[keyc_f].Add(ess.Credit.Value);
                    }
                }
                // 平時加權平均
                if (ess.ScoreA.HasValue && ess.Credit.HasValue)
                {
                    if (Program.ScoreValueMap.ContainsKey(ess.ScoreA.Value))
                    {
                        if (Program.ScoreValueMap[ess.ScoreA.Value].AllowCalculation)
                        {
                            scDict[keya].Add(_Calculator.ParseSubjectScore(Program.ScoreValueMap[ess.ScoreA.Value].Score.Value * ess.Credit.Value)); // 抓成績計算規則
                            scDict[keyc_a].Add(ess.Credit.Value);
                        }
                    }
                    else
                    {
                        scDict[keya].Add(_Calculator.ParseSubjectScore(ess.ScoreA.Value * ess.Credit.Value)); // 抓成績計算規則
                        scDict[keyc_a].Add(ess.Credit.Value);
                    }
                }

                // 學分
                if (ess.Credit.HasValue)
                    scDict[keyc].Add(ess.Credit.Value);

                // 有成績的學分數
                if (ess.Credit.HasValue && ess.ScoreT.HasValue)
                    scDict[keycs].Add(ess.Credit.Value);
            }


            foreach (string name in DomainNameList)
            {
                //string ss = "_加權", sc = "_學分", sa = "_平時加權", sf = "_定期加權", scs = "_學分值";
                ExamDomainScore eds = new ExamDomainScore();
                // 處理彈性課程
                if (string.IsNullOrEmpty(name))
                    eds.DomainName = "彈性課程";
                else
                    eds.DomainName = name;

                string keyc = name + sc;
                string keyc_f = name + sc_f;
                string keyc_a = name + sc_a;
                string keycs = name + scs;
                // 定期學分加總
                if (scDict.ContainsKey(keyc_f))
                    eds.CreditF = scDict[keyc_f].Sum();
                // 平時學分加總
                if (scDict.ContainsKey(keyc_a))
                    eds.CreditA = scDict[keyc_a].Sum();
                // 學分加總
                if (scDict.ContainsKey(keyc))
                    eds.Credit = scDict[keyc].Sum();

                // 有成績學分
                if (scDict.ContainsKey(keycs))
                    eds.Credit1 = scDict[keycs].Sum();

                string keyss = name + ss;
                string keysf = name + sf;
                string keysa = name + sa;
                // 領域總成績加權平均
                if (eds.Credit1.HasValue)
                    if (eds.Credit1.Value > 0)
                        if (scDict.ContainsKey(keyss))
                            eds.ScoreT = _Calculator.ParseDomainScore(scDict[keyss].Sum() / eds.Credit1.Value);

                // 領域定期成績加權平均
                if (eds.CreditF.HasValue)
                    if (eds.CreditF.Value > 0)
                        if (scDict.ContainsKey(keysf))
                            eds.ScoreF = _Calculator.ParseDomainScore(scDict[keysf].Sum() / eds.CreditF.Value);

                // 領域平時成績加權平均
                if (eds.CreditA.HasValue)
                    if (eds.CreditA.Value > 0)
                        if (scDict.ContainsKey(keysa))
                            eds.ScoreA = _Calculator.ParseDomainScore(scDict[keysa].Sum() / eds.CreditA.Value);

                _ExamDomainScoreDict.Add(name, eds);
            }

        }

        /// <summary>
        /// 評量領域總成績加權平均
        /// </summary>
        /// <returns></returns>
        public decimal? GetDomainWAvgScoreA(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            decimal cc = 0;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue && sc.Credit1.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value * sc.Credit1.Value;
                        cc += sc.Credit1.Value;
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue && sc.Credit1.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value * sc.Credit1.Value;
                        cc += sc.Credit1.Value;
                    }
                }
            }

            if (cc > 0)
                score = _Calculator.ParseDomainScore((ss.HasValue ? ss.Value : 0) / cc);

            return score;
        }


        /// <summary>
        /// 評量領域定期加權平均
        /// </summary>
        /// <param name="all">是否排除彈性課程</param>
        /// <returns></returns>
        public decimal? GetDomainWAvgScoreF(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            decimal cc = 0;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue && sc.CreditF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.CreditF.Value;
                                cc++;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.CreditF.Value;
                            cc += sc.CreditF.Value;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue && sc.CreditF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.CreditF.Value;
                                cc++;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.CreditF.Value;
                            cc += sc.CreditF.Value;
                        }
                    }
                }
            }

            if (cc > 0)
                score = _Calculator.ParseDomainScore((ss.HasValue ? ss.Value : 0) / cc);

            return score;
        }


        /// <summary>
        /// 取得定期評量成績(評量加定期)之算術平均 從科目算上去
        /// </summary>
        /// <param name="all"> 是否包含彈性課程</param>
        /// <returns></returns>
        public decimal? GetDomainArithmeticMeanScoreA(bool all)
        {
            decimal? arithmeticMeanScoreA = null;
            decimal? totalScore=null;
            decimal subjCount = 0;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {

                    if (sc.ScoreA.HasValue && sc.CreditA.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreA.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreA.Value].AllowCalculation)
                            {
                                totalScore = (totalScore.HasValue ? totalScore.Value : 0) + Program.ScoreValueMap[sc.ScoreA.Value].Score.Value * sc.CreditA.Value;
                                subjCount++;
                            }
                        }
                        else
                        {
                            totalScore = (totalScore.HasValue ? totalScore.Value : 0) + sc.ScoreA.Value * sc.CreditA.Value;
                            subjCount++;
                        }
                    }
                }
            }
            else // 不計算彈性課程
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;
                    // 使用有成績計算加權
                    if (sc.ScoreA.HasValue && sc.CreditA.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreA.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreA.Value].AllowCalculation)
                            {
                                totalScore = (totalScore.HasValue ? totalScore.Value : 0) + Program.ScoreValueMap[sc.ScoreA.Value].Score.Value * sc.CreditA.Value;
                                subjCount++;
                            }
                        }
                        else
                        {
                            totalScore = (totalScore.HasValue ? totalScore.Value : 0) + sc.ScoreA.Value * sc.CreditA.Value;
                            subjCount++;
                        }
                    }
                }
            }

            if (subjCount > 0)
                arithmeticMeanScoreA = _Calculator.ParseDomainScore((totalScore.HasValue ? totalScore.Value : 0) / subjCount);

            return arithmeticMeanScoreA;
        }



        public decimal? GetDomainArithmeticMeanScoreF(bool all)
        {
            decimal? arithmeticMeanScoreF = null;
            decimal? totalScore = null;
            decimal subjCount = 0;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    if (sc.ScoreF.HasValue && sc.CreditF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                totalScore = (totalScore.HasValue ? totalScore.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.CreditF.Value;
                                subjCount++;
                            }
                        }
                        else
                        {
                            totalScore = (totalScore.HasValue ? totalScore.Value : 0) + sc.ScoreF.Value * sc.CreditF.Value;
                            subjCount++;
                        }
                    }
                }
            }
            else // 不計算彈性課程
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;
                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue && sc.CreditF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                totalScore = (totalScore.HasValue ? totalScore.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.CreditF.Value;
                                subjCount++;
                            }
                        }
                        else
                        {
                            totalScore = (totalScore.HasValue ? totalScore.Value : 0) + sc.ScoreF.Value * sc.CreditF.Value;
                            subjCount++;
                        }
                    }
                }
            }

            if (subjCount > 0)
                arithmeticMeanScoreF = _Calculator.ParseDomainScore((totalScore.HasValue ? totalScore.Value : 0) / subjCount);

            return arithmeticMeanScoreF;
        }


        /// <summary>
        /// 評量領域總成績算數平均
        /// </summary>
        /// <returns></returns>
        public decimal? GetDomainScore_A(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            decimal cc = 0;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value;
                        cc += 1;
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value;
                        cc += 1;
                    }
                }
            }

            if (cc > 0)
                score = _Calculator.ParseDomainScore((ss.HasValue ? ss.Value : 0) / cc);

            return score;
        }

        /// <summary>
        /// 評量領域定期算數平均
        /// </summary>
        /// <returns></returns>
        public decimal? GetDomainScore_F(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            decimal cc = 0;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value;
                                cc += 1;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value;
                            cc += 1;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value;
                                cc += 1;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value;
                            cc += 1;
                        }
                    }
                }
            }

            if (cc > 0)
                score = _Calculator.ParseDomainScore((ss.HasValue ? ss.Value : 0) / cc);

            return score;
        }



        /// <summary>
        /// 評量科目定期評量加權平均
        /// </summary>
        /// <returns></returns>
        public decimal? GetSubjectScoreAF(bool all, string isRefScore)
        {
            decimal? score = null;
            decimal? ss = null;
            decimal cc = 0;

            if (all)
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    if (sc.ScoreF.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.Credit.Value;
                                cc += sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.Credit.Value;
                            cc += sc.Credit.Value;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    if (sc.ScoreF.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.Credit.Value;
                                cc += sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.Credit.Value;
                            cc += sc.Credit.Value;
                        }
                    }
                }
            }

            if (cc > 0)
                score = _Calculator.ParseSubjectScore((ss.HasValue ? ss.Value : 0) / cc);

            return score;
        }

        /// <summary>
        /// 評量科目平時評量加權平均
        /// </summary>
        /// <returns></returns>
        public decimal? GetSubjectScoreAA(bool all, string isRefScore)
        {
            decimal? score = null;
            decimal? ss = null;
            decimal cc = 0;

            if (all)
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    if (sc.ScoreA.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreA.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreA.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreA.Value].Score.Value * sc.Credit.Value;
                                cc += sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreA.Value * sc.Credit.Value;
                            cc += sc.Credit.Value;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    if (sc.ScoreA.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreA.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreA.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreA.Value].Score.Value * sc.Credit.Value;
                                cc += sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreA.Value * sc.Credit.Value;
                            cc += sc.Credit.Value;
                        }
                    }
                }
            }

            if (cc > 0)
                score = _Calculator.ParseSubjectScore((ss.HasValue ? ss.Value : 0) / cc);

            return score;
        }

        /// <summary>
        /// 評量科目總成績加權平均
        /// </summary>
        /// <returns></returns>
        public decimal? GetSubjectScoreAT(bool all, string isRefScore)
        {
            decimal? score = null;
            decimal? ss = null;
            decimal cc = 0;

            if (all)
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    if (sc.ScoreT.HasValue && sc.Credit.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value * sc.Credit.Value;
                        cc += sc.Credit.Value;
                    }
                }
            }
            else
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    if (sc.ScoreT.HasValue && sc.Credit.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value * sc.Credit.Value;
                        cc += sc.Credit.Value;
                    }
                }
            }

            if (cc > 0)
                score = _Calculator.ParseSubjectScore((ss.HasValue ? ss.Value : 0) / cc);

            return score;
        }


        /// <summary>
        /// 取的算術平均
        /// </summary>
        /// <param name="isContainFlex">是否包含彈性課程</param>
        /// <param name="scoreType">領域成績 OR 科目成績</param>
        /// <param name="enumScoreComposition">定期、平時、定期加平時</param>
        /// <returns></returns>
        public decimal? GetScoreArithmeticＭean(bool isContainFlex, EnumScoreType scoreType, EnumScoreComposition enumScoreComposition)
        {
            Dictionary<string, IScore> dicScoreInfo = new Dictionary<string, IScore>();
            decimal? result = null;
         
                if (scoreType == EnumScoreType.領域)
                {
                    dicScoreInfo = _ExamDomainScoreDict;
                }
                else if (scoreType == EnumScoreType.科目)
                {
                    dicScoreInfo = _ExamSubjectScoreDict;
                }
                else if (scoreType == EnumScoreType.參考科目)
                {

                    dicScoreInfo = this._RefExamSubjectScoreDict;
                }

                decimal? totalScore = null;
                decimal SubjCount = 0;

                if (isContainFlex)
                {
                    foreach (IScore sc in dicScoreInfo.Values)
                    {
                        if (this.GetScore(sc, enumScoreComposition).HasValue && sc.Credit.HasValue)
                        {
                            totalScore = (totalScore.HasValue ? totalScore.Value : 0) + GetScore(sc, enumScoreComposition).Value;
                            SubjCount++;
                        }
                    }
                }
                else
                {
                    foreach (IScore sc in dicScoreInfo.Values)
                    {
                        // 過濾彈性課程
                        if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                            continue;

                        if (this.GetScore(sc, enumScoreComposition).HasValue && sc.Credit.HasValue)
                        {
                        totalScore = (totalScore.HasValue ? totalScore.Value : 0) + this.GetScore(sc, enumScoreComposition).Value;
                            SubjCount++;
                        }
                    }
                }

                if (SubjCount > 0)
                    result = _Calculator.ParseSubjectScore((totalScore.HasValue ? totalScore.Value : 0) / SubjCount);

          

            // 確認 是哪一種成績要計算 取用不同dictionary

         
            return result;
        }




        /// <summary>
        /// 取的算術總分
        /// </summary>
        /// <param name="isContainFlex">是否包含彈性課程</param>
        /// <param name="scoreType">領域成績 OR 科目成績</param>
        /// <param name="enumScoreComposition">定期、平時、定期加平時</param>
        /// <returns></returns>
        public decimal? GetScoreArithmeticTotal(bool isContainFlex, EnumScoreType scoreType, EnumScoreComposition enumScoreComposition)
        {
            Dictionary<string, IScore> dicScoreInfo = new Dictionary<string, IScore>();

            // 確認 是哪一種成績要計算 取用不同dictionary
            if (scoreType == EnumScoreType.領域)
            {
                dicScoreInfo = _ExamDomainScoreDict;
            }
            else if (scoreType == EnumScoreType.科目)
            {
                dicScoreInfo = _ExamSubjectScoreDict;
            }
            else if (scoreType == EnumScoreType.參考科目)
            {
                dicScoreInfo = _RefExamSubjectScoreDict;

            }


            decimal? result = null;

            decimal? totalScore = null;
            //if (dicScoreInfo.Count > 0) //如果有成績紀錄(至少有一科)
            //{
            //    totalScore = 0;
            //}

            if (isContainFlex)
            {
                foreach (IScore sc in dicScoreInfo.Values)
                {
                    if (this.GetScore(sc, enumScoreComposition).HasValue && sc.Credit.HasValue)
                    {
                        totalScore = (totalScore.HasValue ? totalScore.Value : 0) + GetScore(sc, enumScoreComposition).Value;
                    }
                }
            }
            else
            {
                foreach (IScore sc in dicScoreInfo.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    if (this.GetScore(sc, enumScoreComposition).HasValue && sc.Credit.HasValue)
                    {
                        totalScore = (totalScore.HasValue ? totalScore.Value : 0) + this.GetScore(sc, enumScoreComposition).Value;

                    }
                }
            }

            result = totalScore;

            return result;
        }






        /// <summary>
        /// 取的Score 看是要用哪一種 回傳不同
        /// </summary>
        /// <returns>定期加平時 or 定期 or 平時</returns>
        internal decimal? GetScore(IScore socreInfo, EnumScoreComposition EnumScoreComposition)
        {
            if (EnumScoreComposition == EnumScoreComposition.成績)
            {
                if (socreInfo.ScoreT.HasValue && Program.ScoreValueMap.ContainsKey(socreInfo.ScoreT.Value))
                {
                    return Program.ScoreValueMap[socreInfo.ScoreT.Value].Score;
                }
                return socreInfo.ScoreT;
            }
            else if (EnumScoreComposition == EnumScoreComposition.定期成績)
            {
                if (socreInfo.ScoreF.HasValue && Program.ScoreValueMap.ContainsKey(socreInfo.ScoreF.Value))
                {
                    return Program.ScoreValueMap[socreInfo.ScoreF.Value].Score;
                }
                return socreInfo.ScoreF;
            }
            else if (EnumScoreComposition == EnumScoreComposition.平時成績)  // 因目前沒有平時成績排名需求 應該用不到
            {
                if (socreInfo.ScoreA.HasValue && Program.ScoreValueMap.ContainsKey(socreInfo.ScoreA.Value))
                {
                    return Program.ScoreValueMap[socreInfo.ScoreA.Value].Score;
                }
                return socreInfo.ScoreA;
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// 評量領域總成績加權總分
        /// </summary>
        /// <returns></returns>
        public decimal? GetDomainScoreS(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue && sc.Credit1.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value * sc.Credit1.Value;
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue && sc.Credit1.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value * sc.Credit1.Value;
                    }
                }
            }

            score = ss;

            return score;
        }

        /// <summary>
        /// 評量領域定期總成績加權總分
        /// </summary>
        /// <returns></returns>
        public decimal? GetDomainScoreSF(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue && sc.CreditF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.CreditF.Value;
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.CreditF.Value;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue && sc.CreditF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.CreditF.Value;
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.CreditF.Value;
                        }
                    }
                }
            }

            score = ss;

            return score;
        }


        /// <summary>
        /// 評量領域總成績加權總分
        /// </summary>
        /// <returns></returns>
        public decimal? GetDomainScore_S(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value;
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreT.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreT.Value;
                    }
                }
            }

            score = ss;

            return score;
        }

        /// <summary>
        /// 評量領域定期總成績加權總分
        /// </summary>
        /// <returns></returns>
        public decimal? GetDomainScore_SF(bool all)
        {
            decimal? score = null;
            decimal? ss = null;
            if (all)
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamDomainScore sc in _ExamDomainScoreDict.Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    // 使用有成績計算加權
                    if (sc.ScoreF.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value;
                        }
                    }
                }
            }

            score = ss;

            return score;
        }



        /// <summary>
        /// 評量科目定期評量加權總分
        /// </summary>
        /// <returns></returns>
        public decimal? GetSubjectScoreSF(bool all, string isRefScore)
        {
            decimal? score = null;
            decimal? ss = null;

            if (all)
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    if (sc.ScoreF.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.Credit.Value;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    if (sc.ScoreF.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreF.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreF.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreF.Value].Score.Value * sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreF.Value * sc.Credit.Value;
                        }
                    }
                }
            }

            score = ss;

            return score;
        }

        /// <summary>
        /// 評量科目平時評量加權總分
        /// </summary>
        /// <returns></returns>
        public decimal? GetSubjectScoreSA(bool all, string isRefScore)
        {
            decimal? score = null;
            decimal? ss = null;

            if (all)
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    if (sc.ScoreA.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreA.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreA.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreA.Value].Score.Value * sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreA.Value * sc.Credit.Value;
                        }
                    }
                }
            }
            else
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    if (sc.ScoreA.HasValue && sc.Credit.HasValue)
                    {
                        if (Program.ScoreValueMap.ContainsKey(sc.ScoreA.Value))
                        {
                            if (Program.ScoreValueMap[sc.ScoreA.Value].AllowCalculation)
                            {
                                ss = (ss.HasValue ? ss.Value : 0) + Program.ScoreValueMap[sc.ScoreA.Value].Score.Value * sc.Credit.Value;
                            }
                        }
                        else
                        {
                            ss = (ss.HasValue ? ss.Value : 0) + sc.ScoreA.Value * sc.Credit.Value;
                        }
                    }
                }
            }

            score = ss;

            return score;
        }

        /// <summary>
        /// 評量科目總成績加權總分
        /// </summary>
        /// <returns></returns>
        public decimal? GetSubjectScoreST(bool all, string isRefScore)
        {
            decimal? score = null;
            decimal? ss = null;

            if (all)
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    if (sc.ScoreT.HasValue && sc.Credit.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + _Calculator.ParseSubjectScore(sc.ScoreT.Value * sc.Credit.Value);
                    }
                }
            }
            else
            {
                foreach (ExamSubjectScore sc in this.GetSubjScoreDictionary(isRefScore).Values)
                {
                    // 過濾彈性課程
                    if (string.IsNullOrEmpty(sc.DomainName) || sc.DomainName == "彈性課程")
                        continue;

                    if (sc.ScoreT.HasValue && sc.Credit.HasValue)
                    {
                        ss = (ss.HasValue ? ss.Value : 0) + _Calculator.ParseSubjectScore(sc.ScoreT.Value * sc.Credit.Value);
                    }
                }
            }

            score = ss;

            return score;
        }

        /// <summary>
        /// 取得成績Dictionary  因為有要印的當次識別的 跟參考識別
        /// </summary>
        /// <param name="ifRefScore"></param>
        /// <returns></returns>
        public Dictionary<string, IScore> GetSubjScoreDictionary(string ifRefScore)
        {
            if (ifRefScore == "參考")
            {
                return this._RefExamSubjectScoreDict;
            }
            else
            {

                return this._ExamSubjectScoreDict;
            }
        }
    }
}
