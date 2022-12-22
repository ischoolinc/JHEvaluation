using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using JHSchool.Evaluation.Calculation;

namespace HsinChuExamScore_JH.DAO
{
    /// <summary>
    /// 評量科目成績
    /// </summary>
    public class ExamSubjectScore : IScore
    {
        /// <summary>
        /// 領域名稱
        /// </summary>
        public string DomainName { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName { get; set; }

        /// <summary>
        /// 學分數
        /// </summary>
        public decimal? Credit { get; set; }

        /// <summary>
        /// 定期評量
        /// </summary>
        public decimal? ScoreF { get; set; }

        /// <summary>
        /// 平時評量
        /// </summary>
        public decimal? ScoreA { get; set; }

        /// <summary>
        /// 總成績
        /// </summary>
        public decimal? ScoreT { get; set; }

        /// <summary>
        /// 文字評量
        /// </summary>
        public string Text { get; set; }


        /// <summary>
        /// (參考試別) 總成績
        /// </summary>
        public decimal? RefScoreT { get; set; }

        /// <summary>
        /// (參考試別) 平時成績
        /// </summary>
        public decimal? RefScoreA { get; set; }

        /// <summary>
        /// (參考試别) 定期評量
        /// </summary>
        public decimal? RefScoreF { get; set; }

        /// <summary>
        /// 定期評量 百分比
        /// </summary>
        public decimal? ScorePercentage { get; set; }

        /// <summary>
        /// 平時評量 百分比
        /// </summary>
        public decimal? AssessScorePercentage { get; set; }

        /// <summary>
        /// 計算總成績(國中的評量成績=定期評量&平時評量)
        /// </summary>
        /// <param name="IRecord">評量設定</param>
        /// <param name="isReferrnceScore">是否為參考試別成績</param>
        /// <returns></returns>
        public void GetTotalScore(Boolean isReferrnceScore, JHAEIncludeRecord IRecord, Dictionary<string, decimal> ScorePercentageHSDict)
        {
            AEIncludeData aedata;
            if (IRecord == null)
            {
                aedata = new AEIncludeData();
            }
            else
            {
                aedata = new AEIncludeData(IRecord);
            }
            ScorePercentage = AssessScorePercentage = 0;
            if (IRecord != null && ScorePercentageHSDict.ContainsKey(IRecord.RefAssessmentSetupID))
            {
                ScorePercentage = ScorePercentageHSDict[IRecord.RefAssessmentSetupID] * 0.01M;
                AssessScorePercentage = (100 - ScorePercentageHSDict[IRecord.RefAssessmentSetupID]) * 0.01M;
            }

            if (!aedata.UseScore && aedata.UseAssignmentScore)
            {
                ScorePercentage = 0;
                AssessScorePercentage = 1;
            }
            if (!aedata.UseScore && !aedata.UseAssignmentScore)
            {
                ScorePercentage = 0;
                AssessScorePercentage = 0;
            }
            if (ScoreF.HasValue && Program.ScoreValueMap.ContainsKey(ScoreF.Value))
            {
                if (!Program.ScoreValueMap[ScoreF.Value].AllowCalculation)
                {
                    ScorePercentage = 0;
                    AssessScorePercentage = 1;
                }
            }
            if (ScoreA.HasValue && Program.ScoreValueMap.ContainsKey(ScoreA.Value))
            {
                if (!Program.ScoreValueMap[ScoreA.Value].AllowCalculation)
                {
                    ScorePercentage = 1;
                    AssessScorePercentage = 0;
                }
            }
            if (ScorePercentage > 0 && ScoreF.HasValue)
            {
                if (Program.ScoreValueMap.ContainsKey(ScoreF.Value))
                {
                    if (Program.ScoreValueMap[ScoreF.Value].AllowCalculation)
                    {
                        ScoreT = (ScoreT == null ? 0 : ScoreT) + Program.ScoreValueMap[ScoreF.Value].Score.Value * ScorePercentage;
                    }
                }
                else
                {
                    if (ScoreA.HasValue)
                        ScoreT = (ScoreT == null ? 0 : ScoreT) + ScoreF.Value * ScorePercentage;
                    else //有定期評量，沒有平時評量
                        ScoreT = ScoreF.Value;
                }
            }
            if (AssessScorePercentage > 0 && ScoreA.HasValue)
            {
                if (Program.ScoreValueMap.ContainsKey(ScoreA.Value))
                {
                    if (Program.ScoreValueMap[ScoreA.Value].AllowCalculation)
                    {
                        ScoreT = (ScoreT == null ? 0 : ScoreT) + Program.ScoreValueMap[ScoreA.Value].Score.Value * AssessScorePercentage;
                    }
                }
                else
                {
                    if (ScoreF.HasValue)
                        ScoreT = (ScoreT == null ? 0 : ScoreT) + ScoreA.Value * AssessScorePercentage;
                    else //有平時評量，沒有定期評量
                        ScoreT = ScoreA.Value;
                }
            }
        }
        public void GetRefTotalScore(Boolean isReferrnceScore, JHAEIncludeRecord IRecord, Dictionary<string, decimal> ScorePercentageHSDict)
        {
            AEIncludeData aedata;
            if (IRecord == null)
            {
                aedata = new AEIncludeData();
            }
            else
            {
                aedata = new AEIncludeData(IRecord);
            }
            ScorePercentage = AssessScorePercentage = 0;

            if (IRecord != null && ScorePercentageHSDict.ContainsKey(IRecord.RefAssessmentSetupID))
            {
                ScorePercentage = ScorePercentageHSDict[IRecord.RefAssessmentSetupID] * 0.01M;
                AssessScorePercentage = (100 - ScorePercentageHSDict[IRecord.RefAssessmentSetupID]) * 0.01M;
            }

            if (!aedata.UseScore && aedata.UseAssignmentScore)
            {
                ScorePercentage = 0;
                AssessScorePercentage = 1;
            }
            if (!aedata.UseScore && !aedata.UseAssignmentScore)
            {
                ScorePercentage = 0;
                AssessScorePercentage = 0;
            }
            if (ScoreF.HasValue && Program.ScoreValueMap.ContainsKey(ScoreF.Value))
            {
                if (!Program.ScoreValueMap[ScoreF.Value].AllowCalculation)
                {
                    ScorePercentage = 0;
                    AssessScorePercentage = 1;
                }
            }
            if (ScoreA.HasValue && Program.ScoreValueMap.ContainsKey(ScoreA.Value))
            {
                if (!Program.ScoreValueMap[ScoreA.Value].AllowCalculation)
                {
                    ScorePercentage = 1;
                    AssessScorePercentage = 0;
                }
            }
            if (isReferrnceScore) // 如果有含參考試別 就要計算 參考試的總成績
            {
                if (RefScoreF.HasValue && Program.ScoreValueMap.ContainsKey(RefScoreF.Value))
                {
                    if (!Program.ScoreValueMap[RefScoreF.Value].AllowCalculation)
                    {
                        ScorePercentage = 0;
                        AssessScorePercentage = 1;
                    }
                }
                if (RefScoreA.HasValue && Program.ScoreValueMap.ContainsKey(RefScoreA.Value))
                {
                    if (!Program.ScoreValueMap[RefScoreA.Value].AllowCalculation)
                    {
                        ScorePercentage = 1;
                        AssessScorePercentage = 0;
                    }
                }
                if (ScorePercentage > 0 && RefScoreF.HasValue)
                {
                    if (Program.ScoreValueMap.ContainsKey(RefScoreF.Value))
                    {
                        if (Program.ScoreValueMap[RefScoreF.Value].AllowCalculation)
                        {
                            RefScoreT = (RefScoreT == null ? 0 : RefScoreT) + Program.ScoreValueMap[RefScoreF.Value].Score.Value * ScorePercentage;
                        }
                    }
                    else
                    {
                        RefScoreT = (RefScoreT == null ? 0 : RefScoreT) + RefScoreF.Value * ScorePercentage;
                    }
                }
                if (AssessScorePercentage > 0 && RefScoreA.HasValue)
                {
                    if (Program.ScoreValueMap.ContainsKey(RefScoreA.Value))
                    {
                        if (Program.ScoreValueMap[RefScoreA.Value].AllowCalculation)
                        {
                            RefScoreT = (RefScoreT == null ? 0 : RefScoreT) + Program.ScoreValueMap[RefScoreA.Value].Score.Value * AssessScorePercentage;
                        }
                    }
                    else
                    {
                        RefScoreT = (RefScoreT == null ? 0 : RefScoreT) + RefScoreA.Value * AssessScorePercentage;
                    }
                }
            }
        }
    }
}
