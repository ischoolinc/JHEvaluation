using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using HsinChu.JHEvaluation.Data;

namespace HsinChu.ClassExamScoreAvgComparison.Model
{
    internal class CourseScore
    {
        public string CourseID { get; private set; }
        public decimal? Score { get; private set; }

        private decimal? _score;
        private decimal? _assignment_score;

        public CourseScore(string courseID, decimal? score, decimal? assignmentScore)
        {
            _score = score;
            _assignment_score = assignmentScore;
            CourseID = courseID;
        }

        public void CalculateScore(HC.JHAEIncludeRecord ae, string type, Dictionary<decimal, DAL.ScoreMap> scoreMapDic)
        {
            //if (ae.UseScore) Score = _score;

            if (type == "定期")
            {
                if (ae.UseScore)
                {
                    if (_score.HasValue)
                    {
                        if (scoreMapDic.ContainsKey(_score.Value))
                        {
                            if (scoreMapDic[_score.Value].AllowCalculation) //可以被計算 ex缺考0分
                            {
                                Score = scoreMapDic[_score.Value].Score.Value;
                            }
                            else  //不能被計算 ex 免試
                            {

                            }

                        }
                        else
                        {
                            Score = _score;
                        }
                    }
                    else
                    {
                        Score = _score;
                    }
                }
            }
            else if (type == "定期加平時")
            {
                if (ae.UseScore && ae.UseAssignmentScore)
                {
                    //對照後的 定期評量分數
                    decimal? score = null;

                    //對照後的 平時評量分數
                    decimal? assignmentScore = null;

                    //找對照
                    if (_score.HasValue)
                    {
                        if (scoreMapDic.ContainsKey(_score.Value))
                        {
                            if (scoreMapDic[_score.Value].AllowCalculation) //可以被計算 ex缺考0分
                                score = scoreMapDic[_score.Value].Score.Value;
                            else  //不能被計算 ex 免試
                            {
                                score = null;
                            }
                        }
                        else
                        {
                            score = _score.Value;
                        }
                    }
                    if (_assignment_score.HasValue)
                    {
                        if (scoreMapDic.ContainsKey(_assignment_score.Value))
                        {
                            if (scoreMapDic[_assignment_score.Value].AllowCalculation) //可以被計算 ex缺考0分
                                assignmentScore = scoreMapDic[_assignment_score.Value].Score.Value;
                            else  //不能被計算 ex 免試
                            {
                                assignmentScore = null;
                            }
                        }
                        else
                        {
                            assignmentScore = _assignment_score.Value;
                        }
                    }

                    //
                    if (score != null && assignmentScore != null)
                    {
                        if (Global.ScorePercentageHSDict.ContainsKey(ae.RefAssessmentSetupID))
                        {
                            decimal ff = Global.ScorePercentageHSDict[ae.RefAssessmentSetupID];
                            decimal f = score.Value * ff * 0.01M;
                            decimal a = assignmentScore.Value * (100 - ff) * 0.01M;

                            Score = f + a;
                        }
                        else
                            Score = _score.Value * 0.5M + _assignment_score.Value * 0.5M;


                    }
                    else if (score != null)
                        Score = score;
                    else if (assignmentScore != null)
                        Score = assignmentScore;
                }
                else if (ae.UseScore)
                    Score = _score;
                else if (ae.UseAssignmentScore)
                    Score = _assignment_score;
            }
        }
    }
}
