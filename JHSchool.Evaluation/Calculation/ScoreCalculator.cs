using System;
using System.Xml;
using JHSchool.Data;

namespace JHSchool.Evaluation.Calculation
{
    public class ScoreCalculator : JHSchool.Evaluation.ScoreCalculator
    {
        public ScoreCalculator(JHScoreCalcRuleRecord record) : base(record)
        {
        }
    }
}
