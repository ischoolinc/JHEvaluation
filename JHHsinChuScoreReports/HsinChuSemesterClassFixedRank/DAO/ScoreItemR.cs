using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JHSchool.Evaluation.Calculation;

namespace HsinChuSemesterClassFixedRank.DAO
{
    /// <summary>
    /// 統計用
    /// </summary>
    public class ScoreItemR
    {
        public string ScoreKey { get; set; }

        public string ScoreOriginKey { get; set; }

        public string PassCountKey { get; set; }

        public string PassOriginKey { get; set; }

        // 成績
        List<decimal> ScoreList = new List<decimal>();

        // 原始成績
        List<decimal> ScoreOriginList = new List<decimal>();


        public void AddScore(decimal score)
        {
            ScoreList.Add(score);
        }

        /// <summary>
        /// 原始成績
        /// </summary>
        /// <param name="score"></param>
        public void AddScoreOrigin(decimal score)
        {
            ScoreOriginList.Add(score);
        }

        /// <summary>
        /// 及格人數
        /// </summary>
        /// <returns></returns>
        public int GetPassScoreCount()
        {
            int value = 0;
            if (ScoreList.Count > 0)
                value = ScoreList.Where(x => x >= 60).Count();
            return value;
        }

        /// <summary>
        /// 原始成績及格人數
        /// </summary>
        /// <returns></returns>
        public int GetPassScoreOriginCount()
        {
            int value = 0;
            if (ScoreOriginList.Count > 0)
                value = ScoreOriginList.Where(x => x >= 60).Count();
            return value;
        }

        /// <summary>
        /// 取得平均,傳入第幾位四捨五入
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        public decimal GetAvgScore(ScoreCalculator cal, string scoreType)
        {
            decimal value = 0;

            if (ScoreList.Count > 0)
            {
                if (scoreType == "subject")
                {
                    value = cal.ParseSubjectScore(ScoreList.Average());
                }

                if (scoreType == "domain")
                {
                    value = cal.ParseDomainScore(ScoreList.Average());
                }
            }
            return value;
        }

        /// <summary>
        /// 原始分數取得平均,傳入第幾位四捨五入
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        public decimal GetAvgScoreOrigin(ScoreCalculator cal, string scoreType)
        {
            decimal value = 0;
            if (ScoreOriginList.Count > 0)
            {
                if (scoreType == "subject")
                {
                    value = cal.ParseSubjectScore(ScoreOriginList.Average());
                }

                if (scoreType == "domain")
                {
                    value = cal.ParseDomainScore(ScoreOriginList.Average());
                }

            }
            return value;
        }
    }
}
