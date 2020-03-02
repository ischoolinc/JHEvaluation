using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JHSchool.Evaluation.Calculation;

namespace HsinChuSemesterClassFixedRank.DAO
{
    public class ScoreItemC
    {
        public string ItemName { get; set; }

        // 總分
        List<decimal> ScoreList = new List<decimal>();
        // 加權總分
        List<decimal> ScoreListA = new List<decimal>();
        // 學分數
        List<decimal> CreditList = new List<decimal>();

        // 原始成績總分
        List<decimal> ScoreOriginList = new List<decimal>();
        // 原始成績加權總分
        List<decimal> ScoreOriginListA = new List<decimal>();


        /// <summary>
        /// 總分
        /// </summary>
        public decimal SumScore = 0;
        /// <summary>
        /// 加權總分
        /// </summary>
        public decimal SumScoreA = 0;
        /// <summary>
        /// 平均
        /// </summary>
        public decimal AvgScore = 0;
        /// <summary>
        /// 加權平均
        /// </summary>
        public decimal AvgScoreA = 0;

        /// <summary>
        /// 總分(原始)
        /// </summary>
        public decimal SumScoreOrigin = 0;
        /// <summary>
        /// 加權總分(原始)
        /// </summary>
        public decimal SumScoreOriginA = 0;
        /// <summary>
        /// 平均(原始)
        /// </summary>
        public decimal AvgScoreOrigin = 0;
        /// <summary>
        /// 加權平均(原始)
        /// </summary>
        public decimal AvgScoreOriginA = 0;

        public void AddScore(decimal score)
        {
            ScoreList.Add(score);
        }

        public void AddOriginScore(decimal score)
        {
            ScoreOriginList.Add(score);
        }

        public void AddOriginScore(decimal score, decimal credit)
        {
            ScoreOriginListA.Add(score * credit);
        }

        public void AddScore(decimal score, decimal credit)
        {
            ScoreListA.Add(score * credit);
        }

        public void AddCredit(decimal credit)
        {
            CreditList.Add(credit);
        }

        /// <summary>
        /// 計算總計,傳入平均四捨五入數
        /// </summary>
        public void CalScore(ScoreCalculator cal,string scoreType)
        {
            // 總分
            SumScore = ScoreList.Sum();

            // 加權總分
            SumScoreA = ScoreListA.Sum();

            int cCot = CreditList.Count();
            decimal sCot = CreditList.Sum();

            if (scoreType =="subject")
            {
                // 平均
                if (cCot > 0)
                    AvgScore = cal.ParseSubjectScore(SumScore / cCot);

                // 加權平均
                if (sCot > 0)
                    AvgScoreA = cal.ParseSubjectScore(SumScoreA / sCot);
            }

            if (scoreType == "domain")
            {
                // 平均
                if (cCot > 0)
                    AvgScore = cal.ParseDomainScore(SumScore / cCot);

                // 加權平均
                if (sCot > 0)
                    AvgScoreA = cal.ParseDomainScore(SumScoreA / sCot);
            }

           
            // 原始成績
            // 總分
            SumScoreOrigin = ScoreOriginList.Sum();

            // 加權總分
            SumScoreOriginA = ScoreOriginListA.Sum();


            if (scoreType == "subject")
            {
                // 平均
                if (cCot > 0)
                    AvgScoreOrigin = cal.ParseSubjectScore(SumScoreOrigin / cCot);

                // 加權平均
                if (sCot > 0)
                    AvgScoreOriginA = cal.ParseSubjectScore(SumScoreOriginA / sCot);
            }

            if (scoreType == "domain")
            {
                // 平均
                if (cCot > 0)
                    AvgScoreOrigin = cal.ParseDomainScore(SumScoreOrigin / cCot);

                // 加權平均
                if (sCot > 0)
                    AvgScoreOriginA = cal.ParseDomainScore(SumScoreOriginA / sCot);
            }

             
        }
    }
}
