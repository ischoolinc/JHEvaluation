using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HsinChuExamScoreClassFixedRank.DAO
{
    public class StudentInfo
    {
        public string StudentID { get; set; }
        public string Name { get; set; }
        public string SeatNo { get; set; }

        public string ClassID { get; set; }

        // 領域與科目
        public List<DomainInfo> DomainInfoList = new List<DomainInfo>();

        /// <summary>
        /// 總分
        /// </summary>
        public decimal? SumScore { get; set; }

        /// <summary>
        /// 加權總分
        /// </summary>
        public decimal? SumScoreA { get; set; }

        /// <summary>
        /// 平均
        /// </summary>
        public decimal? AvgScore { get; set; }

        /// <summary>
        /// 加權平均
        /// </summary>
        public decimal? AvgScoreA { get; set; }

        /// <summary>
        /// 平均班排名
        /// </summary>
        public int? ClassAvgRank { get; set; }

        /// <summary>
        /// 加權平均班排名
        /// </summary>
        /// 
        public int? ClassAvgRankA { get; set; }

        /// <summary>
        /// 平均年排名
        /// </summary>
        public int? YearAvgRank { get; set; }

        /// <summary>
        /// 加權平均年排名
        /// </summary>
        /// 
        public int? YearAvgRankA { get; set; }

        /// <summary>
        /// 總分班排名
        /// </summary>
        public int? ClassSumRank { get; set; }

        /// <summary>
        /// 加權總分班排名
        /// </summary>
        /// 
        public int? ClassSumRankA { get; set; }

        /// <summary>
        /// 總分年排名
        /// </summary>
        public int? YearSumRank { get; set; }

        /// <summary>
        /// 加權總分年排名
        /// </summary>
        /// 
        public int? YearSumRankA { get; set; }


        /// <summary>
        /// 計算成績
        /// </summary>
        public void CalScore()
        {
            // 初始
            SumScore = SumScoreA = AvgScore = AvgScoreA = 0;

            decimal cot = 0; decimal credit = 0;
            foreach (DomainInfo domain in DomainInfoList)
            {
                if (domain.Score.HasValue && domain.Credit.HasValue)
                {
                    SumScore += domain.Score.Value;
                    SumScoreA += domain.Score.Value * domain.Credit.Value;

                    cot += 1;
                    credit += domain.Credit.Value;
                }
            }
            if (cot > 0)
                AvgScore = SumScore / cot;

            if (credit > 0)
                AvgScoreA = SumScoreA / credit;

        }       
    }
}
