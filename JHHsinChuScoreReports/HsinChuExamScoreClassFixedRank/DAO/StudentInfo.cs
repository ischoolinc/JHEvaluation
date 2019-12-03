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
        /// 總分(定期)
        /// </summary>
        public decimal? SumScoreF { get; set; }

        /// <summary>
        /// 加權總分(定期)
        /// </summary>
        public decimal? SumScoreAF { get; set; }

        /// <summary>
        /// 平均(定期)
        /// </summary>
        public decimal? AvgScoreF { get; set; }

        /// <summary>
        /// 加權平均(定期)
        /// </summary>
        public decimal? AvgScoreAF { get; set; }

        

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
        /// 平均班排名(定期)
        /// </summary>
        public int? ClassAvgRankF { get; set; }

        /// <summary>
        /// 加權平均班排名(定期)
        /// </summary>
        /// 
        public int? ClassAvgRankAF { get; set; }

        /// <summary>
        /// 平均年排名(定期)
        /// </summary>
        public int? YearAvgRankF { get; set; }

        /// <summary>
        /// 加權平均年排名(定期)
        /// </summary>
        /// 
        public int? YearAvgRankAF { get; set; }

        /// <summary>
        /// 總分班排名(定期)
        /// </summary>
        public int? ClassSumRankF { get; set; }

        /// <summary>
        /// 加權總分班排名(定期)
        /// </summary>
        /// 
        public int? ClassSumRankAF { get; set; }

        /// <summary>
        /// 總分年排名(定期)
        /// </summary>
        public int? YearSumRankF { get; set; }

        /// <summary>
        /// 加權總分年排名(定期)
        /// </summary>
        /// 
        public int? YearSumRankAF { get; set; }

        /// <summary>
        /// 計算成績
        /// </summary>
        public void CalScore()
        {
            // 初始
            SumScore = SumScoreA = AvgScore = AvgScoreA = SumScoreF = SumScoreAF = AvgScoreF = AvgScoreAF = 0;

            decimal cot = 0, credit = 0, cotF = 0, creditF = 0;


            // 計算全部
            //foreach (DomainInfo domain in DomainInfoList)
            //{
            //    if (domain.Score.HasValue && domain.Credit.HasValue)
            //    {
            //        SumScore += domain.Score.Value;
            //        SumScoreA += domain.Score.Value * domain.Credit.Value;

            //        cot += 1;
            //        credit += domain.Credit.Value;
            //    }
            //}
            //if (cot > 0)
            //    AvgScore = SumScore / cot;

            //if (credit > 0)
            //    AvgScoreA = SumScoreA / credit;

            // 只有領域，依領域為主
            foreach (DomainInfo domain in DomainInfoList)
            {
                if (Global.SelOnlyDomainList.Contains(domain.Name))
                {
                    // 總成績
                    if (domain.Score.HasValue && domain.Credit.HasValue)
                    {
                        SumScore += domain.Score.Value;
                        SumScoreA += domain.Score.Value * domain.Credit.Value;

                        cot += 1;
                        credit += domain.Credit.Value;
                    }

                    // 定期
                    if (domain.ScoreF.HasValue && domain.Credit.HasValue)
                    {
                        SumScoreF += domain.ScoreF.Value;
                        SumScoreAF += domain.ScoreF.Value * domain.Credit.Value;
                        cotF += 1;
                        creditF += domain.Credit.Value;
                    }
                }
            }

            // 只有科目來計算
            foreach (DomainInfo domain in DomainInfoList)
            {
                foreach(SubjectInfo subj in domain.SubjectInfoList)
                {
                    if (Global.SelOnlySubjectList.Contains(subj.Name))
                    {
                        // 總成績
                        if (subj.Score.HasValue && subj.Credit.HasValue)
                        {
                            SumScore += subj.Score.Value;
                            SumScoreA += subj.Score.Value * subj.Credit.Value;

                            cot += 1;
                            credit += subj.Credit.Value;
                        }

                        // 定期
                        if (subj.ScoreF.HasValue && subj.Credit.HasValue)
                        {
                            SumScoreF += subj.ScoreF.Value;
                            SumScoreAF += subj.ScoreF.Value * subj.Credit.Value;

                            cotF += 1;
                            creditF += subj.Credit.Value;
                        }
                    }
                }
            }

            if (cot > 0)
                AvgScore = SumScore / cot;

            if (credit > 0)
                AvgScoreA = SumScoreA / credit;

            if (cotF > 0)
                AvgScoreF = SumScoreF / cotF;

            if (creditF > 0)
                AvgScoreAF = SumScoreAF / creditF;
        }
    }
}
