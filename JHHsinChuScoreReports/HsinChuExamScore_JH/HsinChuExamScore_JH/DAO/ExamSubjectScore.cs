using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        public decimal? FixScorePercentage { get; set; }

        /// <summary>
        /// 計算總成績(國中的評量成績=定期評量&平時評量)
        /// </summary>
        /// <param name="fixScoreWeight">(評分樣板、定期平量比重)</param>
        /// <param name="isReferrnceScore">是否為參考試別成績</param>
        /// <returns></returns>
        public void  GetTotalScore(Boolean isReferrnceScore, decimal fixScoreWeight )
        {
            decimal fixScorePercentage = fixScoreWeight * 0.01M;
            decimal assessScorePercentage = (100 - fixScoreWeight) * 0.01M; // 取得定期，評量由100-定期

           
                if (ScoreF != null && ScoreA != null)
                {
                    this.ScoreT = this.ScoreA * assessScorePercentage + this.ScoreF * fixScorePercentage;
                }

                else if (ScoreF != null && ScoreA == null)
                {
                    this.ScoreT = ScoreF;
                }
                else if (ScoreF == null && ScoreA != null)
                {
                    this.ScoreT = ScoreA;
                }

            if (isReferrnceScore) // 如果有含參考試別 就要計算 參考試的總成績
            {
                if (RefScoreF != null && RefScoreA != null)
                {
                    this.RefScoreT = this.RefScoreA * assessScorePercentage + this.RefScoreF * fixScorePercentage;
                }
                else if (RefScoreF != null && RefScoreA == null) 
                {
                    this.RefScoreT = ScoreF;
                }
                else if (RefScoreF == null && RefScoreA != null)
                {
                    this.RefScoreT = RefScoreA;
                }
            }
        }
    }
}
