using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HsinChuExamScoreClassFixedRank.DAO
{
    public class SubjectInfo
    {

        public string Name { get; set; }

        public string DomainName { get; set; }

        // 總成績
        public decimal? Score { get; set; }
        // 學分數
        public decimal? Credit { get; set; }

        // 定期
        public decimal? ScoreF { get; set; }
        // 定期比例
        public decimal ScoreFP { get; set; }
        // 平時
        public decimal? ScoreA { get; set; }
        // 平時比例
        public decimal ScoreAP { get; set; }

        /// <summary>
        ///  計算科目評量總成績
        /// </summary>
        public void CalcScore()
        {
            // 定期與評量都有
            if (ScoreA.HasValue && ScoreF.HasValue)
            {
                Score = ScoreA.Value * ScoreAP + ScoreF.Value * ScoreFP;
            }
            else if (ScoreA.HasValue && ScoreF.HasValue == false)
            {
                // 只有評量
                Score = ScoreA.Value;
            }
            else if (ScoreA.HasValue == false && ScoreF.HasValue)
            {
                // 只有定期
                Score = ScoreF.Value;
            }
        }
    }
}
