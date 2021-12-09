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
        //  使用定期評量
        public bool UseScore { get; set; }
        //  使用平時評量
        public bool UseAssignmentScore { get; set; }

        /// <summary>
        ///  計算科目評量總成績
        /// </summary>
        public void CalcScore()
        {
            // 定期與評量都有
            //if (ScoreA.HasValue && ScoreF.HasValue)
            //{
            //    Score = ScoreA.Value * ScoreAP + ScoreF.Value * ScoreFP;
            //}
            //else if (ScoreA.HasValue && ScoreF.HasValue == false)
            //{
            //    // 只有平時
            //    Score = ScoreA.Value;
            //}
            //else if (ScoreA.HasValue == false && ScoreF.HasValue)
            //{
            //    // 只有定期
            //    Score = ScoreF.Value;
            //}
            if (ScoreFP > 0 && ScoreF.HasValue)
            {
                if (Program.ScoreValueMap.ContainsKey(ScoreF.Value))
                {
                    if (Program.ScoreValueMap[ScoreF.Value].AllowCalculation)
                    {
                        Score = (Score==null?0:Score) + Program.ScoreValueMap[ScoreF.Value].Score.Value * ScoreFP;
                    }
                }
                else
                {
                    Score = (Score == null ? 0 : Score) + ScoreF.Value * ScoreFP;
                }
            }
            if (ScoreAP > 0 && ScoreA.HasValue)
            {
                if (Program.ScoreValueMap.ContainsKey(ScoreA.Value))
                {
                    if (Program.ScoreValueMap[ScoreA.Value].AllowCalculation)
                    {
                        Score = (Score == null ? 0 : Score) + Program.ScoreValueMap[ScoreA.Value].Score.Value * ScoreAP;
                    }
                }
                else
                {
                    Score = (Score == null ? 0 : Score) + ScoreA.Value * ScoreAP;
                }
            }
        }
    }
}
