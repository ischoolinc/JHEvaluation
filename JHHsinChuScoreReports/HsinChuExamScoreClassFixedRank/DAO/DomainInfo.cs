using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HsinChuExamScoreClassFixedRank.DAO
{
    public class DomainInfo
    {
        public string Name { get; set; }
        // 加權平均
        public decimal? Score { get; set; }
        // 學分數
        public decimal? Credit { get; set; }

        // 加權平均(定期)
        public decimal? ScoreF { get; set; }

        public List<SubjectInfo> SubjectInfoList = new List<SubjectInfo>();

        // 計算領域評量總成績
        public void CalScore()
        {
            // 使用加權平均計算
            decimal? sumScore=null, sumScoreF=null;
            decimal sumCredit = 0, sumCreditF = 0;
            foreach (SubjectInfo si in SubjectInfoList)
            {
                if (si.Score.HasValue && si.Credit.HasValue)
                {
                    sumScore = (sumScore==null?0: sumScore) + si.Score.Value * si.Credit.Value;
                    sumCredit += si.Credit.Value;
                }

                if (si.ScoreF.HasValue && si.Credit.HasValue)
                {
                    if (Program.ScoreValueMap.ContainsKey(si.ScoreF.Value))
                    {
                        if (Program.ScoreValueMap[si.ScoreF.Value].AllowCalculation)
                        {                            
                            sumScoreF =(sumScoreF==null?0:sumScoreF) + Program.ScoreValueMap[si.ScoreF.Value].Score.Value * si.Credit.Value;
                            sumCreditF += si.Credit.Value;
                        }
                    }
                    else
                    {
                        sumScoreF = (sumScoreF == null ? 0 : sumScoreF) + si.ScoreF.Value * si.Credit.Value;
                        sumCreditF += si.Credit.Value;
                    }
                }
            }

            if (sumCredit == 0)
            {
                Score = null;
                Credit = null;
            }else
            {
                if (sumScore != null)
                {
                    Score = sumScore / sumCredit;
                    Credit = sumCredit;
                }
            }

            if (sumCreditF > 0)
            {
                if (sumScoreF != null)
                {
                    ScoreF = sumScoreF / sumCreditF;
                }
            }
        }
    }
}
