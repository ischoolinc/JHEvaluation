﻿using System;
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
            decimal sumScore = 0,sumCredit = 0,sumScoreF = 0,sumCreditF = 0;
            foreach(SubjectInfo si in SubjectInfoList)
            {
                if (si.Score.HasValue && si.Credit.HasValue)
                {
                    sumScore += si.Score.Value * si.Credit.Value;
                    sumCredit += si.Credit.Value;
                }

                if (si.ScoreF.HasValue && si.Credit.HasValue)
                {
                    sumScoreF += si.ScoreF.Value * si.Credit.Value;
                    sumCreditF += si.Credit.Value;
                }
            }

            if (sumCredit == 0)
            {
                Score = null;
                Credit = null;
            }else
            {
                Score = sumScore / sumCredit;
                Credit = sumCredit;
            }

            if (sumCreditF > 0)
            {
                ScoreF = sumScoreF / sumCreditF;
            }
        }
    }
}