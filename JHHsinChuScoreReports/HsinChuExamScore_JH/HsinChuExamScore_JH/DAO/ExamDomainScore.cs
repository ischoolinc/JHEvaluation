﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HsinChuExamScore_JH.DAO
{
    /// <summary>
    /// 評量領域成績
    /// </summary>
    public class ExamDomainScore : IScore
    {
        /// <summary>
        /// 領域名稱
        /// </summary>
        public string DomainName { get; set; }

        /// <summary>
        /// 定期評量學分加總
        /// </summary>
        public decimal? CreditF { get; set; }

        /// <summary>
        /// 平時評量學分加總
        /// </summary>
        public decimal? CreditA { get; set; }

        /// <summary>
        /// 領域學分
        /// </summary>
        public decimal? Credit { get; set; }

        /// <summary>
        /// 有成績的領域學分
        /// </summary>
        public decimal? Credit1 { get; set; }

        /// <summary>
        /// 領域成績(總成績)
        /// </summary>
        public decimal? ScoreT { get; set; }

        /// <summary>
        /// 領域平時成績
        /// </summary>
        public decimal? ScoreA { get; set; }

        /// <summary>
        /// 領域定期成績
        /// </summary>
        public decimal? ScoreF { get; set; }


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
       public  decimal? RefScoreF { get; set; }

    }
}
