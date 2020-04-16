using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HsinChuExamScore_JH.DAO
{
    /// <summary>
    /// 國中成績 
    /// </summary>
    public interface IScore
    {
        /// <summary>
        /// 總成績(總成績)
        /// </summary>
         decimal? ScoreT { get; set; }

        /// <summary>
        /// 平時成績
        /// </summary>
         decimal? ScoreA { get; set; }

        /// <summary>
        /// 定期成績
        /// </summary>
         decimal? ScoreF { get; set; }

        /// <summary>
        /// 學分數
        /// </summary>
        decimal? Credit { get; set; }

        /// <summary>
        /// 領域名稱
        /// </summary>
         string DomainName { get; set; }
       
        
        /// <summary>
        /// (參考試別) 總成績
        /// </summary>
        decimal? RefScoreT { get; set; }

        /// <summary>
        /// (參考試別) 平時成績
        /// </summary>
        decimal? RefScoreA { get; set; }

        /// <summary>
        /// (參考試别) 定期評量
        /// </summary>
        decimal? RefScoreF { get; set; }
    }
}
