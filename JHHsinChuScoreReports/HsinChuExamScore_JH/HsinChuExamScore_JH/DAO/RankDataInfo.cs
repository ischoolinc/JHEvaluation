using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HsinChuExamScore_JH.DAO
{
    /// <summary>
    /// 排名資料欄位
    /// </summary>
    public class RankDataInfo
    {
        public int matrix_count = 0;  // 總人數
        public int level_gte100 = 0;  // 100以上
        public int level_90 = 0;  // 90以上
        public int level_80 = 0;  // 80以上
        public int level_70 = 0;  // 70以上
        public int level_60 = 0;  // 60以上
        public int level_50 = 0;  // 50以上
        public int level_40 = 0;  // 40以上
        public int level_30 = 0;  // 30以上
        public int level_20 = 0;  // 20以上
        public int level_10 = 0;  // 10以上
        public int level_lt10 = 0;  // 10以下
        public decimal avg_top_25 = 0;  // 頂標
        public decimal avg_top_50 = 0;  // 高標
        public decimal avg = 0;  // 均標
        public decimal avg_bottom_50 = 0;  // 低標
        public decimal avg_bottom_25 = 0;  // 底標
        public int rank = 0;  // 排名
        public int pr = 0;  // PR
        public int percentile = 0;  // 百分比
      
    }
}
