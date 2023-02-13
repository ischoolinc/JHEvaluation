using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HsinChu.ClassExamScoreAvgComparison.DAL
{
    public class ScoreMap
    {
        public string UseText { set; get; }
        public bool AllowCalculation { set; get; }
        public decimal? Score { set; get; }
        public bool Active { set; get; }
        public decimal UseValue { set; get; }

    }
}
