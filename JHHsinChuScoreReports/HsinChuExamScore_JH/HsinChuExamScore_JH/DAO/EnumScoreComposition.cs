using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HsinChuExamScore_JH.DAO
{
    /// <summary>
    /// 由於國中定期評量成績是由部分法規是由  定期加平時 組成 但學校成績單較常使用定期評量
    /// </summary>
    public enum EnumScoreComposition
    {
       成績, // 定期加平時
       定期成績,
       平時成績
    }
}
