using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Evaluation
{
    public class Global
    {
        public static void LoadDomainMap108()
        {
            DomainMap108List.Clear();

            // 加入國語文、英語 主要是因為高雄版在資料庫內沒有語文領域，讓使用者可以明確資料這2個領域也會被算入
            //2022/08 高雄加入本土語文
            DomainMap108List = new List<string>(new string[] { "語文", "數學", "自然科學", "科技", "社會", "藝術", "健康與體育", "綜合活動", "藝術與人文", "自然與生活科技", "國語文", "英語", "本土語文" });
        }

        /// <summary>
        ///  108 課綱領域對照
        /// </summary>
        public static List<string> DomainMap108List = new List<string>();
    }
}
