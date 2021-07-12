using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KaoHsiung.StudentRecordReport
{
    public class DomainSorter
    {
        //private static List<string> _list = new List<string>(new string[] { "國語文", "英語", "數學", "社會", "藝術與人文", "自然與生活科技", "健康與體育", "綜合活動", "彈性課程" });
        private static List<string> _list = new List<string>(new string[] { "國語文","英語","數學","社會","自然科學","自然與生活科技","藝術","藝術與人文","健康與體育","綜合活動","科技","實用語文","實用數學","社會適應","生活教育","休閒教育","職業教育","特殊需求","體育專業","藝術才能專長","彈性課程" });

        public static int Sort1(string x, string y)
        {
            int ix = _list.IndexOf(x);
            int iy = _list.IndexOf(y);

            if (ix >= 0 && iy >= 0) return ix.CompareTo(iy);
            else if (ix >= 0) return -1;
            else if (iy >= 0) return 1;
            else return x.CompareTo(y);
        }
    }
}
