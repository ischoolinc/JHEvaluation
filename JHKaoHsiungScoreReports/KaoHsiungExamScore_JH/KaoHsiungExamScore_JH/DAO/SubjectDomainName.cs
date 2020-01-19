using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KaoHsiungExamScore_JH.DAO
{
    public class SubjectDomainName
    {
        // 2020/1/16 新增。
        public string SCAttendID { get; set; }

        public string SubjectName { get; set; }

        public string DomainName { get; set; }

        public decimal? Credit { get; set; }
    }
}
