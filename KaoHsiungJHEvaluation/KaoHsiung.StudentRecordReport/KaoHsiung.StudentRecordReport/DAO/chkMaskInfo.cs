using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KaoHsiung.StudentRecordReport.DAO
{
    // 檢查是否遮罩使用
    public class chkMaskInfo
    {
        // 遮罩姓名
        public bool isMaskName { get; set; }

        // 遮罩身分證號
        public bool isMaskIDNumber { get; set; }
        // 遮罩生日
        public bool isMaskBirthday { get; set; }

        // 遮罩電話
        public bool isMaskPhone { get; set; }

        // 遮罩地址
        public bool isMaskAddress { get; set; }
    }
}
