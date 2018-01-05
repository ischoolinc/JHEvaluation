using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace KaoHsiung.ReaderScoreImport_DomainMakeUp.UDT
{
    [TableName("ReaderScoreImport.ClassCode_DomainMakeUp")]
    public class ClassCode_DomainMakeUp : ActiveRecord
    {
        [Field(Field = "ClassName", Indexed = true)]
        public string ClassName { get; set; }

        [Field(Field = "Code", Indexed = false)]
        public string Code { get; set; }
    }
}
