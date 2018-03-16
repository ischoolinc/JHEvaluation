using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace KaoHsiung.ReaderScoreImport_DomainMakeUp.UDT
{
    [TableName("ReaderScoreImport.DomainCode_DomainMakeUp")]
    public class DomainCode_DomainMakeUp : ActiveRecord
    {
        [Field(Field = "Domain", Indexed = true)]
        public string Domain { get; set; }

        [Field(Field = "Code", Indexed = false)]
        public string Code { get; set; }
    }
}
