using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp.UDT
{
    [TableName("ReaderScoreImport.SubjectCode_SubjectMakeUp")]
    public class SubjectCode_SubjectMakeUp : ActiveRecord
    {
        [Field(Field = "Subject", Indexed = true)]
        public string Subject { get; set; }

        [Field(Field = "Code", Indexed = false)]
        public string Code { get; set; }
    }
}
