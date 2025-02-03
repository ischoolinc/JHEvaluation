using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;


namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp.UDT
{
    [TableName("ReaderScoreImport.ClassCode_SubjectMakeUp")]
    public class ClassCode_SubjectMakeUp : ActiveRecord
    {
        [Field(Field = "ClassName", Indexed = true)]
        public string ClassName { get; set; }

        [Field(Field = "Code", Indexed = false)]
        public string Code { get; set; }
    }
}
