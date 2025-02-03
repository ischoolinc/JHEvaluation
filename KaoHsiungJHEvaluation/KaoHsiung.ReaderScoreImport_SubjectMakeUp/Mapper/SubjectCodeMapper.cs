using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.UDT;

namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp.Mapper
{
    internal class SubjectCodeMapper : CodeMapper
    {
        private static CodeMapper _instance;

        public static CodeMapper Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new SubjectCodeMapper();
                return _instance;
            }
        }

        private SubjectCodeMapper()
        {
        }

        protected override void LoadCodes()
        {
            base.LoadCodes();

            AccessHelper helper = new AccessHelper();

            foreach (SubjectCode_SubjectMakeUp item in helper.Select<SubjectCode_SubjectMakeUp>())
            {
                if (!CodeMap.ContainsKey(item.Code))
                    CodeMap.Add(item.Code, item.Subject);
            }
        }
    }
}
