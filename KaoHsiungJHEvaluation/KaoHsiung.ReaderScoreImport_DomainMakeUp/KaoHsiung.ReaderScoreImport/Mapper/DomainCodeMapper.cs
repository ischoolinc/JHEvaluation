using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using KaoHsiung.ReaderScoreImport_DomainMakeUp.UDT;

namespace KaoHsiung.ReaderScoreImport_DomainMakeUp.Mapper
{
    internal class DomainCodeMapper : CodeMapper
    {
        private static CodeMapper _instance;

        public static CodeMapper Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DomainCodeMapper();
                return _instance;
            }
        }

        private DomainCodeMapper()
        {
        }

        protected override void LoadCodes()
        {
            base.LoadCodes();

            AccessHelper helper = new AccessHelper();

            foreach (DomainCode_DomainMakeUp item in helper.Select<DomainCode_DomainMakeUp>())
            {
                if (!CodeMap.ContainsKey(item.Code))
                    CodeMap.Add(item.Code, item.Domain);
            }
        }
    }
}
