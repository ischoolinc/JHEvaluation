﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using KaoHsiung.ReaderScoreImport_DomainMakeUp.UDT;

namespace KaoHsiung.ReaderScoreImport_DomainMakeUp.Mapper
{
    internal class ClassCodeMapper : CodeMapper
    {
        private static CodeMapper _instance;

        public static CodeMapper Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ClassCodeMapper();
                return _instance;
            }
        }

        private ClassCodeMapper()
        {
        }

        protected override void LoadCodes()
        {
            base.LoadCodes();

            AccessHelper helper = new AccessHelper();

            foreach (ClassCode_DomainMakeUp item in helper.Select<ClassCode_DomainMakeUp>())
            {
                if (!CodeMap.ContainsKey(item.Code))
                    CodeMap.Add(item.Code, item.ClassName);
            }
        }
    }
}
