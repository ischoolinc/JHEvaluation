﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp.Validation
{
    internal interface IRecordValidator<T>
    {
        string Validate(T record);
    }
}
