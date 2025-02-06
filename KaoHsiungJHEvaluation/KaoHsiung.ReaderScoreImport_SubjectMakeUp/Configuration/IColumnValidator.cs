using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp
{
    interface IColumnValidator
    {
        bool IsValid(string input);
        string GetErrorMessage();
    }
}
