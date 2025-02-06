using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.Model;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.Mapper;

namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp.Validation.RecordValidators
{
    internal class SubjectCodeValidator : IRecordValidator<RawData>
    {
        #region IRecordValidator<RawData> 成員
        public string Validate(RawData record)
        {
            if (SubjectCodeMapper.Instance.CheckCodeExists(record.SubjectCode))
                return string.Empty;
            else
                return string.Format("領域代碼「{0}」不存在。", record.SubjectCode);
        }
        #endregion
    }
}
