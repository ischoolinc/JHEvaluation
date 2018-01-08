using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using KaoHsiung.ReaderScoreImport_DomainMakeUp.Model;
using KaoHsiung.ReaderScoreImport_DomainMakeUp.Mapper;

namespace KaoHsiung.ReaderScoreImport_DomainMakeUp.Validation.RecordValidators
{
    internal class DomainCodeValidator : IRecordValidator<RawData>
    {
        #region IRecordValidator<RawData> 成員
        public string Validate(RawData record)
        {
            if (DomainCodeMapper.Instance.CheckCodeExists(record.DomainCode))
                return string.Empty;
            else
                return string.Format("領域代碼「{0}」不存在。", record.DomainCode);
        }
        #endregion
    }
}
