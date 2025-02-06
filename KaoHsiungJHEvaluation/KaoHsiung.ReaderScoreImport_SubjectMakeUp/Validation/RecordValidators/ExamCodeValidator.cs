using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.Model;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.Mapper;

namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp.Validation.RecordValidators
{
    internal class ExamCodeValidator : IRecordValidator<RawData>
    {
        #region IRecordValidator<RawData> 成員
        public string Validate(RawData record)
        {
            //if (ExamCodeMapper.Instance.CheckCodeExists(record.ExamCode))
            //    return string.Empty;
            //else
            //    return string.Format("試別代碼「{0}」不存在。", record.ExamCode);
            //if (record.ExamCode == "01")
            //    return string.Empty;
            //else
            //    return "補考成績匯入試別代碼僅能為「01」。";
            return string.Empty;
        }
        #endregion
    }
}
