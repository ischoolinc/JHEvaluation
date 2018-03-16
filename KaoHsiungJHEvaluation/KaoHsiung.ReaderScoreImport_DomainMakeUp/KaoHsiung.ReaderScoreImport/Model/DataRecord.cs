using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using KaoHsiung.ReaderScoreImport_DomainMakeUp.Mapper;

namespace KaoHsiung.ReaderScoreImport_DomainMakeUp.Model
{
    internal class DataRecord
    {
        private RawData _raw;

        //學號{7}，班級{3}，座號{2}，試別{2}，科目{5}，成績{6}
        public string StudentNumber { get; set; }
        public string Class { get; set; }
        public string SeatNo { get; set; }
        public string Exam { get; set; }
        //public List<string> Subjects { get; set; }
        public string Domain { get; set; }


        public decimal Score { get; set; }

        public DataRecord(RawData raw)
        {
            _raw = raw;
            //Subjects = new List<string>();

            StudentNumber = _raw.StudentNumber;
            Class = ClassCodeMapper.Instance.Map(_raw.ClassCode);
            SeatNo = _raw.SeatNo;
            Exam = ExamCodeMapper.Instance.Map(_raw.ExamCode);

            //// 2018/1/5 穎驊註解，不確定為啥以前要這樣寫，感覺科目以前會帶有 "," 逗號，現在統一使用領域來寫
            //string subjects = DomainCodeMapper.Instance.Map(_raw.DomainCode);
            //foreach (string subject in subjects.Split(new string[]{","}, StringSplitOptions.RemoveEmptyEntries))
            //{
            //    string s = subject.Trim();
            //    if(!Subjects.Contains(s))
            //        Subjects.Add(s);
            //}

            Domain = DomainCodeMapper.Instance.Map(_raw.DomainCode);

            Score = ParseScore(_raw.Score);
        }

        private decimal ParseScore(string p)
        {
            decimal d;
            if (decimal.TryParse(p.Trim(), out d))
                return d;
            else
                throw new Exception();
        }
    }

    internal class DataRecordCollection : List<DataRecord>
    {
        internal void ConvertFromRawData(List<RawData> _raws)
        {
            foreach (RawData raw in _raws)
            {
                DataRecord dr = new DataRecord(raw);
                Add(dr);
            }
        }
    }
}
