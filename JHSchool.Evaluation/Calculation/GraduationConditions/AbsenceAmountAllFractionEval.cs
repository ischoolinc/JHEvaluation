using System.Collections.Generic;
using System.Xml;
using JHSchool.Behavior.BusinessLogic;
using JHSchool.Data;

namespace JHSchool.Evaluation.Calculation.GraduationConditions
{
    /// <summary>
    /// 判斷全部學期的缺曠節數是否小於上課時數的某百分比
    /// </summary>
    internal class AbsenceAmountAllFractionEval : IEvaluative
    {
        private EvaluationResult _result;
        private Dictionary<string, decimal> _types, _typeWeight,_AvoidMapping;
        private Dictionary<string, int> _periodMapping;
        List<string> _SelectedType;
        private decimal _amount = 100;
        private int _dayPeriod;

        /// <summary>
        /// XML參數建構式
        /// <![CDATA[
        /// <條件 Checked="True" Type="AbsenceAmountEachFraction" 假別="一般:事假,病假,曠課" 節數="1/2"/>
        /// ]]>
        /// </summary>
        /// <param name="element"></param>
        public AbsenceAmountAllFractionEval(XmlElement element)
        {
            _result = new EvaluationResult();
            _periodMapping = new Dictionary<string, int>();
            _types = new Dictionary<string, decimal>();
            _typeWeight = new Dictionary<string, decimal>();

            foreach (JHPeriodMappingInfo info in JHSchool.Data.JHPeriodMapping.SelectAll())
            {
                if (!_periodMapping.ContainsKey(info.Type))
                    _periodMapping.Add(info.Type, 0);
                _periodMapping[info.Type]++;
            }

            _SelectedType = new List<string>();
            string types = element.GetAttribute("假別");
            foreach (string typeline in types.Split(';'))
            {
                if (typeline.Split(':').Length != 2) continue;

                string type = typeline.Split(':')[0];

                decimal weight = 1;
                if (type.EndsWith(")") && type.Contains("("))
                {
                    type = type.Substring(0, type.Length - 1);
                    //weight = type.Split('(')[1];
                    if (!decimal.TryParse(type.Split('(')[1], out weight))
                        weight = 1;
                    type = type.Split('(')[0];
                }

                _SelectedType.Add(type);
                foreach (string absence in typeline.Split(':')[1].Split(','))
                {
                    _types.Add(type + ":" + absence, weight);

                    //各節次type的權重
                    if (!_typeWeight.ContainsKey(type))
                    {
                        _typeWeight.Add(type, weight);
                    }
                }
                    
            }

            _dayPeriod = 0;
            foreach (string type in _SelectedType)
            {
                if (_periodMapping.ContainsKey(type))
                    _dayPeriod += _periodMapping[type];
            }

            //foreach (int gy in new List<int>(_gradeYearMapping.Keys))
            //{
            //    decimal num = _gradeYearMapping[gy];
            //    num = num * dayPeriod;
            //    num = num * _amount / 100;
            //    _gradeYearMapping[gy] = num;
            //}

            string amount = element.GetAttribute("節數");
            if (amount.Contains("/"))
            {
                decimal f1 = decimal.Parse(amount.Split('/')[0]);
                decimal f2 = decimal.Parse(amount.Split('/')[1]);

                _amount = f1 / f2 * 100;
            }
            else if (amount.EndsWith("%"))
                _amount = decimal.Parse(amount.Substring(0, amount.Length - 1));
            //SchoolDayCountG1
            //SCHOOL_HOLIDAY_CONFIG_STRING
            //CONFIG_STRING
            //<條件 Checked="True" Type="AbsenceAmountEachFraction"
            //      假別="一般:事假,病假,曠課" 節數="1/2"/>

        }

        #region IEvaluative 成員

        Dictionary<string, bool> IEvaluative.Evaluate(IEnumerable<StudentRecord> list)
        {
            _result.Clear();

            Dictionary<string, bool> passList = new Dictionary<string, bool>();

            Dictionary<string, List<AutoSummaryRecord>> morals = new Dictionary<string, List<AutoSummaryRecord>>();
            foreach (AutoSummaryRecord record in AutoSummary.Select(list.AsKeyList(), null))
            {
                if (!morals.ContainsKey(record.RefStudentID))
                    morals.Add(record.RefStudentID, new List<AutoSummaryRecord>());
                morals[record.RefStudentID].Add(record);
            }

            foreach (StudentRecord student in list)
            {
                _AvoidMapping = new Dictionary<string, decimal>();
                List<K12.Data.AbsenceMappingInfo> infoList = K12.Data.AbsenceMapping.SelectAll();
                foreach (K12.Data.AbsenceMappingInfo each in infoList)
                {
                    if (each.Noabsence)
                    {
                        foreach (string type in _typeWeight.Keys)
                        {
                            _AvoidMapping.Add(type + ":" + each.Name, 0);
                        }
                    }
                }

                passList.Add(student.ID, true);

                Dictionary<SemesterInfo, int> gyMapping = new Dictionary<SemesterInfo, int>();
                Dictionary<SemesterInfo, decimal> schoolDayMapping = new Dictionary<SemesterInfo, decimal>();
                foreach (K12.Data.SemesterHistoryItem item in UIConfig._StudentSHistoryRecDict[student.ID].SemesterHistoryItems)
                {
                    SemesterInfo info = new SemesterInfo();
                    info.SchoolYear = item.SchoolYear;
                    info.Semester = item.Semester;
                    if (!gyMapping.ContainsKey(info))
                    {
                        gyMapping.Add(info, item.GradeYear);

                        if (item.SchoolDayCount.HasValue)
                        {
                            decimal num = (decimal)item.SchoolDayCount.Value;

                            //設定臨界值
                            decimal newNum = 0;
                            foreach (string type in _SelectedType)
                            {
                                newNum += num * _periodMapping[type] * _typeWeight[type];
                            }
                            //newNum = newNum * _amount / 100;
                            schoolDayMapping.Add(info, newNum);

                            //num *= _dayPeriod;
                            //num = num * _amount / 100;
                            //schoolDayMapping.Add(info, num);
                        }
                    }
                }

                if (!morals.ContainsKey(student.ID)) continue;
                Dictionary<SemesterInfo, decimal> counter = new Dictionary<SemesterInfo, decimal>();
                foreach (AutoSummaryRecord record in morals[student.ID])
                {
                    SemesterInfo info = new SemesterInfo();
                    info.SchoolYear = record.SchoolYear;
                    info.Semester = record.Semester;

                    foreach (AbsenceCountRecord acRecord in record.AbsenceCounts)
                    {
                        //加總各項核可的假別數
                        if (_AvoidMapping.ContainsKey(acRecord.PeriodType + ":" + acRecord.Name))
                        {
                            decimal weight = _typeWeight.ContainsKey(acRecord.PeriodType) ? _typeWeight[acRecord.PeriodType] : 0;
                            _AvoidMapping[acRecord.PeriodType + ":" + acRecord.Name] += acRecord.Count * weight;
                        }

                        if (!_types.ContainsKey(acRecord.PeriodType + ":" + acRecord.Name)) continue;

                        if (!counter.ContainsKey(info))
                            counter.Add(info, 0);
                        counter[info] += acRecord.Count * _types[acRecord.PeriodType + ":" + acRecord.Name];
                    }
                }

                List<ResultDetail> resultList = new List<ResultDetail>();
                decimal count = 0;
                decimal schoolDay = 0;

                foreach (SemesterInfo info in counter.Keys)
                {
                    count += counter[info];
                }

                foreach (KeyValuePair<SemesterInfo, decimal> kvp in schoolDayMapping)
                {
                    schoolDay += kvp.Value;
                }

                //循環要扣除的假別數
                foreach (string elem in _AvoidMapping.Keys)
                {
                    //不是要count的假別就扣除總節數
                    if (!_types.ContainsKey(elem))
                    {
                        schoolDay -= _AvoidMapping[elem];
                    }
                }

                //總節數乘上設定比例
                schoolDay *= _amount / 100;

                if (count > schoolDay)
                {
                    ResultDetail rd = new ResultDetail(student.ID, "0", "0");
                    rd.AddMessage("上課天數不足");
                    rd.AddDetail("上課天數不足(累計" + count + "節)");
                    resultList.Add(rd);
                }

                if (resultList.Count > 0)
                {
                    _result.Add(student.ID, resultList);
                    passList[student.ID] = false;
                }
            }

            return passList;
        }

        public EvaluationResult Result
        {
            get { return _result; }
        }

        #endregion
    }
}