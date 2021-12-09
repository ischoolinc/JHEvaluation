using System;
using System.Collections.Generic;
using System.Text;
using JHSchool.Data;
using System.Xml;

namespace JHEvaluation.ScoreCalculation.ScoreStruct
{
    public class AttendScore : IScore
    {
        /*
        <Extension>
	        <Effort/>
	        <Text/>
            <OrdinarilyEffort/> <!--平時評量努力程度-->
            <OrdinarilyScore/> <!--平時評量成績-->
        </Extension>
         */

        //傳入學年度學期,判斷103-1以前的學年度5:5而以後6:4比例計算
        public AttendScore(JHSCAttendRecord attend, decimal? weight, decimal? period, bool toSemester, string domain,int? schoolYear)
        {
            RawAttend = attend;
            Subscores = new TakeScoreCollection();
            Value = attend.Score;
            Weight = weight;
            Period = period;
            Effort = attend.Effort;
            Text = attend.Text;
            ToSemester = toSemester;
            Domain = domain;
            SchoolYear = schoolYear;

            OrdinarilyEffort = attend.OrdinarilyEffort;
            OrdinarilyScore = attend.OrdinarilyScore;
        }

        /// <summary>
        /// 以 ExamID 為 Key 的成績集合。
        /// </summary>
        public TakeScoreCollection Subscores { get; private set; }

        #region IScore 成員

        /// <summary>
        /// 課程總成績。
        /// </summary>
        public decimal? Value { get; set; }

        /// <summary>
        /// 權重。
        /// </summary>
        public decimal? Weight { get; private set; }

        #endregion

        public int? Effort { get; set; }

        /// <summary>
        /// 課程文字評量。
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// 是否列入到學期成績。
        /// </summary>
        public bool ToSemester { get; private set; }

        /// <summary>
        /// 節數。
        /// </summary>
        public decimal? Period { get; private set; }

        /// <summary>
        /// 領域。
        /// </summary>
        public string Domain { get; private set; }

        /// <summary>
        /// 平時評量努力程度(高雄)
        /// </summary>
        public int? OrdinarilyEffort { get; set; }

        /// <summary>
        /// 平時評量分數評量(高雄)
        /// </summary>
        public decimal? OrdinarilyScore { get; set; }

        public JHSCAttendRecord RawAttend { get; private set; }

        /// <summary>
        /// 學年度
        /// </summary>
        public int? SchoolYear { get; set; }
    }

    /// <summary>
    /// 評量成績。
    /// </summary>
    public class TakeScore : IScore
    {
        /*
         * JHSCETakeRecord
        <Extension>
	        <Effort>0</Effort>
	        <Text/>
	        <Score>56</Score>
	        <AssignmentScore>56</AssignmentScore>
        </Extension>
         */

        /*
         * JHAEIncludeRecord
        <Extension>
	        <UseScore>是</UseScore>
	        <UseEffort>否</UseEffort>
	        <UseText>否</UseText>
	        <UseAssignmentScore>是</UseAssignmentScore>
        </Extension>
         */

        public TakeScore(JHSCETakeRecord sceTake, AEIncludeData include)
        {
            SCETakeData scedata = new SCETakeData(sceTake);
            AEIncludeData aedata = include; //為了效率...
            Text = scedata.Text;
            Effort = null;
            Value = null;
            Weight = include.Weight;

            if (Program.Mode == ModuleMode.KaoHsiung)
            {
                if (aedata.UseScore) Value = scedata.Score;
                if (aedata.UseEffort) Effort = scedata.Effort;
            }
            else if (Program.Mode == ModuleMode.HsinChu)
            {
                // 預設取得比例
                decimal scoreWeight = 0, assignmentWeight = 0, totalWeight = 0;
                decimal? totalScore = null;

                // 使用系統內預設比例
                if (Util.ScorePercentageHSDict.ContainsKey(aedata.RefAssessmentSetupID))
                {
                    scoreWeight = Util.ScorePercentageHSDict[aedata.RefAssessmentSetupID];
                    assignmentWeight = 100 - scoreWeight;
                }
                if (aedata.UseScore && !aedata.UseAssignmentScore)
                {
                    scoreWeight = 100;
                    assignmentWeight = 0;
                }
                if (!aedata.UseScore && aedata.UseAssignmentScore)
                {
                    scoreWeight = 0;
                    assignmentWeight = 100;
                }
                if (!aedata.UseScore && !aedata.UseAssignmentScore)
                {
                    scoreWeight = 0;
                    assignmentWeight = 0;
                }
                if (scedata.Score.HasValue && Program.ScoreValueMap.ContainsKey(scedata.Score.Value))
                {
                    if (!Program.ScoreValueMap[scedata.Score.Value].AllowCalculation)
                    {
                        scoreWeight = 0;
                        assignmentWeight = 100;
                    }
                }
                if (scedata.AssignmentScore.HasValue && Program.ScoreValueMap.ContainsKey(scedata.AssignmentScore.Value))
                {
                    if (!Program.ScoreValueMap[scedata.AssignmentScore.Value].AllowCalculation)
                    {
                        scoreWeight = 100;
                        assignmentWeight = 0;
                    }
                }
                if (scoreWeight>0 && scedata.Score.HasValue)
                {
                    if (Program.ScoreValueMap.ContainsKey(scedata.Score.Value))
                    {
                        if (Program.ScoreValueMap[scedata.Score.Value].AllowCalculation)
                        {
                            totalScore = (totalScore == null ? 0 : totalScore) + Program.ScoreValueMap[scedata.Score.Value].Score.Value * scoreWeight;
                            totalWeight += scoreWeight;
                        }
                    }
                    else
                    {
                        totalScore = (totalScore == null ? 0 : totalScore) + scedata.Score.Value * scoreWeight;
                        totalWeight += scoreWeight;
                    }
                }
                if (assignmentWeight > 0 && scedata.AssignmentScore.HasValue)
                {
                    if (Program.ScoreValueMap.ContainsKey(scedata.AssignmentScore.Value))
                    {
                        if (Program.ScoreValueMap[scedata.AssignmentScore.Value].AllowCalculation)
                        {
                            totalScore = (totalScore == null ? 0 : totalScore) + Program.ScoreValueMap[scedata.AssignmentScore.Value].Score.Value * assignmentWeight;
                            totalWeight += assignmentWeight;
                        }
                    }
                    else
                    {
                        totalScore = (totalScore == null ? 0 : totalScore) + scedata.AssignmentScore.Value * assignmentWeight;
                        totalWeight += assignmentWeight;
                    }
                }

                //if (aedata.UseScore && scedata.Score.HasValue)
                //{
                //    totalScore += scedata.Score.Value * scoreWeight;
                //    totalWeight += scoreWeight;
                //}

                //if (aedata.UseAssignmentScore && scedata.AssignmentScore.HasValue)
                //{
                //    totalScore += scedata.AssignmentScore.Value * assignmentWeight;
                //    totalWeight += assignmentWeight;
                //}

                if (totalWeight != 0)
                    Value = totalScore / totalWeight;
                else
                    Value = null;

                //// 原本作法
                //decimal sum = 0, weight = 0;

                //if (aedata.UseScore && scedata.Score.HasValue)
                //{
                //    sum += scedata.Score.Value;
                //    weight++;
                //}

                //if (aedata.UseAssignmentScore && scedata.AssignmentScore.HasValue)
                //{
                //    sum += scedata.AssignmentScore.Value;
                //    weight++;
                //}

                //if (weight > 0)
                //    Value = (sum / weight);

            }
            else
                throw new ArgumentException(string.Format("沒有此種縣市的處理方式({0})。", Program.Mode.ToString()));
        }

        #region IScore 成員

        /// <summary>
        /// 成績。
        /// </summary>
        public decimal? Value { get; set; }

        /// <summary>
        /// 權重。
        /// </summary>
        public decimal? Weight { get; set; }

        #endregion

        /// <summary>
        /// 努力程度。
        /// </summary>
        public int? Effort { get; private set; }

        /// <summary>
        /// 文字。
        /// </summary>
        public string Text { get; private set; }

        #region HelperClasses
        private class SCETakeData
        {
            public SCETakeData(JHSCETakeRecord record)
            {
                Score = record.Score;
                Text = record.Text;

                XmlElement xmlrecord = record.ToXML();

                #region 嘗試取得 AssignmentScore
                XmlNode assignment = xmlrecord.SelectSingleNode("Extension/Extension/AssignmentScore");
                if (assignment != null) AssignmentScore = ParseDecimalAllowNull(assignment.InnerText);
                #endregion

                #region 嘗試取得 Effort
                XmlNode effort = xmlrecord.SelectSingleNode("Extension/Extension/Effort");
                if (effort != null) Effort = (int?)ParseDecimalAllowNull(effort.InnerText);
                #endregion
            }

            public decimal? Score { get; private set; }

            public decimal? AssignmentScore { get; private set; }

            public string Text { get; private set; }

            public int? Effort { get; private set; }

            private decimal? ParseDecimalAllowNull(string p)
            {
                decimal d;
                if (decimal.TryParse(p, out d))
                    return d;
                return null;
            }
        }

        public class AEIncludeData
        {
            public AEIncludeData(JHAEIncludeRecord record)
            {
                Weight = record.Weight;
                UseScore = record.UseScore;
                UseText = record.UseText;
                UseEffort = false;
                UseAssignmentScore = false;

                RefAssessmentSetupID = record.RefAssessmentSetupID;

                XmlElement xmlrecord = record.ToXML();

                #region 嘗試取得 UseAssignmentScore
                XmlNode assignment = xmlrecord.SelectSingleNode("Extension/Extension/UseAssignmentScore");
                if (assignment != null) UseAssignmentScore = ParseBool(assignment.InnerText);
                #endregion

                #region 嘗試取得 UseEffort
                XmlNode effort = xmlrecord.SelectSingleNode("Extension/Extension/UseEffort");
                if (effort != null) UseEffort = ParseBool(effort.InnerText);
                #endregion
            }

            public bool UseScore { get; private set; }

            public bool UseAssignmentScore { get; private set; }

            public bool UseEffort { get; private set; }

            public bool UseText { get; private set; }

            public decimal Weight { get; private set; }

            private bool ParseBool(string p)
            {
                if (p == "是") return true;
                else return false;
            }

            public string RefAssessmentSetupID { get; private set; }
        }
        #endregion
    }
}
