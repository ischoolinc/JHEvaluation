using JHSchool.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace HsinChuExamScore_JH.DAO
{
    public class AEIncludeData
    {
        public AEIncludeData() { }
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
}
