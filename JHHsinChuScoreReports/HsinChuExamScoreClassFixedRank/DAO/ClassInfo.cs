using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HsinChuExamScoreClassFixedRank.DAO
{
    public class ClassInfo
    {
        public string ClassName { get; set; }
        public string ClassID { get; set; }
        public int GradeYear { get; set; }


        // 班級各項及格人數
        public Dictionary<string, int> ClassScorePassCountDict = new Dictionary<string, int>();

        // 班級各項平均
        public Dictionary<string, decimal> ClassAvgScoreDict = new Dictionary<string, decimal>();

        public List<StudentInfo> Students = new List<StudentInfo>();

        /// <summary>
        /// 計算班級各項及格人數
        /// </summary>
        public void CalClassPassCount()
        {
            foreach (StudentInfo si in Students)
            {
                foreach (DomainInfo di in si.DomainInfoList)
                {
                    if (!ClassScorePassCountDict.ContainsKey(di.Name))
                        ClassScorePassCountDict.Add(di.Name, 0);

                    if (di.Score.HasValue && di.Score.Value >= 60)
                    {
                        ClassScorePassCountDict[di.Name]++;
                    }

                    foreach (SubjectInfo subj in di.SubjectInfoList)
                    {
                        string key = di.Name + "_" + subj.Name;
                        if (!ClassScorePassCountDict.ContainsKey(key))
                            ClassScorePassCountDict.Add(key, 0);

                        if (subj.Score.HasValue && subj.Score.Value >= 60)
                            ClassScorePassCountDict[key]++;
                    }
                }
            }
        }

        /// <summary>
        /// 計算班級各項平均
        /// </summary>
        public void CalClassAvgScore()
        {
            List<string> itemList = new List<string>();
            foreach (StudentInfo si in Students)
            {
                foreach (DomainInfo di in si.DomainInfoList)
                {
                    if (!itemList.Contains(di.Name))
                        itemList.Add(di.Name);

                    string dName = di.Name + "總分";
                    string dNameC = di.Name + "個數";
                    string dNameAvg = di.Name + "平均";
                    if (!ClassAvgScoreDict.ContainsKey(dName))
                    {
                        ClassAvgScoreDict.Add(dName, 0);
                        ClassAvgScoreDict.Add(dNameC, 0);
                        ClassAvgScoreDict.Add(dNameAvg, 0);
                    }


                    if (di.Score.HasValue)
                    {
                        ClassAvgScoreDict[dName] += di.Score.Value;
                        ClassAvgScoreDict[dNameC] += 1;
                    }

                    foreach (SubjectInfo subj in di.SubjectInfoList)
                    {
                        string itemName = di.Name + "_" + subj.Name;
                        if (!itemList.Contains(itemName))
                            itemList.Add(itemName);

                        string key = itemName + "總分";
                        string keyC = itemName + "個數";
                        string keyAvg = itemName + "平均";
                        if (!ClassAvgScoreDict.ContainsKey(key))
                        {
                            ClassAvgScoreDict.Add(key, 0);
                            ClassAvgScoreDict.Add(keyC, 0);
                            ClassAvgScoreDict.Add(keyAvg, 0);
                        }


                        if (subj.Score.HasValue)
                        {
                            ClassAvgScoreDict[key] += subj.Score.Value;
                            ClassAvgScoreDict[keyC] += 1;
                        }

                    }
                }
            }

            // 計算平均
            foreach (string itemName in itemList)
            {
                string k1 = itemName + "總分";
                if (ClassAvgScoreDict.ContainsKey(k1))
                {
                    ClassAvgScoreDict[itemName + "平均"] = Math.Round(ClassAvgScoreDict[k1] / ClassAvgScoreDict[itemName + "個數"], Global.parseNumebr, MidpointRounding.AwayFromZero);
                }
            }
        }
    }
}
