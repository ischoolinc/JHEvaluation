using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Evaluation.ClassExtendControls.Ribbon.Score.SemsGoodStudentReport.Data
{
    public class ClassData
    {
        public K12.Data.ClassRecord ClassRecord { get; set; }
        public List<StudentData> Students { get; set; }

        public ClassData(K12.Data.ClassRecord classRecord)
        {
            ClassRecord = classRecord;
            Students = new List<StudentData>();

            foreach (K12.Data.StudentRecord sr in classRecord.Students)
            {
                StudentData sd = new StudentData();
                sd.Student = sr;
                sd.Score = 0;

                Students.Add(sd);
            }
        }

        public void Sort()
        {
            Students.Sort(new StudentComparer());
        }

        public List<StudentData> GetTopStudent(int rank)
        {
            List<StudentData> studentRank = new List<StudentData>();

            decimal currentScore = decimal.MaxValue;
            int currentRank = 0;
           
            for (int i = 0; i < Students.Count; i++)
            {
                StudentData sd = Students[i];

                if (sd.Score == currentScore)
                {
                    sd.Rank = currentRank;
                }
                else
                {
                    sd.Rank = (i + 1);
                    currentRank = sd.Rank;
                    currentScore = sd.Score;
                }                
            }

            foreach (StudentData sd in Students)
            {
                if (sd.Rank <= rank)
                    studentRank.Add(sd);
            }

            return studentRank;
        }
    }
}
