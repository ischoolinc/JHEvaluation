using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;

namespace KaoHsingReExamScoreReport.DAO
{
    public class StudentData
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 學生姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public string GradeYear { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string SeatNo { get; set; }

        public string ClassID { get; set; }
        /// <summary>
        /// 學期成績
        /// </summary>
        public JHSemesterScoreRecord SemesterScoreRecord = new JHSemesterScoreRecord();
    }
}
