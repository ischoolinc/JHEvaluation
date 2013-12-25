using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using FISCA;
using FISCA.Presentation;
using FISCA.Deployment;
using JHSchool.Data;
using JHEvaluation.SemesterScoreContentItem.Forms;

namespace JHEvaluation.SemesterScoreContentItem
{
    public static class Program
    {
        [MainMethod()]
        public static void Main()
        {
            Global.Params = ModuleLoader.GetDeployParametsers(typeof(Program), "Mode=KaoHsiung");

            //學生學期成績
            K12.Presentation.NLDPanels.Student.AddDetailBulider(new DetailBulider<SemesterScoreItem>());

            JHSchool.SF.Evaluation.SemesterScoreEditor.RegisterHandler(delegate(string studentId)
            {
                SemesterScoreEditor form;
                form = new SemesterScoreEditor(JHStudent.SelectByID(studentId));
                return form.ShowDialog();
            });
            JHSchool.SF.Evaluation.SemesterScoreEditor.RegisterHandler(delegate(string studentId, int schoolYear, int semester)
            {
                SemesterScoreEditor form;
                form = new SemesterScoreEditor(JHStudent.SelectByID(studentId), JHSemesterScore.SelectBySchoolYearAndSemester(studentId, schoolYear, semester));
                return form.ShowDialog();
            });
        }
    }

    internal class Global
    {
        internal static DeployParameters Params { get; set; }
    }
}
