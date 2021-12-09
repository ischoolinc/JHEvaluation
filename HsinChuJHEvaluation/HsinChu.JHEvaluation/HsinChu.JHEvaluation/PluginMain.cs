using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA;
using FISCA.Presentation;
using JHSchool;

namespace HsinChu.JHEvaluation
{
    public static class PluginMain
    {

        public static Dictionary<string, ScoreMap> ScoreTextMap = new Dictionary<string, ScoreMap>();
        public static Dictionary<decimal, ScoreMap> ScoreValueMap = new Dictionary<decimal, ScoreMap>();

        [Dependency("JHSchool.Evaluation")]
        [MainMethod("HsinChu.JHEvaluation")]
        public static void Main()
        {
            //if (System.IO.File.Exists(System.IO.Path.Combine(Application.StartupPath, "新竹開發")))
            //{
            SetupStudent.Init();
            SetupCourse.Init();
            SetupEduAdmin.Init();
            //}
        }
    }
}
