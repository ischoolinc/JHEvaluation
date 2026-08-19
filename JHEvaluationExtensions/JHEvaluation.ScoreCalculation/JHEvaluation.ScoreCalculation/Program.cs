using System;
using System.Collections.Generic;
using System.Text;
using FISCA;
using FISCA.Deployment;
using Framework.Security;
using FISCA.Presentation;
using K12.Presentation;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Framework;
using JHSchool;
using DataRationality;
using JHSchool.Evaluation.StudentExtendControls;

namespace JHEvaluation.ScoreCalculation
{
    public static class Program
    {
        private static readonly object InitializationSync = new object();
        private static bool _initializationCompleted;
        private static Exception _initializationFailure;

        internal static ModuleMode Mode { get; private set; }

        //權限代碼。
        private const string CalcStudentCode = "JHSchool.Student.Ribbon0057";
        private const string CalcAdminCode = "JHSchool.EduAdmin.Ribbon0045";
        private const string BatchSHistoryCode = "JHSchool.EduAdmin.Ribbon.BatchSemesterHistory"; //批次管理學期歷程。
        private const string GradFilteStudentCode = "JHSchool.Student.Ribbon0058"; //學生畢業資格審查。
        private const string GradFilteAdminCode = "JHSchool.EduAdmin.Ribbon0046"; //教務畢業資格審查。

        public static Dictionary<string, ScoreStruct.ScoreMap> ScoreTextMap = new Dictionary<string, ScoreStruct.ScoreMap>();
        public static Dictionary<decimal, ScoreStruct.ScoreMap> ScoreValueMap = new Dictionary<decimal, ScoreStruct.ScoreMap>();

        [MainMethod()]
        public static void Main()
        {
# if LocalDebug
            DeployParameters dparams = ModuleLoader.GetDeployParametsers(typeof(Program), "Mode=KaoHsiung");
            TestMode(dparams);
            return;
# else
            InitializeModule();
#endif
        }

        private static void InitializeModule()
        {
            lock (InitializationSync)
            {
                if (_initializationCompleted)
                    return;

                if (_initializationFailure != null)
                    throw new InvalidOperationException("The score calculation module previously failed to initialize.", _initializationFailure);

                ModuleLoadDiagnostics diagnostics = ModuleLoadDiagnostics.Start();
                try
                {
                    RegistrationContext context = null;

                    diagnostics.Measure("deploy-mode", DeployModeSetup);
                    diagnostics.SetDeploymentMode(Mode);
                    diagnostics.Measure("student-detail-builder", RegisterGraduationDetailBuilder);
                    diagnostics.Measure("score-menus", delegate { context = RegisterScoreMenus(); });
                    diagnostics.Measure("graduation-reports", delegate { RegisterGraduationAndReports(context); });
                    diagnostics.Measure("acl-selection-events", delegate { RegisterAclAndSelectionEvents(context); });
                    diagnostics.Measure("detail-items", RegisterDetailItems);
                    diagnostics.Measure("data-rationality", RegisterDataRationalityCheck);

                    _initializationCompleted = true;
                    diagnostics.Complete(true, null);
                }
                catch (Exception ex)
                {
                    _initializationFailure = ex;
                    diagnostics.Complete(false, ex);
                    throw;
                }
                finally
                {
                    diagnostics.Flush();
                }
            }
        }

        private static void RegisterGraduationDetailBuilder()
        {
            //2017/5/9 穎驊 自JHSchool.Evaluation 搬過來：畢業成績。
            Student.Instance.AddDetailBulider(new DetailBulider<GraduationScoreItem>());
        }

        private static RegistrationContext RegisterScoreMenus()
        {
            RegistrationContext context = new RegistrationContext();

            MenuButton adminScoreMenu = JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["批次作業/檢視"]["成績作業"];
            MenuButton semesterHistoryButton = adminScoreMenu["產生學期歷程"];
            semesterHistoryButton.Enable = Framework.User.Acl[BatchSHistoryCode].Executable;
            semesterHistoryButton.Click += delegate
            {
                new JHEvaluation.ScoreCalculation.SemesterHistory.BatchSemesterHistory().ShowDialog();
            };

            MenuButton studentScoreMenu = NLDPanels.Student.RibbonBarItems["教務"]["成績作業"];
            studentScoreMenu.Enable = false;
            studentScoreMenu["計算科目成績"].Click += delegate { new SubjectScoreCalculate(NLDPanels.Student.SelectedSource).ShowDialog(); };
            studentScoreMenu["計算領域成績"].Click += delegate { new DomainScoreCalculate(NLDPanels.Student.SelectedSource).ShowDialog(); };
            studentScoreMenu["計算學習領域成績"].Click += delegate { new LearningDomainScoreCalculate(NLDPanels.Student.SelectedSource).ShowDialog(); };
            studentScoreMenu["加總學習領域文字描述"].Click += delegate { new DomainTextScoreSum(NLDPanels.Student.SelectedSource).ShowDialog(); };

            JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["批次作業/檢視"]["成績作業"].Size = RibbonBarButton.MenuButtonSize.Large;
            adminScoreMenu.Image = Properties.Resources.calc_save_64;
            adminScoreMenu["批次計算科目成績"].Click += delegate { new SubjectScoreCalculateByGradeyear().ShowDialog(); };
            adminScoreMenu["批次計算領域成績"].Click += delegate { new DomainScoreCalculateByGradeyear().ShowDialog(); };
            adminScoreMenu["批次計算學習領域成績"].Click += delegate { new LearningDomainScoreCalculateByGradeyear().ShowDialog(); };
            adminScoreMenu["批次加總學習領域文字描述"].Click += delegate { new DomainTextScoreSumByGradeyear().ShowDialog(); };

            context.StudentScoreMenu = studentScoreMenu;
            context.AdminScoreMenu = adminScoreMenu;
            return context;
        }

        private static void RegisterGraduationAndReports(RegistrationContext context)
        {
            RibbonBarButton studentGraduation = NLDPanels.Student.RibbonBarItems["教務"]["畢業作業"];
            MenuButton calculateGraduation = studentGraduation["計算畢業成績"];
            calculateGraduation.Enable = User.Acl[CalcStudentCode].Executable;
            calculateGraduation.Click += delegate
            {
                new GraduateScoreCalculate(NLDPanels.Student.SelectedSource).ShowDialog();
            };

            RibbonBarButton reportMenu = Student.Instance.RibbonBarItems["資料統計"]["報表"];
            MenuButton scoreWarningReport = reportMenu["成績相關報表"]["畢業預警報表"];
            MenuButton behaviorWarningReport = reportMenu["學務相關報表"]["畢業預警報表"];
            bool canRunWarningReport = User.Acl["JHSchool.Student.Report0010"].Executable;
            scoreWarningReport.Enable = canRunWarningReport;
            scoreWarningReport.Click += delegate { RunGraduationPredictReport(); };
            behaviorWarningReport.Enable = canRunWarningReport;
            behaviorWarningReport.Click += delegate { RunGraduationPredictReport(); };

            MenuButton studentGraduationFilter = studentGraduation["畢業資格審查"];
            studentGraduationFilter.Enable = User.Acl[GradFilteStudentCode].Executable;
            studentGraduationFilter.Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count == 0) return;
                Form form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationInspectWizard("Student");
                form.ShowDialog();
            };

            RibbonBarButton adminGraduation = JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["批次作業/檢視"]["畢業作業"];
            adminGraduation.Size = RibbonBarButton.MenuButtonSize.Large;
            MenuButton adminCalculateGraduation = adminGraduation["計算畢業成績"];
            bool canCalculateAdmin = User.Acl[CalcAdminCode].Executable;
            adminCalculateGraduation.Enable = canCalculateAdmin;
            adminCalculateGraduation.Visible = false;
            adminCalculateGraduation.Click += delegate { };

            MenuButton adminGraduationFilter = adminGraduation["畢業資格審查"];
            adminGraduationFilter.Enable = User.Acl[GradFilteAdminCode].Executable;
            adminGraduationFilter.Click += delegate
            {
                Form form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationInspectWizard("EduAdmin");
                form.ShowDialog();
            };

            context.StudentCalculateGraduation = calculateGraduation;
            context.StudentGraduationFilter = studentGraduationFilter;
            context.CanCalculateAdmin = canCalculateAdmin;
        }

        private static void RegisterAclAndSelectionEvents(RegistrationContext context)
        {
            Catalog academicCatalog = RoleAclSource.Instance["教務作業"];
            academicCatalog.Add(new RibbonFeature(BatchSHistoryCode, "批次產生學期歷程"));
            context.AdminScoreMenu.Enable = context.CanCalculateAdmin;

            NLDPanels.Student.SelectedSourceChanged += delegate
            {
                bool hasSelectedStudents = NLDPanels.Student.SelectedSource.Count > 0;
                context.StudentScoreMenu.Enable = hasSelectedStudents && User.Acl[CalcStudentCode].Executable;
                context.StudentCalculateGraduation.Enable = hasSelectedStudents && User.Acl[CalcStudentCode].Executable;
                context.StudentGraduationFilter.Enable = hasSelectedStudents && User.Acl[GradFilteStudentCode].Executable;
            };
        }

        private static void RegisterDetailItems()
        {
            Catalog detail = RoleAclSource.Instance["學生"]["資料項目"];
            detail.Add(new DetailItemFeature(typeof(SemesterScoreItem)));
            detail.Add(new DetailItemFeature(typeof(GraduationScoreItem)));
            detail.Add(new DetailItemFeature(typeof(CourseScoreItem)));
        }

        private static void RegisterDataRationalityCheck()
        {
            DataRationalityManager.Checks.Add(new JHSchool.Evaluation.StudentExtendControls.Ribbon.CheckStudentSemHistoryScoreRAT());
        }

        private static void RunGraduationPredictReport()
        {
            if (Student.Instance.SelectedList.Count <= 0) return;
            new JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReport(Student.Instance.SelectedList);
        }

        private static void DeployModeSetup()
        {
            DeployParameters dparams = ModuleLoader.GetDeployParametsers(typeof(Program), "Mode=HsinChu");

            //debug
            //Mode = ModuleMode.HsinChu;
            //return;

            if (dparams["Mode"].ToUpper() == "KaoHsiung".ToUpper())
                Mode = ModuleMode.KaoHsiung; //高雄。
            else
                Mode = ModuleMode.HsinChu;  //新竹。

            //// debug 開發完後要把這段註解!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            //Mode = ModuleMode.KaoHsiung;
        }

        //private static void TestMode(DeployParameters dparams)
        //{
        //    FISCA.Authentication.DSAServices.SetLicense(@"C:\Users\yaoming\Desktop\ischool.dir\SmartSchoolLicense.key");
        //    FISCA.Authentication.DSAServices.Login("admin", "1234");

        //    new CalculationTest(dparams["Mode"]).ShowDialog();
        //}

        private sealed class RegistrationContext
        {
            internal MenuButton StudentScoreMenu { get; set; }
            internal MenuButton AdminScoreMenu { get; set; }
            internal MenuButton StudentCalculateGraduation { get; set; }
            internal MenuButton StudentGraduationFilter { get; set; }
            internal bool CanCalculateAdmin { get; set; }
        }
    }

    internal enum ModuleMode
    {
        /// <summary>
        /// 新竹
        /// </summary>
        HsinChu,
        /// <summary>
        /// 高雄
        /// </summary>
        KaoHsiung
    }
}
