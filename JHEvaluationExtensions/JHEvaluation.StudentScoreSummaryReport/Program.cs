using System;
using System.Collections.Generic;
using System.Text;
using FISCA;
using Framework.Security;
using FISCA.Presentation;
using K12.Presentation;
using FISCA.Deployment;
using System.Windows.Forms;
using JHSchool.Data;
using System.IO;

namespace JHEvaluation.StudentScoreSummaryReport
{
    public static class Program
    {
        internal static ModuleMode Mode { get; private set; }

        //權限代碼。
        private const string PermissionCode = "JHEvaluation.Student.StudentScoreSummaryReport";
        private const string PermissionCodeEnglish = "JHEvaluation.Student.StudentScoreSummaryReportEnglish";
        private const string PermissionCode2022 = "JHEvaluation.Student.StudentScoreSummaryReport2022";
        private const string PermissionCodeEnglish2022 = "JHEvaluation.Student.StudentScoreSummaryReportEnglish2022";
        [MainMethod()]
        [STAThread()]
        public static void Main()
        {
            DeployModeSetup();

# if LocalDebug
            DeployParameters dparams = ModuleLoader.GetDeployParametsers(typeof(Program), "Mode=KaoHsiung");
            FISCA.Authentication.DSAServices.SetLicense(@"C:\Users\yaoming\Desktop\ischool\SmartSchoolLicense.key");
            FISCA.Authentication.DSAServices.Login("admin", "1234");

            //JHClassRecord cls = JHClass.SelectByID("24");//正興 301 班。

            List<string> students = new List<string>();
            //foreach (JHStudentRecord each in cls.Students)
            //    students.Add(each.ID);

            //呵!
            students.Clear();
            students.Add("852");

            // test.hc 143136

            Aspose.Words.License lic = new Aspose.Words.License();
            lic.SetLicense(new MemoryStream(Prc.Aspose_Total));
            //new PrintForm(students).ShowDialog();
            new PrintFormEnglish(students).ShowDialog();
# else

            //註冊權限管理項目。
            Catalog detail = RoleAclSource.Instance["學生"]["報表"];
            detail.Add(new ReportFeature(PermissionCode, "在校成績證明書"));
            detail.Add(new ReportFeature(PermissionCodeEnglish, "在校成績證明書(英文)"));
            detail.Add(new ReportFeature(PermissionCode2022, "在校成績證明書(新版)"));
            detail.Add(new ReportFeature(PermissionCodeEnglish2022, "在校成績證明書(英文新版)"));

            //註冊報表功能項目。
            MenuButton mb = NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["在校成績證明書"];
            mb.Enable = false;
            mb.Click += delegate
            {
                new PrintForm(K12.Presentation.NLDPanels.Student.SelectedSource).ShowDialog();
            };
            MenuButton mb2 = NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["在校成績證明書(英文版)"];
            mb2.Enable = false;
            mb2.Click += delegate
            {
                new PrintFormEnglish(K12.Presentation.NLDPanels.Student.SelectedSource).ShowDialog();
            };

            //權限判斷。
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                mb.Enable = (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0) &&
                    Framework.User.Acl[PermissionCode].Executable;

                mb2.Enable = (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0) &&
                    Framework.User.Acl[PermissionCodeEnglish].Executable;
            };
#endif
            #region 在校成績證明書 新版
            MenuButton mb3 = NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["在校成績證明書(新版)"];
            mb3.Enable = false;
            mb3.Click += delegate
            {
                new PrintForm_StudentScoreCertificattion(K12.Presentation.NLDPanels.Student.SelectedSource).ShowDialog();
            };

            MenuButton mb4 = NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["在校成績證明書(英文新版)"];
            mb4.Enable = false;
            mb4.Click += delegate
            {
                new PrintForm_StudentScoreCertificattion_English(K12.Presentation.NLDPanels.Student.SelectedSource).ShowDialog();
            };
            //權限判斷。
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                mb3.Enable = (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0) &&
                    Framework.User.Acl[PermissionCode2022].Executable;

                mb4.Enable = (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0) &&
                    Framework.User.Acl[PermissionCodeEnglish2022].Executable;
            };
            #endregion
        }

        private static void DeployModeSetup()
        {
            DeployParameters dparams = ModuleLoader.GetDeployParametsers(typeof(Program), "Mode=KaoHsiung");

            if (dparams["Mode"].ToUpper() == "KaoHsiung".ToUpper())
                Mode = ModuleMode.KaoHsiung; //高雄。
            else
                Mode = ModuleMode.HsinChu;  //新竹。

            //Mode = ModuleMode.KaoHsiung; //高雄。

            Mode = ModuleMode.HsinChu;  //新竹。
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
