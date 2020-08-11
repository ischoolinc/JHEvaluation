using FISCA;
using FISCA.Deployment;
using FISCA.Presentation;
using Framework;
using Framework.Security;
using K12.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HsinChu.StudentRecordReportVer2
{
    public static class Program
    {
        //權限代碼
        private const string PermissionCode_StudentRecordReport = "HsinChu.JHEvaluation.Student.StudentReport";

        [MainMethod()]
        public static void Main()
        {
            //註冊權限管理項目
            Catalog detail = RoleAclSource.Instance["學生"]["報表"];
            detail.Add(new ReportFeature(PermissionCode_StudentRecordReport, "學籍表-十二年國教適用"));

            //註冊報表功能項目
            MenuButton mb0 = NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["學籍表-十二年國教適用"];
            mb0.Enable = User.Acl[PermissionCode_StudentRecordReport].Executable;
            mb0.Click += delegate
            {
                new PrintForm_StudentRecordReport(NLDPanels.Student.SelectedSource).ShowDialog();
            };

            //權限判斷
            NLDPanels.Student.SelectedSourceChanged += delegate
            {
                mb0.Enable = (NLDPanels.Student.SelectedSource.Count > 0)
                    && Framework.User.Acl[PermissionCode_StudentRecordReport].Executable;
            };
        }
    }
}
