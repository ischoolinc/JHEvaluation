using FISCA;
using FISCA.Presentation;
using FISCA.Permission;
using K12.Presentation;
using FISCA.Presentation.Controls;

namespace DomainScoreReport
{
    public class Program
    {
        [MainMethod("DomainScoreReport")]
        public static void Main()
        {
            #region 學生
            {
                string code = "Student-Domain-Score-Report";
                RoleAclSource.Instance["學生"]["報表"].Add(new RibbonFeature(code, "成績預警通知單"));
                MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["成績相關報表"]["成績預警通知單"].Enable = UserAcl.Current[code].Executable;
                MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["成績相關報表"]["成績預警通知單"].Click += delegate
                {
                    if (NLDPanels.Student.SelectedSource.Count > 0)
                    {
                        (new ExportStudentDomainScore(NLDPanels.Student.SelectedSource)).Export();
                    }
                    else
                    {
                        MsgBox.Show("請選擇要列印成績單的學生。");
                    }
                };
            }
            #endregion

            #region 班級
            {
                string code = "Class-Domain-Score-Report";
                RoleAclSource.Instance["班級"]["報表"].Add(new RibbonFeature(code, "成績預警通知單"));
                MotherForm.RibbonBarItems["班級", "資料統計"]["報表"]["成績相關報表"]["成績預警通知單"].Enable = UserAcl.Current[code].Executable;
                MotherForm.RibbonBarItems["班級", "資料統計"]["報表"]["成績相關報表"]["成績預警通知單"].Click += delegate
                {
                    if (NLDPanels.Class.SelectedSource.Count > 0)
                    {
                        (new frmExportClassDomainScore(NLDPanels.Class.SelectedSource)).Show();
                    }
                    else
                    {
                        MsgBox.Show("請選擇要列印成績單的班級");
                    }
                };
            }
            #endregion

            #region 教務作業
            {
                string code = "School-Domain-Score-Report";
                RoleAclSource.Instance["教務作業"].Add(new RibbonFeature(code, "領域不及格人數統計表"));
                MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["領域不及格人數統計表"].Enable = UserAcl.Current[code].Executable;
                MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["領域不及格人數統計表"].Click += delegate
                {
                    (new frmExportSchoolDomainScore()).Show();
                };
            }
            #endregion

        }
    }
}
