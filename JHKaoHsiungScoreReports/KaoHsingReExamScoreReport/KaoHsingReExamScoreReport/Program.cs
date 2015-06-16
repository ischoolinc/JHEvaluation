using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KaoHsingReExamScoreReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            RibbonBarItem rbItem2 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbItem2["報表"]["成績相關報表"]["補考名單(給導師)"].Enable = UserAcl.Current["KaoHsingReExamScoreReport.ReDomainForTeacherForm"].Executable;
            rbItem2["報表"]["成績相關報表"]["補考名單(給導師)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    Forms.ReDomainForTeacherForm rdft = new Forms.ReDomainForTeacherForm(K12.Presentation.NLDPanels.Class.SelectedSource);
                    rdft.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選班級");
                    return;
                }

            };

            // 學期成績總表
            Catalog catalog1b = RoleAclSource.Instance["班級"]["功能按鈕"];
            catalog1b.Add(new RibbonFeature("KaoHsingReExamScoreReport.ReDomainForTeacherForm", "補考名單(給導師)"));
        }
    }
}
