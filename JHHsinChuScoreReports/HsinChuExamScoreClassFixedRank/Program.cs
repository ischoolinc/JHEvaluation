using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using FISCA.Permission;
using FISCA.Presentation;


namespace HsinChuExamScoreClassFixedRank
{
    public class Program
    {

        [FISCA.MainMethod]
        public static void Main()
        {
            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbItem1["報表"]["成績相關報表"]["班級評量成績通知單(固定排名)"].Enable = UserAcl.Current["JH.Student.HsinChuExamScoreClassFixedRank"].Executable;
            rbItem1["報表"]["成績相關報表"]["班級評量成績通知單(固定排名)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {

                    Form.PrintForm pf = new Form.PrintForm();
                    pf.SetClassIDList(K12.Presentation.NLDPanels.Class.SelectedSource);
                    pf.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選班級");
                    return;
                }

            };

            // 評量成績通知單
            Catalog catalog1 = RoleAclSource.Instance["班級"]["功能按鈕"];
            catalog1.Add(new RibbonFeature("JH.Student.HsinChuExamScoreClassFixedRank", "班級評量成績通知單(固定排名)"));
        }
    }
}
