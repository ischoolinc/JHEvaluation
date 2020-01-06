using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using FISCA.Permission;
using FISCA.Presentation;

namespace HsinChuSemesterScoreFixed_JH
{
    /// <summary>
    /// 新竹學期成績單
    /// </summary>
    public class Program
    {
        static DataTable _dtEpost = new DataTable();

        [FISCA.MainMethod]
        public static void Main()
        {
            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["學生", "資料統計"];
            rbItem1["報表"]["成績相關報表"]["學期成績通知單(固定排名)"].Enable = UserAcl.Current["9C776D1D-5C08-482A-95EB-91CBD965011C"].Executable;
            rbItem1["報表"]["成績相關報表"]["學期成績通知單(固定排名)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && K12.Presentation.NLDPanels.Student.SelectedSource.Count < 111)
                {
                    PrintForm pf = new PrintForm(K12.Presentation.NLDPanels.Student.SelectedSource);
                    pf.ShowDialog();
                }
               
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選學生");
                    return;
                }
            };

            RibbonBarItem rbItem2 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbItem2["報表"]["成績相關報表"]["學期成績通知單(固定排名)"].Enable = UserAcl.Current["9C776D1D-5C08-482A-95EB-91CBD965011C"].Executable;
            rbItem2["報表"]["成績相關報表"]["學期成績通知單(固定排名)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && K12.Presentation.NLDPanels.Class.SelectedSource.Count < 4)
                {
                    List<string> StudentIDList = Utility.GetClassStudentIDList1ByClassID(K12.Presentation.NLDPanels.Class.SelectedSource);
                    PrintForm pf = new PrintForm(StudentIDList);
                    pf.ShowDialog();
                }
                
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選班級");
                    return;
                }

            };
            // 學期成績通知單
            Catalog catalog1a = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog1a.Add(new RibbonFeature("9C776D1D-5C08-482A-95EB-91CBD965011C", "學期成績通知單(固定排名)"));

          
        }


    }
}
