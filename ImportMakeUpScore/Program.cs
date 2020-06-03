using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Presentation;
using FISCA.Permission;

namespace ImportMakeUpScore
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            string ItemRegCode = "53D58211-8A4E-4795-9142-5B44FD5DD7B1";

            RibbonBarItem item = MotherForm.RibbonBarItems["班級", "資料統計"];
            item["報表"]["成績相關報表"]["補考成績匯入表"].Enable = false;

            item["報表"]["成績相關報表"]["補考成績匯入表"].Click += delegate
            {
                MainForm mf = new MainForm(K12.Presentation.NLDPanels.Class.SelectedSource);
                mf.ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && FISCA.Permission.UserAcl.Current[ItemRegCode].Executable)
                {
                    item["報表"]["成績相關報表"]["補考成績匯入表"].Enable = true;
                }
                else
                    item["報表"]["成績相關報表"]["補考成績匯入表"].Enable = false;
            };

            Catalog cal = RoleAclSource.Instance["班級"]["功能按鈕"];
            cal.Add(new RibbonFeature(ItemRegCode, "補考成績匯入表"));
        }

    }
}
