﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Permission;
using FISCA.Presentation;

namespace HsinChu.JHEvaluation
{
    class SetupEduAdmin
    {
        internal static void Init()
        {
            #region RibbonBar

            #region 教務作業/成績作業
            RibbonBarButton rbItem = JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["基本設定"]["設定"];
            //rbItem["評量設定"].Image = HsinChu.JHEvaluation.Properties.Resources.評量設定;
            rbItem["評分樣版設定"].Enable = Framework.User.Acl["JHSchool.EduAdmin.Ribbon0010"].Executable;
            rbItem["評分樣版設定"].Click += delegate
            {
                new HsinChu.JHEvaluation.CourseExtendControls.Ribbon.AssessmentSetupManager().ShowDialog();
            };

            RibbonBarButton rbItem1 = JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["批次作業/檢視"]["成績作業"];
            //rbItem1["評量輸入狀況檢視"].Image = Properties.Resources.成績輸入檢查;
            rbItem1["評量輸入狀況檢視"].Enable = Framework.User.Acl["JHSchool.EduAdmin.Ribbon0020"].Executable;
            rbItem1["評量輸入狀況檢視"].Click += delegate
            {
                new HsinChu.JHEvaluation.EduAdminExtendControls.Ribbon.CourseScoreStatusForm().ShowDialog();
            };
            #endregion

            #region 教務作業/基本設定/設定/評量成績缺考/免試
            var Permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            Permission.Add(new RibbonFeature("EduAdmin_Button_ScoreMap", "設定評量成績缺考/免試"));
            RibbonBarButton rbItem2 = JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["基本設定"]["設定"];
            rbItem2["評量成績缺考/免試"].Enable = UserAcl.Current["EduAdmin_Button_ScoreMap"].Executable;
            rbItem2["評量成績缺考/免試"].Click += delegate
            {
                new ConfigControls.Ribbon.ScoreValueManager().ShowDialog();
            };

            #endregion

            #region 教務作業/成績作業/評量缺免成績查詢調整
            var Permission2 = RoleAclSource.Instance["教務作業"];
            Permission2.Add(new RibbonFeature("EduAdmin_Button_CheckCourseScoreInput", "評量缺免成績查詢調整"));
            RibbonBarButton rbItem3 = JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["批次作業/檢視"]["成績作業"];
            rbItem3["評量缺免成績查詢調整"].Enable = UserAcl.Current["EduAdmin_Button_CheckCourseScoreInput"].Executable;
            rbItem3["評量缺免成績查詢調整"].Click += delegate
            {
                new EduAdminExtendControls.Ribbon.CheckCourseScoreInput().ShowDialog();
            };

            #endregion

            #endregion
        }
    }
}
