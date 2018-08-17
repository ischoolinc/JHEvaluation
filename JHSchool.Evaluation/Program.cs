using System.Collections.Generic;
using System.Windows.Forms;
using FISCA;
using FISCA.Presentation;
using Framework;
using Framework.Security;
using JHSchool.Affair;
using JHSchool.Evaluation.CourseExtendControls;
using JHSchool.Evaluation.CourseExtendControls.Ribbon;
using JHSchool.Evaluation.EduAdminExtendControls.Ribbon;
using System;
using DataRationality;
using K12.Presentation;
using FISCA.Data;
using K12.Data;
using FISCA.Authentication;
using FISCA.LogAgent;
using System.IO;

namespace JHSchool.Evaluation
{
    public static class Program
    {
        [MainMethod("JHSchool.Evaluation")]
        public static void Main()
        {
            #region SyncAllBackground
            //授課教師
            if (!TCInstruct.Instance.Loaded) TCInstruct.Instance.SyncAllBackground();
            //修課記錄
            //if (!SCAttend.Instance.Loaded) SCAttend.Instance.SyncAllBackground();
            //課程規劃
            if (!ProgramPlan.Instance.Loaded) ProgramPlan.Instance.SyncAllBackground();
            //計算規則
            if (!ScoreCalcRule.Instance.Loaded) ScoreCalcRule.Instance.SyncAllBackground();
            //評量設定
            if (!AssessmentSetup.Instance.Loaded) AssessmentSetup.Instance.SyncAllBackground();
            #endregion

            #region SetupPresentation
            //授課教師
            TCInstruct.Instance.SetupPresentation();
            //修課記錄
            SCAttend.Instance.SetupPresentation();
            //課程規劃
            ProgramPlan.Instance.SetupPresentation();
            //計算規則
            ScoreCalcRule.Instance.SetupPresentation();
            //評量設定
            AssessmentSetup.Instance.SetupPresentation();
            #endregion

            #region ContentItem 資料項目
            //學期成績
            //Student.Instance.AddDetailBulider(new DetailBulider<SemesterScoreItem>());
            //Student.Instance.AddDetailBulider(new DetailBulider<DLScoreItem>());
            //課程基本資訊
            Course.Instance.AddDetailBulider(new JHSchool.Legacy.ContentItemBulider<CourseExtendControls.BasicInfoItem>());

            //成績計算
            Course.Instance.AddDetailBulider(new DetailBulider<ScoreCalcSetupItem>());
            //修課學生

            Course.Instance.AddDetailBulider(new JHSchool.Legacy.ContentItemBulider<SCAttendItem>());
            //電子報表(因相關功能未完成先註)
            //Course.Instance.AddDetailBulider(new JHSchool.Legacy.ContentItemBulider<CourseExtendControls.ElectronicPaperItem>());
            //班級課程規劃
            Class.Instance.AddDetailBulider(new DetailBulider<ClassExtendControls.ClassBaseInfoItem_Extend>());

            //暫時註解
            //個人活動表現紀錄
            //Student.Instance.AddDetailBulider(new DetailBulider<StudentActivityRecordItem>());

            // 2017/5/9 穎驊  註解 下面位子 搬去 JHEvaluation.ScoreCalculation
            //畢業成績
            //Student.Instance.AddDetailBulider(new DetailBulider<GraduationScoreItem>());

            //修課及評量成績
            //Student.Instance.AddDetailBulider(new DetailBulider<CourseScoreItem>());
            #endregion

            #region RibbonBar

            NLDPanels.Student.RibbonBarItems["教務"]["成績作業"].Size = RibbonBarButton.MenuButtonSize.Large;
            NLDPanels.Student.RibbonBarItems["教務"]["成績作業"].Image = Properties.Resources.calc_save_64;

            NLDPanels.Student.RibbonBarItems["教務"]["畢業作業"].Size = RibbonBarButton.MenuButtonSize.Large;
            NLDPanels.Student.RibbonBarItems["教務"]["畢業作業"].Image = Properties.Resources.graduated_write_64;

            #region 學生/資料統計/報表
            RibbonBarButton rbButton = Student.Instance.RibbonBarItems["資料統計"]["報表"];

            // 2017/5/9 穎驊  註解 下面位子 搬去 JHEvaluation.ScoreCalculation
            //rbButton["成績相關報表"]["畢業預警報表"].Enable = User.Acl["JHSchool.Student.Report0010"].Executable;
            //rbButton["成績相關報表"]["畢業預警報表"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count <= 0) return;
            //    JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReport report = new JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReport(Student.Instance.SelectedList);
            //};
            //rbButton["學務相關報表"]["畢業預警報表"].Enable = User.Acl["JHSchool.Student.Report0010"].Executable;
            //rbButton["學務相關報表"]["畢業預警報表"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count <= 0) return;
            //    JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReport report = new JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReport(Student.Instance.SelectedList);
            //};
            #endregion

            #region 學生/資料統計/匯入匯出
            RibbonBarButton rbItemExport = Student.Instance.RibbonBarItems["資料統計"]["匯出"];
            RibbonBarButton rbItemImport = Student.Instance.RibbonBarItems["資料統計"]["匯入"];
            //rbItemExport["匯出學期科目成績"].Enable = User.Acl["JHSchool.Student.Ribbon0180"].Executable;
            //rbItemExport["匯出學期科目成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.ExportSemesterSubjectScore();
            //    JHSchool.Evaluation.ImportExport.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
            //    exporter.InitializeExport(wizard);
            //    wizard.ShowDialog();
            //};
            //rbItemExport["匯出學期領域成績"].Enable = User.Acl["JHSchool.Student.Ribbon0181"].Executable;
            //rbItemExport["匯出學期領域成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.ExportSemesterDomainScore();
            //    JHSchool.Evaluation.ImportExport.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
            //    exporter.InitializeExport(wizard);
            //    wizard.ShowDialog();
            //};
            //rbItemExport["匯出畢業成績"].Enable = User.Acl["JHSchool.Student.Ribbon0182"].Executable;
            //rbItemExport["匯出畢業成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.ExportGradScore();
            //    JHSchool.Evaluation.ImportExport.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
            //    exporter.InitializeExport(wizard);
            //    wizard.ShowDialog();
            //};
            rbItemExport["成績相關匯出"]["匯出課程成績"].Enable = User.Acl["JHSchool.Student.Ribbon0183"].Executable;
            rbItemExport["成績相關匯出"]["匯出課程成績"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.ExportCourseScore();
                JHSchool.Evaluation.ImportExport.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            rbItemExport["成績相關匯出"]["匯出學期歷程"].Enable = User.Acl["JHSchool.Student.Ribbon0169"].Executable;
            rbItemExport["成績相關匯出"]["匯出學期歷程"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.ExportSemesterHistory();
                JHSchool.Evaluation.ImportExport.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            //rbItemExport["匯出評量成績"].Enable = User.Acl["JHSchool.Student.Ribbon0184"].Executable;
            //rbItemExport["匯出評量成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.ExportExamScore();
            //    JHSchool.Evaluation.ImportExport.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
            //    exporter.InitializeExport(wizard);
            //    wizard.ShowDialog();
            //};

            //rbItemImport["匯入學期科目成績"].Enable = User.Acl["JHSchool.Student.Ribbon0190"].Executable;
            //rbItemImport["匯入學期科目成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.ImportSemesterSubjectScore();
            //    JHSchool.Evaluation.ImportExport.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ImportStudentV2(importer.Text, importer.Image);
            //    importer.InitializeImport(wizard);
            //    wizard.ShowDialog();
            //};
            //rbItemImport["匯入學期領域成績"].Enable = User.Acl["JHSchool.Student.Ribbon0191"].Executable;
            //rbItemImport["匯入學期領域成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.ImportSemesterDomainScore();
            //    JHSchool.Evaluation.ImportExport.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ImportStudentV2(importer.Text, importer.Image);
            //    importer.InitializeImport(wizard);
            //    wizard.ShowDialog();
            //};
            //rbItemImport["匯入畢業成績"].Enable = User.Acl["JHSchool.Student.Ribbon0192"].Executable;
            //rbItemImport["匯入畢業成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.ImportGradScore();
            //    JHSchool.Evaluation.ImportExport.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ImportStudentV2(importer.Text, importer.Image);
            //    importer.InitializeImport(wizard);
            //    wizard.ShowDialog();
            //};
            rbItemImport["成績相關匯入"]["匯入課程成績"].Enable = User.Acl["JHSchool.Student.Ribbon0193"].Executable;
            rbItemImport["成績相關匯入"]["匯入課程成績"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.ImportCourseScore();
                JHSchool.Evaluation.ImportExport.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };

            rbItemImport["成績相關匯入"]["匯入學期歷程"].Enable = User.Acl["JHSchool.Student.Ribbon0170"].Executable;
            rbItemImport["成績相關匯入"]["匯入學期歷程"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.ImportSemesterHistory();
                JHSchool.Evaluation.ImportExport.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };

            //rbItemImport["匯入評量成績"].Enable = User.Acl["JHSchool.Student.Ribbon0194"].Executable;
            //rbItemImport["匯入評量成績"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.ImportExamScore();
            //    JHSchool.Evaluation.ImportExport.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.ImportStudentV2(importer.Text, importer.Image);
            //    importer.InitializeImport(wizard);
            //    wizard.ShowDialog();
            //};
            #endregion

            #region 學生/教務
            //移到成績計算模組。
            RibbonBarItem rbItem = Student.Instance.RibbonBarItems["教務"];
            //rbItem["畢業資格審查"].Enable = User.Acl["JHSchool.Student.Ribbon0058"].Executable;
            //rbItem["畢業資格審查"].Image = JHSchool.Evaluation.CourseExtendControls.Ribbon.Resources.graduation_64;
            //rbItem["畢業資格審查"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count == 0) return;
            //    Form form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationInspectWizard("Student");
            //    form.ShowDialog();
            //};
            #endregion

            #region 學生/成績/排名
            /** 程式碼移動到 JHEvaluation.Rating Module 中。*/

            //rbButton = Student.Instance.RibbonBarItems["成績"]["排名"];
            //rbButton.Enable = User.Acl["JHSchool.Student.Ribbon0059"].Executable;
            //rbButton["評量科目成績排名"].Enable = true;
            //rbButton["評量科目成績排名"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count > 0)
            //    {
            //        JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormExamSubject form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormExamSubject();
            //        form.SetRatingStudents(K12.Presentation.NLDPanels.Student.SelectedSource);
            //        form.ShowDialog();
            //    }
            //};

            ////rbButton.Enable = User.Acl[""].Executable;
            //rbButton["評量領域成績排名"].Enable = true;
            //rbButton["評量領域成績排名"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count > 0)
            //    {
            //        JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormExamDomain form= new JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormExamDomain();
            //        form.SetRatingStudents(K12.Presentation.NLDPanels.Student.SelectedSource);
            //        form.ShowDialog();
            //    }
            //};

            //rbButton["學期科目成績排名"].Enable = true;
            //rbButton["學期科目成績排名"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count > 0)
            //    {
            //        JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemesterSubject form= new JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemesterSubject();
            //        form.SetRatingStudents(K12.Presentation.NLDPanels.Student.SelectedSource);
            //        form.ShowDialog();
            //    }
            //};

            //rbButton["學期領域成績排名"].Enable = true;
            //rbButton["學期領域成績排名"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count > 0)
            //    {
            //        JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemesterDomain form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemesterDomain();
            //        form.SetRatingStudents(K12.Presentation.NLDPanels.Student.SelectedSource);
            //        form.ShowDialog();
            //    }
            //};

            //rbButton["學期科目成績排名(多學期)"].Enable = true;
            //rbButton["學期科目成績排名(多學期)"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count > 0)
            //    {
            //        JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemestersSubject form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemestersSubject();
            //        form.SetRatingStudents(K12.Presentation.NLDPanels.Student.SelectedSource);
            //        form.ShowDialog();
            //    }
            //};

            //rbButton["學期領域成績排名(多學期)"].Enable = true;
            //rbButton["學期領域成績排名(多學期)"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count > 0)
            //    {
            //        JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemestersDomain form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormSemestersDomain();
            //        form.SetRatingStudents(K12.Presentation.NLDPanels.Student.SelectedSource);
            //        form.ShowDialog();
            //    }
            //};

            //rbButton["畢業成績排名"].Enable = true;
            //rbButton["畢業成績排名"].Click += delegate
            //{
            //    if (Student.Instance.SelectedList.Count > 0)
            //    {
            //        Form form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.Rating.FormGraduation();
            //        form.ShowDialog();
            //    }
            //};
            #endregion

            #region 班級/資料統計/報表
            MenuButton semsGoodStudentReportButton = Class.Instance.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["學期優異表現名單"];
            semsGoodStudentReportButton.Enable = User.Acl["JHSchool.Class.Report0180"].Executable;
            semsGoodStudentReportButton.Click += delegate
            {
                if (Class.Instance.SelectedList.Count > 0)
                {
                    Form form = new JHSchool.Evaluation.ClassExtendControls.Ribbon.Score.SemsGoodStudentReport.SemsGoodStudentReport();
                    form.ShowDialog();
                }
            };
            #endregion

            #region 班級/教務
            rbButton = K12.Presentation.NLDPanels.Class.RibbonBarItems["教務"]["班級開課"];
            //rbButton = Class.Instance.RibbonBarItems["成績"]["班級開課"];
            rbButton.Enable = User.Acl["JHSchool.Class.Ribbon0070"].Executable;
            rbButton.Image = Properties.Resources.organigram_refresh_64;
            rbButton["依課程規劃表開課"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                    JHSchool.Evaluation.ClassExtendControls.Ribbon.CreateCoursesByProgramPlan.Run();
            };
            rbButton["直接開課"].Click += delegate
            {
                if (Class.Instance.SelectedList.Count > 0) new JHSchool.Evaluation.ClassExtendControls.Ribbon.CreateCoursesDirectly();
            };
            #endregion

            #region 課程/編輯
            rbItem = Course.Instance.RibbonBarItems["編輯"];
            rbButton = rbItem["新增"];
            rbButton.Size = RibbonBarButton.MenuButtonSize.Large;
            rbButton.Image = JHSchool.Evaluation.CourseExtendControls.Ribbon.Resources.btnAddCourse;
            rbButton.Enable = User.Acl["JHSchool.Course.Ribbon0000"].Executable;
            rbButton.Click += delegate
            {
                new JHSchool.Evaluation.CourseExtendControls.Ribbon.AddCourse().ShowDialog();
            };

            rbButton = rbItem["刪除"];
            rbButton.Size = RibbonBarButton.MenuButtonSize.Large;
            rbButton.Image = JHSchool.Evaluation.CourseExtendControls.Ribbon.Resources.btnDeleteCourse;
            rbButton.Enable = false;
            rbButton.Click += delegate
            {
                //2018/3/26 穎驊註記 下列為舊的 刪除課程，一次僅能一筆，且如果內有資料，則無法刪除，也不會有Log 紀錄
                //if (Course.Instance.SelectedKeys.Count == 1)
                //{
                //    JHSchool.Data.JHCourseRecord record = JHSchool.Data.JHCourse.SelectByID(Course.Instance.SelectedKeys[0]);
                //    //int CourseAttendCot = Course.Instance.Items[record.ID].GetAttendStudents().Count;
                //    List<JHSchool.Data.JHSCAttendRecord> scattendList = JHSchool.Data.JHSCAttend.SelectByStudentIDAndCourseID(new List<string>() { }, new List<string>() { record.ID });
                //    int attendStudentCount = 0;
                //    foreach (JHSchool.Data.JHSCAttendRecord scattend in scattendList)
                //    {
                //        if (scattend.Student.Status == K12.Data.StudentRecord.StudentStatus.一般)
                //            attendStudentCount++;
                //    }

                //    if (attendStudentCount > 0)
                //        MsgBox.Show(record.Name + " 有" + attendStudentCount.ToString() + "位修課學生，請先移除修課學生後再刪除課程.");
                //    else
                //    {
                //        string msg = string.Format("確定要刪除「{0}」？", record.Name);
                //        if (MsgBox.Show(msg, "刪除課程", MessageBoxButtons.YesNo) == DialogResult.Yes)
                //        {
                //            #region 自動刪除非一般學生的修課記錄
                //            List<JHSchool.Data.JHSCAttendRecord> deleteSCAttendList = new List<JHSchool.Data.JHSCAttendRecord>();
                //            foreach (JHSchool.Data.JHSCAttendRecord scattend in scattendList)
                //            {
                //                JHSchool.Data.JHStudentRecord stuRecord = JHSchool.Data.JHStudent.SelectByID(scattend.RefStudentID);
                //                if (stuRecord == null) continue;
                //                if (stuRecord.Status != K12.Data.StudentRecord.StudentStatus.一般)
                //                    deleteSCAttendList.Add(scattend);
                //            }
                //            List<string> studentIDs = new List<string>();
                //            foreach (JHSchool.Data.JHSCAttendRecord scattend in deleteSCAttendList)
                //                studentIDs.Add(scattend.RefStudentID);
                //            List<JHSchool.Data.JHSCETakeRecord> sceList = JHSchool.Data.JHSCETake.SelectByStudentAndCourse(studentIDs, new List<string>() { record.ID });
                //            JHSchool.Data.JHSCETake.Delete(sceList);
                //            JHSchool.Data.JHSCAttend.Delete(deleteSCAttendList);
                //            #endregion

                //            JHSchool.Data.JHCourse.Delete(record);
                //            //CourseRecordEditor crd = Course.Instance.Items[record.ID].GetEditor();
                //            //crd.Remove = true;
                //            //crd.Save();
                //            // 加這主要是重新整理
                //            Course.Instance.SyncDataBackground(record.ID);
                //        }
                //        else
                //            return;
                //    }
                //}

                if (MsgBox.Show("本功能將會將所選全部課程移除，請問是否繼續?", "批次刪除課程", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    List<JHSchool.Data.JHCourseRecord> record_list = JHSchool.Data.JHCourse.SelectByIDs(Course.Instance.SelectedKeys);

                    //課程資料
                    List<string> courseIDList = new List<string>();

                    foreach (JHSchool.Data.JHCourseRecord r in record_list)
                    {
                        courseIDList.Add(r.ID);
                    }

                    // 學生參與課堂紀錄
                    List<JHSchool.Data.JHSCAttendRecord> scattendList = JHSchool.Data.JHSCAttend.SelectByStudentIDAndCourseID(new List<string>() { }, courseIDList);

                    List<string> scattendIDList = new List<string>();
                    List<string> studentIDList = new List<string>();

                    foreach (JHSchool.Data.JHSCAttendRecord r in scattendList)
                    {
                        scattendIDList.Add(r.ID);
                        studentIDList.Add(r.RefStudentID);
                    }

                    //學生課堂考試紀錄
                    List<JHSchool.Data.JHSCETakeRecord> sceList = JHSchool.Data.JHSCETake.SelectByStudentAndCourse(new List<string>() { }, courseIDList);

                    Dictionary<string, List<JHSchool.Data.JHSCETakeRecord>> studentScoreDict = new Dictionary<string, List<Data.JHSCETakeRecord>>();

                    //用RefStudentID 整理個學生的成績
                    foreach (JHSchool.Data.JHSCETakeRecord jhscetr in sceList)
                    {
                        if (!studentScoreDict.ContainsKey(jhscetr.RefStudentID))
                        {
                            studentScoreDict.Add(jhscetr.RefStudentID, new List<Data.JHSCETakeRecord>());
                            studentScoreDict[jhscetr.RefStudentID].Add(jhscetr);
                        }
                        else
                        {
                            studentScoreDict[jhscetr.RefStudentID].Add(jhscetr);
                        }
                    }
                    List<string> WarningList = new List<string>();

                    System.Text.StringBuilder sb1 = new System.Text.StringBuilder();
                    foreach (KeyValuePair<string, List<JHSchool.Data.JHSCETakeRecord>> p in studentScoreDict)
                    {
                        sb1.AppendLine("學生:" + p.Value[0].Student.Name + ",於");

                        foreach (JHSchool.Data.JHSCETakeRecord r in p.Value)
                        {
                            sb1.AppendLine("課程:" + r.Course.Name + ",試別:" + r.Exam.Name + ",含有分數:" + r.Score + "資料");
                        }

                        WarningList.Add(sb1.ToString());

                        sb1.Clear();
                    }

                    string warning_total = string.Format("發現「{0}」位學生在課程具有成績資料，請問是否仍要刪除？", WarningList.Count);

                    System.Text.StringBuilder sb2 = new System.Text.StringBuilder();
                    sb2.AppendLine(warning_total);

                    foreach (string s in WarningList)
                    {
                        sb2.AppendLine(s);
                    }
                    
                    // 假如刪除的學生中，有人有成績資料，則提醒使用者是否要刪掉
                    if (sceList.Count > 0)
                    {
                        DeleteCourseWarningForm dcwf = new DeleteCourseWarningForm(warning_total, studentScoreDict);
                        
                        if (dcwf.ShowDialog() == DialogResult.Yes)
                        {
                            DeleteCourseWithLogSQL();

                        }

                        //if (MsgBox.Show(sb2.ToString(), "批次刪除課程", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        //{
                        //    DeleteCourseWithLogSQL();
                        //}
                    }
                    else
                    {
                        DeleteCourseWithLogSQL();
                    }
                }
            };



            //2018/3/26 穎驊註解，將此功能一併併入刪除處理
            //////2018/3/9 穎驊註解，此為因應高雄小組項目 [09-05][04] 批次刪除課程功能 所新增的功能
            //rbButton = rbItem["批次刪除修課學生"];
            //rbButton.Size = RibbonBarButton.MenuButtonSize.Large;
            //rbButton.Image = JHSchool.Evaluation.CourseExtendControls.Ribbon.Resources.btnDeleteStudent_Image;
            ////2018/3/9 穎驊註解，暫時找不到期註冊開放權限的Code，先跟隨 刪除課程的設定
            //rbButton.Enable = false;
            //rbButton.Click += delegate
            //{
            //    if (MsgBox.Show("本功能將會將所選課程的全部修課學生移除，請問是否繼續?", "批次刪除修課學生", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //    {
            //        List<JHSchool.Data.JHCourseRecord> record_list = JHSchool.Data.JHCourse.SelectByIDs(Course.Instance.SelectedKeys);

            //        //課程資料
            //        List<string> courseIDList = new List<string>();

            //        foreach (JHSchool.Data.JHCourseRecord r in record_list)
            //        {
            //            courseIDList.Add(r.ID);
            //        }

            //        // 學生參與課堂紀錄
            //        List<JHSchool.Data.JHSCAttendRecord> scattendList = JHSchool.Data.JHSCAttend.SelectByStudentIDAndCourseID(new List<string>() { }, courseIDList);

            //        List<string> scattendIDList = new List<string>();

            //        foreach (JHSchool.Data.JHSCAttendRecord r in scattendList)
            //        {
            //            scattendIDList.Add(r.ID);
            //        }

            //        //學生課堂考試紀錄
            //        List<JHSchool.Data.JHSCETakeRecord> sceList = JHSchool.Data.JHSCETake.SelectByIDs(scattendIDList);


            //        // 假如刪除的學生中，有人有成績資料，則提醒使用者是否要刪掉
            //        if (sceList.Count > 0)
            //        {
            //            if (MsgBox.Show("發現有學生在課程具有成績資料，請問是否仍要刪除?", "批次刪除修課學生", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //            {


            //            }
            //        }
            //    }
            //};

            RibbonBarButton CouItem1 = Course.Instance.RibbonBarItems["編輯"]["刪除"];
            //RibbonBarButton CouItem2 = Course.Instance.RibbonBarItems["編輯"]["批次刪除修課學生"];
            Course.Instance.SelectedListChanged += delegate
            {
                // 課程刪除不能多選
                //CouItem1.Enable = (Course.Instance.SelectedList.Count < 2) && User.Acl["JHSchool.Course.Ribbon0010"].Executable;
                CouItem1.Enable = User.Acl["JHSchool.Course.Ribbon0010"].Executable;

                ////2018/3/9 穎驊註解，批次刪除學生至少要選一
                //CouItem2.Enable = (Course.Instance.SelectedList.Count > 0) && User.Acl["JHSchool.Course.Ribbon0010"].Executable;
            };
            #endregion



            #region 課程/資料統計/報表
            rbButton = Course.Instance.RibbonBarItems["資料統計"]["報表"];
            rbButton["學生修課清單"].Enable = User.Acl["JHSchool.Course.Report0000"].Executable;
            rbButton["學生修課清單"].Click += delegate
            {
                if (Course.Instance.SelectedList.Count >= 1)
                {
                    new JHSchool.Evaluation.ClassExtendControls.Ribbon.StuinCourse.StuinCourse();
                }
                else
                    MsgBox.Show("請選擇課程");
            };
            #endregion

            #region 課程/資料統計/匯入匯出
            RibbonBarItem rbItemCourseImportExport = Course.Instance.RibbonBarItems["資料統計"];
            rbItemCourseImportExport["匯出"]["匯出課程修課學生"].Enable = User.Acl["JHSchool.Course.Ribbon0031"].Executable;
            rbItemCourseImportExport["匯出"]["匯出課程修課學生"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.Course.ExportCourseStudents("");
                JHSchool.Evaluation.ImportExport.Course.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.Course.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };
            rbItemCourseImportExport["匯入"]["匯入課程修課學生"].Enable = User.Acl["JHSchool.Course.Ribbon0021"].Executable;
            rbItemCourseImportExport["匯入"]["匯入課程修課學生"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.Course.ImportCourseStudents("");
                JHSchool.Evaluation.ImportExport.Course.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.Course.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };

            //rbItemCourseImportExport = FISCA.Presentation.MotherForm.RibbonBarItems["社團作業", "資料統計"];
            //rbItemCourseImportExport["匯出"]["匯出社團參與學生"].Enable = User.Acl["JHSchool.Course.Ribbon0031"].Executable;
            //rbItemCourseImportExport["匯出"]["匯出社團參與學生"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Evaluation.ImportExport.Course.ExportCourseStudents("社團");
            //    JHSchool.Evaluation.ImportExport.Course.ExportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.Course.ExportStudentV2(exporter.Text, exporter.Image, "社團", K12.Presentation.NLDPanels.Course.SelectedSource);
            //    exporter.InitializeExport(wizard);
            //    wizard.ShowDialog();
            //};
            //rbItemCourseImportExport["匯入"]["匯入社團參與學生"].Enable = User.Acl["JHSchool.Course.Ribbon0021"].Executable;
            //rbItemCourseImportExport["匯入"]["匯入社團參與學生"].Click += delegate
            //{
            //    SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Evaluation.ImportExport.Course.ImportCourseStudents("社團");
            //    JHSchool.Evaluation.ImportExport.Course.ImportStudentV2 wizard = new JHSchool.Evaluation.ImportExport.Course.ImportStudentV2(importer.Text, importer.Image);
            //    importer.InitializeImport(wizard);
            //    wizard.ShowDialog();
            //};
            #endregion

            #region 課程/教務
            RibbonBarButton group = Course.Instance.RibbonBarItems["教務"]["分組上課"];
            group.Size = RibbonBarButton.MenuButtonSize.Medium;
            group.Image = Properties.Resources.meeting_refresh_64;
            group.Enable = User.Acl["JHSchool.Course.Ribbon0060"].Executable;
            group.Click += delegate
            {
                if (Course.Instance.SelectedList.Count > 0)
                    new JHSchool.Evaluation.Legacy.SwapAttendStudents(Course.Instance.SelectedList.Count).ShowDialog();
            };

            ////課程超過7項,則"分組上課"不能點擊
            //Course.Instance.SelectedListChanged += delegate
            //{
            //    //分組上課不能超過七個課程。
            //    group.Enable = (Course.Instance.SelectedList.Count <= 7) && User.Acl["JHSchool.Course.Ribbon0060"].Executable;
            //};

            RibbonBarItem scores = Course.Instance.RibbonBarItems["教務"];
            //scores["成績輸入"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //scores["成績輸入"].Image = Resources.exam_write_64;
            //scores["成績輸入"].Enable = User.Acl["JHSchool.Course.Ribbon0070"].Executable;
            //scores["成績輸入"].Click += delegate
            //{
            //    if (Course.Instance.SelectedList.Count == 1)
            //    {
            //        //new JHSchool.Evaluation.CourseExtendControls.Ribbon.EditCourseScore(Course.Instance.SelectedList[0]).ShowDialog();
            //        CourseRecord courseRecord = Course.Instance.SelectedList[0];
            //        if (courseRecord.GetAssessmentSetup() == null)
            //            MsgBox.Show("課程 '" + courseRecord.Name + "' 沒有評量設定。");
            //        else
            //            new JHSchool.Evaluation.CourseExtendControls.Ribbon.CourseScoreInputForm(courseRecord).ShowDialog();
            //    }
            //};
            //Course.Instance.SelectedListChanged += delegate
            //{
            //    scores["成績輸入"].Enable = Course.Instance.SelectedList.Count == 1 && User.Acl["JHSchool.Course.Ribbon0070"].Executable;
            //};

            //scores["成績計算"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //scores["成績計算"].Image = Resources.calcScore;
            //scores["成績計算"].Enable = User.Acl["JHSchool.Course.Ribbon0080"].Executable;
            //scores["成績計算"].Click += delegate
            //{
            //    new JHSchool.Evaluation.CourseExtendControls.Ribbon.CalculateionWizard().ShowDialog();
            //};

            // UNDONE: 等有討論出什麼結論再說吧…
            //scores["合科什麼鬼的"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //scores["合科什麼鬼的"].Enable = true;
            //scores["合科什麼鬼的"].Click += delegate
            //{
            //    List<JHSchool.Data.JHCourseRecord> courseList = JHSchool.Data.JHCourse.SelectByIDs(K12.Presentation.NLDPanels.Course.SelectedSource);
            //    if (courseList.Count <= 0) return;

            //    if (SubjectCombinationConfigForm.CheckAssessmentSetup(courseList) == true)
            //    {
            //        new SubjectCombinationConfigForm(courseList).ShowDialog();
            //    }
            //};
            #endregion

            #region 教務作業/課務作業

            rbItem = EduAdmin.Instance.RibbonBarItems["基本設定"];

            rbItem["管理"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem["管理"].Image = Properties.Resources.network_lock_64;
            rbItem["管理"]["領域資料管理"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon.DomainList"].Executable;
            rbItem["管理"]["領域資料管理"].Click += delegate
            {
                new DomainListTable().ShowDialog();
            };

            rbItem["管理"]["科目資料管理"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon.DomainList"].Executable;
            rbItem["管理"]["科目資料管理"].Click += delegate
            {
                new SubjectListTable().ShowDialog();
            };

            //rbItem["管理"]["評量名稱管理"].Image = Resources.評量名稱管理;
            rbItem["管理"]["評量名稱管理"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0000"].Executable;
            rbItem["管理"]["評量名稱管理"].Click += delegate
            {
                new JHSchool.Evaluation.CourseExtendControls.Ribbon.ExamManager().ShowDialog();
            };

            //rbItem["等第對照表"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //rbItem["等第對照表"].Image = Resources.對照表;
            rbItem["管理"]["等第對照管理"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0031"].Executable;
            rbItem["管理"]["等第對照管理"].Click += delegate
            {
                new JHSchool.Evaluation.EduAdminExtendControls.Ribbon.ScoreMappingTable().ShowDialog();
            };
            #endregion

            #region 教務作業/成績作業
            rbItem["對照/代碼"].Image = Properties.Resources.notepad_lock_64;
            rbItem["對照/代碼"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem["對照/代碼"]["文字描述代碼表"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0032"].Executable;
            rbItem["對照/代碼"]["文字描述代碼表"].Click += delegate
            {
                TextMappingTable text = new TextMappingTable();
                text.ShowDialog();
            };

            //rbItem["努力程度對照表"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //rbItem["努力程度對照表"].Image = Resources.對照表;
            rbItem["對照/代碼"]["努力程度代碼表"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0030"].Executable;
            rbItem["對照/代碼"]["努力程度代碼表"].Click += delegate
            {
                new JHSchool.Evaluation.EduAdminExtendControls.Ribbon.EffortDegreeTable().ShowDialog();
            };
            EffortDegreeTable.CheckDefault();

            //如果需要此分類再進行應用
            rbItem["設定"].Image = Properties.Resources.sandglass_unlock_64;
            rbItem["設定"].Size = RibbonBarButton.MenuButtonSize.Large;

            ScoreMappingTable.CheckDefault();

            //rbItem["課程規劃表"].Size = RibbonBarButton.MenuButtonSize.Large;
            //rbItem["課程規劃表"].Image = ClassExtendControls.Ribbon.Resources.btnProgramPlan_Image;
            rbItem["設定"]["課程規劃表"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0050"].Executable;
            rbItem["設定"]["課程規劃表"].Click += delegate
            {
                new ProgramPlanManager().ShowDialog();
            };

            //rbItem["成績計算規則"].Size = RibbonBarButton.MenuButtonSize.Large;
            //rbItem["成績計算規則"].Image = JHSchool.Evaluation.CourseExtendControls.Ribbon.Resources.course_plan;
            rbItem["設定"]["成績計算規則"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0040"].Executable;
            rbItem["設定"]["成績計算規則"].Click += delegate
            {
                //if (Control.ModifierKeys == Keys.Shift)
                new JHSchool.Evaluation.EduAdminExtendControls.Ribbon.ScoreCalcRuleManager().ShowDialog();
            };


            //移到成績計算模組。
            //rbItem["畢業資格審查"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0046"].Executable;
            //rbItem["畢業資格審查"].Image = JHSchool.Evaluation.CourseExtendControls.Ribbon.Resources.graduation_64;
            //rbItem["畢業資格審查"].Click += delegate
            //{
            //    Form form = new JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationInspectWizard("EduAdmin");
            //    form.ShowDialog();
            //};

            //rbItem["評量輸入狀況"].Image = Resources.成績輸入檢查;
            //rbItem["評量輸入狀況"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0020"].Executable;
            //rbItem["評量輸入狀況"].Click += delegate
            //{
            //    new JHSchool.Evaluation.EduAdminExtendControls.Ribbon.CourseScoreStatusForm().ShowDialog();
            //};

            //rbItem["評量設定"].Image = Resources.評量設定;
            //rbItem["評量設定"].Enable = User.Acl["JHSchool.EduAdmin.Ribbon0010"].Executable;
            //rbItem["評量設定"].Click += delegate
            //{
            //    new JHSchool.Evaluation.CourseExtendControls.Ribbon.AssessmentSetupManager().ShowDialog();
            //};

            //rbItem["特殊教育領域設定"].Image = Resources.;

            //rbItem["特殊教育領域設定"].Enable = true;
            //rbItem["特殊教育領域設定"].Click += delegate
            //{
            //    new SpecialEduDomainTable().ShowDialog();
            //};
            #endregion

            #endregion

            //--------------------------------下面還沒整理的分隔線--------------------------------

            #region 註冊權限管理

            //學生
            Catalog detail = RoleAclSource.Instance["學生"]["資料項目"];

            // 2017/5/9 穎驊  註解 下面位子 搬去 JHEvaluation.ScoreCalculation
            //detail.Add(new DetailItemFeature(typeof(SemesterScoreItem)));
            //detail.Add(new DetailItemFeature(typeof(GraduationScoreItem)));
            //detail.Add(new DetailItemFeature(typeof(CourseScoreItem)));

            Catalog ribbon = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0055", "課程規劃"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0056", "計算規則"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0057", "計算成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0058", "畢業資格審查"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0059", "排名")); //程式碼移動到 JHEvaluation.Rating Module 中。

            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0180", "匯出學期科目成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0181", "匯出學期領域成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0182", "匯出畢業成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0183", "匯出課程成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0184", "匯出評量成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0169", "匯出學期歷程"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0190", "匯入學期科目成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0191", "匯入學期領域成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0192", "匯入畢業成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0193", "匯入課程成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0194", "匯入評量成績"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0170", "匯入學期歷程"));

            ribbon = RoleAclSource.Instance["學生"]["報表"];
            ribbon.Add(new ReportFeature("JHSchool.Student.Report0010", "畢業預警報表"));

            //班級
            ribbon = RoleAclSource.Instance["班級"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("JHSchool.Class.Ribbon0055", "課程規劃"));
            ribbon.Add(new RibbonFeature("JHSchool.Class.Ribbon0056", "計算規則"));
            ribbon.Add(new RibbonFeature("JHSchool.Class.Ribbon0070", "班級開課"));

            ribbon = RoleAclSource.Instance["班級"]["報表"];
            ribbon.Add(new RibbonFeature("JHSchool.Class.Report0180", "學期優異表現名單"));

            //課程
            ribbon = RoleAclSource.Instance["課程"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("JHSchool.Course.Ribbon0031", "匯出課程修課學生"));
            ribbon.Add(new RibbonFeature("JHSchool.Course.Ribbon0021", "匯入課程修課學生"));
            ribbon.Add(new RibbonFeature("JHSchool.Course.Ribbon.AssignAssessmentSetup", "評量設定")); //增加權限控管 by dylan(2010/11/25)

            detail = RoleAclSource.Instance["課程"]["資料項目"];
            detail.Add(new DetailItemFeature(typeof(CourseExtendControls.BasicInfoItem)));
            detail.Add(new DetailItemFeature(typeof(CourseExtendControls.ScoreCalcSetupItem)));
            detail.Add(new DetailItemFeature(typeof(CourseExtendControls.SCAttendItem)));

            // //電子報表(因相關功能未完成先註)
            //detail.Add(new DetailItemFeature(typeof(CourseExtendControls.ElectronicPaperItem)));

            //教務作業
            ribbon = RoleAclSource.Instance["教務作業"];
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon.DomainList", "領域清單"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon.SubjectList", "科目清單"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0000", "評量名稱管理"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0010", "評量設定"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0020", "評量輸入狀況"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0030", "努力程度對照表"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0031", "等第對照表"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0032", "文字描述代碼表"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0040", "成績計算規則"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0045", "計算成績"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0046", "畢業資格審查"));
            ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0050", "課程規劃表"));

            //建文的舊功能
            //ribbon = RoleAclSource.Instance["學務作業"];
            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0090.2", "幹部名稱管理"));
            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0091", "班級幹部管理"));
            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0092", "社團幹部管理"));
            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0093", "學校幹部管理"));
            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0094", "競賽項目管理"));
            #endregion

            Domain.TestDrive();
            Course.Instance.AddView(new TeacherCategoryView());

            // 2017/5/9 穎驊  註解 下面位子 搬去 JHEvaluation.ScoreCalculation
            // 學生學期歷程與學期成績學年度學期檢查
            //DataRationalityManager.Checks.Add(new StudentExtendControls.Ribbon.CheckStudentSemHistoryScoreRAT());

        }


        //2018/3/23  穎驊備註 ，以下超大包SQL 是參考 弈均所開發的選課模組 github\SHSchool.CourseSelection\SHSchool.CourseSelection\Forms\AdjustSSAttendForm.cs
        // 由恩正所教的WITH AS SQL 寫法，其好處是整包的SQL 會一併執行完，或是全部不執行，
        //如此可以避免 利用ischool API分段 處理資料，可能會造成資料不同步的問題，並且可以一併寫入log
        private static void DeleteCourseWithLogSQL()
        {
            string _actor = DSAServices.UserAccount; ;
            string _client_info = ClientInfo.GetCurrentClientInfo().OutputResult().OuterXml;

            // 兜資料
            List<string> dataList = new List<string>();
            foreach (string id in Course.Instance.SelectedKeys)
            {                
                string data = string.Format(@"
                SELECT
                    {0}::BIGINT AS id
                ", id);

                dataList.Add(data);
            }

            string Data = string.Join(" UNION ALL", dataList);


            string sql = string.Format(@"WITH course_data_row AS(
			 {0}
) ,delete_sc_attend_data AS(
	SELECT 
		sc_attend.id AS sc_attend_id
	FROM 
		course_data_row
		INNER JOIN sc_attend
			ON sc_attend.ref_course_id = course_data_row.id		
) ,delete_sce_take_data AS(
	SELECT 
		sce_take.id AS sce_take_id
	FROM 
		delete_sc_attend_data
		INNER JOIN sce_take 
			ON sce_take.ref_sc_attend_id = delete_sc_attend_data.sc_attend_id		
) ,delete_sce_take AS(
	DELETE
	FROM
		sce_take
	WHERE sce_take.id IN (
		SELECT delete_sce_take_data.sce_take_id
		FROM delete_sce_take_data
		LEFT OUTER JOIN sce_take 
			ON delete_sce_take_data.sce_take_id = sce_take.id
		)
	RETURNING sce_take.*
),delete_sc_attend AS(
	DELETE
	FROM
		sc_attend
	WHERE sc_attend.id IN (
		SELECT delete_sc_attend_data.sc_attend_id
		FROM delete_sc_attend_data
		LEFT OUTER JOIN sc_attend 
			ON delete_sc_attend_data.sc_attend_id = sc_attend.id
		)
	RETURNING sc_attend.*
),delete_course AS(
	DELETE
	FROM
		course
	WHERE course.id IN (
		SELECT course_data_row.id
		FROM course_data_row
		)
	RETURNING course.*
),insert_sce_take_log_student_data AS(
INSERT INTO log(
	actor
	, action_type
	, action
	, target_category
	, target_id
	, server_time
	, client_info
	, action_by
	, description
)
SELECT 
	'{1}'::TEXT AS actor
	, 'Record' AS action_type
	, '課程_刪除' AS action
	, 'student'::TEXT AS target_category
	, student.id AS target_id
	, now() AS server_time
	, '{2}' AS client_info
	, '刪除_學生_課程評量成績'AS action_by   
	, '學生「'|| student.name ||'」，課程「'|| course.course_name || '」修課紀錄，評量成績刪除，刪除內容:「'|| sce_take.extension ||'」 使用者「{1}」，' AS description 
FROM
	delete_sce_take_data
	LEFT OUTER JOIN sce_take ON sce_take.id = delete_sce_take_data.sce_take_id 	
	LEFT OUTER JOIN sc_attend ON sc_attend.id = sce_take.ref_sc_attend_id
	LEFT OUTER JOIN course ON course.id = sc_attend.ref_course_id 	
	LEFT OUTER JOIN student ON student.id = sc_attend.ref_student_id

)INSERT INTO log(
	actor
	, action_type
	, action
	, target_category
	, target_id
	, server_time
	, client_info
	, action_by
	, description
)
SELECT 
	'{1}'::TEXT AS actor
	, 'Record' AS action_type
	, '課程_刪除' AS action
	, 'course'::TEXT AS target_category
	, course.id AS target_id
	, now() AS server_time
	, '{2}' AS client_info
	, '刪除_課程_課程'AS action_by   
	, '課程「'|| course.course_name || '」，課程刪除，包含其所有修課學生修課紀錄、評量成績，使用者「{1}」，' AS description 
FROM
	course_data_row
	LEFT OUTER JOIN course ON course.id = course_data_row.id
                ", Data,_actor,_client_info);

            
            UpdateHelper uh = new UpdateHelper();

            //執行sql
            uh.Execute(sql);

            //2018/3/26 穎驊註記，此段會將上面一包SQL 檔 匯出至debug 資料夾。
            //如果對於拼裝SQL 完整內容有疑問，可以取消下面的註解執行匯出檢查，
            //其完整格式應與Resouse 資料夾內的 SQL.txt 雷同
            #region SQL文字檔            
            //string path = "SQL.txt";
            //FileStream file = new FileStream(path, FileMode.Create);
            //StreamWriter sw = new StreamWriter(file);
            //sw.Write(sql);

            //sw.Flush();
            //sw.Close();
            //file.Close();

            //QueryHelper qh = new QueryHelper();
            //qh.Select(sql);
            #endregion

            //2018/3/23 穎驊註記，因為ischool 主系統 在開啟時，會對於課程作快取，如果不重整，
            //選到本次被刪除的課程會造成系統當機
            // 加這主要是重新整理
            Course.Instance.SyncDataBackground(Course.Instance.SelectedKeys);
        }

    }
}
