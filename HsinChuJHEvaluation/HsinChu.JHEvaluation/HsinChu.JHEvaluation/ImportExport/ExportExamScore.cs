using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.API.PlugIn;
using JHSchool.Data;
using HsinChu.JHEvaluation.Data;
using System.Xml;
using System.Windows.Forms;

namespace HsinChu.JHEvaluation.ImportExport
{

    class ExportExamScore : SmartSchool.API.PlugIn.Export.Exporter
    {
        public ExportExamScore()
        {
            this.Image = null;
            this.Text = "匯出評量成績";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            // 2018.09.06 [ischoolKingdom] Vicky依據 [12-01][01] 多學期成績排名 項目，增加 "領域", "科目", "節數(權重)", "學分" 項目資料。
            // 2018.09.07 [ischoolKingdom] Vicky依據 [12-01][01] 多學期成績排名 新需求更動項目排序。
            wizard.ExportableFields.AddRange("學年度", "學期", "課程名稱", "領域", "科目", "節數(權重)", "學分", "評量名稱", "定期分數", "平時分數", "文字描述");
            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                //學生資訊
                List<JHStudentRecord> students = JHStudent.SelectByIDs(e.List);
                //學生修課資訊
                Dictionary<string, List<JHSCAttendRecord>> scattends = new Dictionary<string, List<JHSchool.Data.JHSCAttendRecord>>();
                //學生修習的課程
                Dictionary<string, JHCourseRecord> courses = new Dictionary<string, JHSchool.Data.JHCourseRecord>();
                //評量成績 key: SCAttendID
                Dictionary<string, List<HC.JHSCETakeRecord>> sces = new Dictionary<string, List<HC.JHSCETakeRecord>>();
                //試別資訊
                Dictionary<string, JHExamRecord> exams = new Dictionary<string, JHSchool.Data.JHExamRecord>();

                #region 取得修課記錄
                foreach (JHSCAttendRecord record in JHSCAttend.SelectByStudentIDAndCourseID(e.List, new string[] { }))
                {
                    if (!scattends.ContainsKey(record.RefStudentID))
                        scattends.Add(record.RefStudentID, new List<JHSchool.Data.JHSCAttendRecord>());
                    scattends[record.RefStudentID].Add(record);

                    if (!courses.ContainsKey(record.RefCourseID))
                        courses.Add(record.RefCourseID, null);
                }
                #endregion

                #region 取得課程資訊
                foreach (JHCourseRecord record in JHCourse.SelectByIDs(new List<string>(courses.Keys)))
                {
                    if (courses.ContainsKey(record.ID))
                        courses[record.ID] = record;
                }
                #endregion

                #region 取得試別資訊
                foreach (JHExamRecord exam in JHExam.SelectAll())
                {
                    if (!exams.ContainsKey(exam.ID))
                        exams.Add(exam.ID, exam);
                }
                #endregion

                #region 取得評量成績
                foreach (HC.JHSCETakeRecord record in JHSCETake.SelectByStudentAndCourse(new List<string>(scattends.Keys), new List<string>(courses.Keys)).AsHCJHSCETakeRecords())
                {
                    if (!sces.ContainsKey(record.RefSCAttendID))
                        sces.Add(record.RefSCAttendID, new List<HC.JHSCETakeRecord>());
                    sces[record.RefSCAttendID].Add(record);
                }
                #endregion


                #region 取得評量成績缺考暨免試設定

                PluginMain.ScoreTextMap.Clear();
                PluginMain.ScoreValueMap.Clear();
                Framework.ConfigData cd = JHSchool.School.Configuration["評量成績缺考暨免試設定"];
                if (!string.IsNullOrEmpty(cd["評量成績缺考暨免試設定"]))
                {
                    XmlElement element = Framework.XmlHelper.LoadXml(cd["評量成績缺考暨免試設定"]);

                    foreach (XmlElement each in element.SelectNodes("Setting"))
                    {
                        var UseText = each.SelectSingleNode("UseText").InnerText;
                        var AllowCalculation = bool.Parse(each.SelectSingleNode("AllowCalculation").InnerText);
                        decimal Score;
                        decimal.TryParse(each.SelectSingleNode("Score").InnerText, out Score);
                        var Active = bool.Parse(each.SelectSingleNode("Active").InnerText);
                        var UseValue = decimal.Parse(each.SelectSingleNode("UseValue").InnerText);

                        if (Active)
                        {
                            if (!PluginMain.ScoreTextMap.ContainsKey(UseText))
                            {
                                PluginMain.ScoreTextMap.Add(UseText, new ScoreMap
                                {
                                    UseText = UseText,
                                    AllowCalculation = AllowCalculation,
                                    Score = Score,
                                    Active = Active,
                                    UseValue = UseValue,
                                });
                            }
                            if (!PluginMain.ScoreValueMap.ContainsKey(UseValue))
                            {
                                PluginMain.ScoreValueMap.Add(UseValue, new ScoreMap
                                {
                                    UseText = UseText,
                                    AllowCalculation = AllowCalculation,
                                    Score = Score,
                                    Active = Active,
                                    UseValue = UseValue,
                                });
                            }
                        }
                    }
                }

                #endregion

                #region 產生 Row Data
                try
                {
                    foreach (JHStudentRecord stu in students)
                    {
                        if (!scattends.ContainsKey(stu.ID))
                        {
                            continue;
                        }

                        foreach (JHSCAttendRecord record in scattends[stu.ID])
                        {
                            if (!sces.ContainsKey(record.ID)) continue;

                            sces[record.ID].Sort(delegate (HC.JHSCETakeRecord x, HC.JHSCETakeRecord y)
                            {
                                return x.RefExamID.CompareTo(y.RefExamID);
                            });

                            foreach (HC.JHSCETakeRecord sce in sces[record.ID])
                            {
                                string examName = sce.RefExamID;
                                if (exams.ContainsKey(sce.RefExamID))
                                    examName = exams[sce.RefExamID].Name;

                                RowData row = new RowData();
                                row.ID = stu.ID;
                                foreach (string field in e.ExportFields)
                                {
                                    if (wizard.ExportableFields.Contains(field))
                                    {
                                        switch (field)
                                        {
                                            case "學年度": row.Add(field, "" + courses[record.RefCourseID].SchoolYear); break;
                                            case "學期": row.Add(field, "" + courses[record.RefCourseID].Semester); break;
                                            case "課程名稱": row.Add(field, courses[record.RefCourseID].Name); break;
                                            // 2018.09.06 [ischoolKingdom] Vicky依據 [12-01][01] 多學期成績排名 項目，增加 "領域", "科目", "節數(權重)", "學分" 項目資料。
                                            case "領域": row.Add(field, courses[record.RefCourseID].Domain); break;
                                            case "科目": row.Add(field, courses[record.RefCourseID].Subject); break;
                                            case "節數(權重)": row.Add(field, "" + courses[record.RefCourseID].Period); break;
                                            case "學分": row.Add(field, "" + courses[record.RefCourseID].Credit); break;

                                            case "評量名稱": row.Add(field, examName); break;
                                            case "定期分數":
                                                if (sce.Score.HasValue)
                                                {
                                                    if (PluginMain.ScoreValueMap.ContainsKey(sce.Score.Value))
                                                    {
                                                        row.Add(field, PluginMain.ScoreValueMap[sce.Score.Value].UseText);
                                                    }
                                                    else
                                                    {
                                                        row.Add(field, sce.Score.Value + "");
                                                    }
                                                }
                                                else
                                                {
                                                    row.Add(field, "");
                                                }
                                                break;
                                            case "平時分數":
                                                if (sce.AssignmentScore.HasValue)
                                                {
                                                    if (PluginMain.ScoreValueMap.ContainsKey(sce.AssignmentScore.Value))
                                                    {
                                                        row.Add(field, PluginMain.ScoreValueMap[sce.AssignmentScore.Value].UseText);
                                                    }
                                                    else
                                                    {
                                                        row.Add(field, sce.AssignmentScore.Value + "");
                                                    }
                                                }
                                                else
                                                {
                                                    row.Add(field, "");
                                                }
                                                break;
                                            case "文字描述": row.Add(field, sce.Text); break;

                                        }
                                    }
                                }
                                e.Items.Add(row);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    ;       // MessageBox.Show(ex.Message);
                }
                #endregion

                FISCA.LogAgent.ApplicationLog.Log("成績系統.匯入匯出", "匯出評量成績", "總共匯出" + e.Items.Count + "筆評量成績。");
            };
        }
    }
}
