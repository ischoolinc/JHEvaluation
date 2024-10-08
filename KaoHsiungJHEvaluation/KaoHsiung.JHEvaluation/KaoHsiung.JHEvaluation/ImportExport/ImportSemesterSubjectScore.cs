﻿using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool.API.PlugIn;
using Framework;
using System.Xml;
using System.Threading;
using JHSchool.Data;
using JHSchool;

namespace KaoHsiung.JHEvaluation.ImportExport
{
    class ImportSemesterSubjectScore : SmartSchool.API.PlugIn.Import.Importer
    {
        public ImportSemesterSubjectScore()
        {
            this.Image = null;
            this.Text = "匯入學期科目成績";
        }

        public override void InitializeImport(SmartSchool.API.PlugIn.Import.ImportWizard wizard)
        {
            Dictionary<string, int> _ID_SchoolYear_Semester_GradeYear = new Dictionary<string, int>();
            Dictionary<string, List<string>> _ID_SchoolYear_Semester_Subject = new Dictionary<string, List<string>>();
            Dictionary<string, JHStudentRecord> _StudentCollection = new Dictionary<string, JHStudentRecord>();
            Dictionary<JHStudentRecord, Dictionary<int, decimal>> _StudentPassScore = new Dictionary<JHStudentRecord, Dictionary<int, decimal>>();
            Dictionary<string, List<JHSemesterScoreRecord>> semsDict = new Dictionary<string, List<JHSemesterScoreRecord>>();

            wizard.PackageLimit = 3000;
            //wizard.ImportableFields.AddRange("領域", "科目", "學年度", "學期", "權數", "節數", "分數評量", "努力程度", "文字描述", "註記");
            //wizard.ImportableFields.AddRange("領域", "科目", "學年度", "學期", "權數", "節數", "分數評量", "文字描述", "註記");

            //2015.1.27 Cloud新增
            //2017/6/16 穎驊新增，因應[02-02][06] 計算學期科目成績新增清空原成績模式 項目， 新增 "刪除"欄位，使使用者能匯入 刪除成績資料
            wizard.ImportableFields.AddRange("領域", "科目", "學年度", "學期", "權數", "節數", "成績", "原始成績", "補考成績", "努力程度", "文字描述", "註記","刪除");
            
            wizard.RequiredFields.AddRange("領域", "科目", "學年度", "學期");

            wizard.ValidateStart += delegate(object sender, SmartSchool.API.PlugIn.Import.ValidateStartEventArgs e)
            {
                #region ValidateStart
                _ID_SchoolYear_Semester_GradeYear.Clear();
                _ID_SchoolYear_Semester_Subject.Clear();
                _StudentCollection.Clear();

                List<JHStudentRecord> list = JHStudent.SelectByIDs(e.List);

                MultiThreadWorker<JHStudentRecord> loader = new MultiThreadWorker<JHStudentRecord>();
                loader.MaxThreads = 3;
                loader.PackageSize = 250;
                loader.PackageWorker += delegate(object sender1, PackageWorkEventArgs<JHStudentRecord> e1)
                {
                    foreach (var item in JHSemesterScore.SelectByStudents(e1.List))
                    {
                        if (!semsDict.ContainsKey(item.RefStudentID))
                            semsDict.Add(item.RefStudentID, new List<JHSchool.Data.JHSemesterScoreRecord>());
                        semsDict[item.RefStudentID].Add(item);
                    }
                };
                loader.Run(list);

                foreach (JHStudentRecord stu in list)
                {
                    if (!_StudentCollection.ContainsKey(stu.ID))
                        _StudentCollection.Add(stu.ID, stu);
                }
                #endregion
            };
            wizard.ValidateRow += delegate(object sender, SmartSchool.API.PlugIn.Import.ValidateRowEventArgs e)
            {
                #region ValidateRow
                int t;
                decimal d;
                JHStudentRecord student;
                if (_StudentCollection.ContainsKey(e.Data.ID))
                {
                    student = _StudentCollection[e.Data.ID];
                }
                else
                {
                    e.ErrorMessage = "壓根就沒有這個學生" + e.Data.ID;
                    return;
                }
                bool inputFormatPass = true;
                #region 驗各欄位填寫格式
                foreach (string field in e.SelectFields)
                {
                    string value = e.Data[field];
                    switch (field)
                    {
                        default:
                            break;
                        case "領域":
                            //if (value == "")
                            //{
                            //    inputFormatPass &= false;
                            //    e.ErrorFields.Add(field, "必須填寫");
                            //}
                            //else if (!Domains.Contains(value))
                            //{
                            //    inputFormatPass &= false;
                            //    e.ErrorFields.Add(field, "必須為七大領域");
                            //}
                            break;
                        case "科目":
                            if (value == "")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填寫");
                            }
                            break;
                        case "學年度":
                            if (value == "" || !int.TryParse(value, out t))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入學年度");
                            }
                            break;
                        case "權數":
                        case "節數":
                            if (value == "" || !decimal.TryParse(value, out d))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入數值");
                            }
                            break;
                        //case "成績年級":
                        //    if (value == "" || !int.TryParse(value, out t))
                        //    {
                        //        inputFormatPass &= false;
                        //        e.ErrorFields.Add(field, "必須填入整數");
                        //    }
                        //    break;
                        case "學期":
                            if (value == "" || !int.TryParse(value, out t) || t > 2 || t < 1)
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入1或2");
                            }
                            break;

                        //case "分數評量":
                        //    if (value != "" && !decimal.TryParse(value, out d))
                        //    {
                        //        inputFormatPass &= false;
                        //        e.ErrorFields.Add(field, "必須填入空白或數值");
                        //    }
                        //    break;

                        //2015.1.27 Cloud新增
                        case "成績":
                            if (value != "" && !decimal.TryParse(value, out d))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或數值");
                            }
                            break;

                        //2015.1.27 Cloud新增
                        case "原始成績":
                            if (value != "" && !decimal.TryParse(value, out d))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或數值");
                            }
                            break;

                        //2015.1.27 Cloud新增
                        case "補考成績":
                            if (value != "" && !decimal.TryParse(value, out d))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或數值");
                            }
                            break;

                        case "努力程度":
                            if (value != "" && !int.TryParse(value, out t))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或數值");
                            }
                            break;

                        case "刪除":
                            if (value != "" && value != "是")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或 '是'");
                            }
                            break;
                    }
                }
                #endregion
                //輸入格式正確才會針對情節做檢驗
                if (inputFormatPass)
                {
                    string errorMessage = "";

                    string subject = e.Data["科目"];
                    string schoolYear = e.Data["學年度"];
                    string semester = e.Data["學期"];
                    int? sy = null;
                    int? se = null;
                    if (int.TryParse(schoolYear, out t))
                        sy = t;
                    if (int.TryParse(semester, out t))
                        se = t;
                    if (sy != null && se != null)
                    {
                        string key = e.Data.ID + "_" + sy + "_" + se;
                        #region 驗證新增科目成績
                        List<JHSemesterScoreRecord> semsList;
                        if (semsDict.ContainsKey(student.ID))
                            semsList = semsDict[student.ID];
                        else
                            semsList = new List<JHSemesterScoreRecord>();
                        foreach (JHSemesterScoreRecord record in semsList)
                        {
                            if (record.SchoolYear != sy) continue;
                            if (record.Semester != se) continue;

                            bool isNewSubjectInfo = true;
                            string message = "";
                            foreach (K12.Data.SubjectScore s in record.Subjects.Values)
                            {
                                if (s.Subject == subject)
                                    isNewSubjectInfo = false;
                            }
                            if (isNewSubjectInfo)
                            {
                                if (!e.WarningFields.ContainsKey("查無此科目"))
                                    e.WarningFields.Add("查無此科目", "學生在此學期並無此筆科目成績資訊，將會新增此科目成績");
                                foreach (string field in new string[] { "領域", "科目", "學年度", "學期", "權數", "節數" })
                                {
                                    if (!e.SelectFields.Contains(field))
                                        message += (message == "" ? "發現此學期無此科目，\n將會新增成績\n缺少成績必要欄位" : "、") + field;
                                }
                                if (message != "")
                                    errorMessage += (errorMessage == "" ? "" : "\n") + message;
                            }
                        }
                        #endregion
                        #region 驗證重複科目資料
                        //string skey = subject + "_" + le;
                        string skey = subject;
                        if (!_ID_SchoolYear_Semester_Subject.ContainsKey(key))
                            _ID_SchoolYear_Semester_Subject.Add(key, new List<string>());
                        if (_ID_SchoolYear_Semester_Subject[key].Contains(skey))
                        {
                            errorMessage += (errorMessage == "" ? "" : "\n") + "同一學期不允許多筆相同科目的資料";
                        }
                        else
                            _ID_SchoolYear_Semester_Subject[key].Add(skey);
                        #endregion
                    }
                    e.ErrorMessage = errorMessage;
                }
                #endregion
            };

            wizard.ImportPackage += delegate(object sender, SmartSchool.API.PlugIn.Import.ImportPackageEventArgs e)
            {
                #region ImportPackage
                Dictionary<string, List<RowData>> id_Rows = new Dictionary<string, List<RowData>>();
                #region 分包裝
                foreach (RowData data in e.Items)
                {
                    if (!id_Rows.ContainsKey(data.ID))
                        id_Rows.Add(data.ID, new List<RowData>());
                    id_Rows[data.ID].Add(data);
                }
                #endregion

                List<JHSemesterScoreRecord> insertList = new List<JHSemesterScoreRecord>();
                List<JHSemesterScoreRecord> updateList = new List<JHSemesterScoreRecord>();

                
                //交叉比對各學生資料
                #region 交叉比對各學生資料
                foreach (string id in id_Rows.Keys)
                {
                    XmlDocument doc = new XmlDocument();
                    JHStudentRecord studentRec = _StudentCollection[id];
                    //該學生的學期科目成績
                    Dictionary<SemesterInfo, JHSemesterScoreRecord> semesterScoreDictionary = new Dictionary<SemesterInfo, JHSemesterScoreRecord>();
                    #region 整理現有的成績資料
                    List<JHSchool.Data.JHSemesterScoreRecord> semsList;
                    if (semsDict.ContainsKey(studentRec.ID))
                        semsList = semsDict[studentRec.ID];
                    else
                        semsList = new List<JHSchool.Data.JHSemesterScoreRecord>();
                    foreach (JHSemesterScoreRecord var in semsList)
                    {
                        SemesterInfo info = new SemesterInfo();
                        info.SchoolYear = var.SchoolYear;
                        info.Semester = var.Semester;

                        if (!semesterScoreDictionary.ContainsKey(info))
                            semesterScoreDictionary.Add(info, var);

                        //string key = var.Subject + "_" + var.Level;
                        //if (!semesterScoreDictionary.ContainsKey(var.SchoolYear))
                        //    semesterScoreDictionary.Add(var.SchoolYear, new Dictionary<int, Dictionary<string, SemesterSubjectScoreInfo>>());
                        //if (!semesterScoreDictionary[var.SchoolYear].ContainsKey(var.Semester))
                        //    semesterScoreDictionary[var.SchoolYear].Add(var.Semester, new Dictionary<string, SemesterSubjectScoreInfo>());
                        //if (!semesterScoreDictionary[var.SchoolYear][var.Semester].ContainsKey(key))
                        //    semesterScoreDictionary[var.SchoolYear][var.Semester].Add(key, var);
                    }
                    #endregion

                    //要匯入的學期科目成績
                    Dictionary<SemesterInfo, Dictionary<string, RowData>> semesterImportScoreDictionary = new Dictionary<SemesterInfo, Dictionary<string, RowData>>();

                    #region 整理要匯入的資料
                    foreach (RowData row in id_Rows[id])
                    {
                        string subject = row["科目"];
                        string schoolYear = row["學年度"];
                        string semester = row["學期"];
                        int sy = int.Parse(schoolYear);
                        int se = int.Parse(semester);

                        SemesterInfo info = new SemesterInfo();
                        info.SchoolYear = sy;
                        info.Semester = se;

                        if (!semesterImportScoreDictionary.ContainsKey(info))
                            semesterImportScoreDictionary.Add(info, new Dictionary<string, RowData>());
                        if (!semesterImportScoreDictionary[info].ContainsKey(subject))
                            semesterImportScoreDictionary[info].Add(subject, row);
                    }
                    #endregion

                    //學期年級重整
                    //Dictionary<SemesterInfo, int> semesterGradeYear = new Dictionary<SemesterInfo, int>();
                    //要變更成績的學期
                    List<SemesterInfo> updatedSemester = new List<SemesterInfo>();
                    //在變更學期中新增加的成績資料
                    Dictionary<SemesterInfo, List<RowData>> updatedNewSemesterScore = new Dictionary<SemesterInfo, List<RowData>>();
                    //要增加成績的學期
                    Dictionary<SemesterInfo, List<RowData>> insertNewSemesterScore = new Dictionary<SemesterInfo, List<RowData>>();
                    //開始處理ImportScore
                    #region 開始處理ImportScore
                    foreach (SemesterInfo info in semesterImportScoreDictionary.Keys)
                    {                        
                        foreach (string subject in semesterImportScoreDictionary[info].Keys)
                        {
                            RowData data = semesterImportScoreDictionary[info][subject];
                            //如果是本來沒有這筆學期的成績就加到insertNewSemesterScore
                            if (!semesterScoreDictionary.ContainsKey(info))
                            {
                                if (!insertNewSemesterScore.ContainsKey(info))
                                    insertNewSemesterScore.Add(info, new List<RowData>());
                                insertNewSemesterScore[info].Add(data);
                            }
                            else
                            {
                                bool hasChanged = false;
                                //修改已存在的資料
                                if (semesterScoreDictionary[info].Subjects.ContainsKey(subject))
                                {
                                    JHSemesterScoreRecord record = semesterScoreDictionary[info];

                                    

                                    #region 直接修改已存在的成績資料的Detail
                                    foreach (string field in e.ImportFields)
                                    {
                                        K12.Data.SubjectScore score = record.Subjects[subject];
                                        string value = data[field];
                                        //"分數評量", "努力程度", "文字描述", "註記"
                                        switch (field)
                                        {
                                            default:
                                                break;
                                            case "領域":
                                                if (score.Domain != value)
                                                {
                                                    score.Domain = value;
                                                    hasChanged = true;
                                                }
                                                break;
                                            case "權數":
                                                if ("" + score.Credit != value)
                                                {
                                                    score.Credit = decimal.Parse(value);
                                                    hasChanged = true;
                                                }
                                                break;
                                            case "節數":
                                                if ("" + score.Period != value)
                                                {
                                                    score.Period = decimal.Parse(value);
                                                    hasChanged = true;
                                                }
                                                break;
                                            //case "成績年級":
                                            //    int gy = int.Parse(data["成績年級"]);
                                            //    if (record.GradeYear != gy)
                                            //    {
                                            //        semesterGradeYear[info] = gy;
                                            //        hasChanged = true;
                                            //    }
                                            //    break;
                                            //case "分數評量":
                                            //    if ("" + score.Score != value)
                                            //    {
                                            //        decimal d;
                                            //        if (decimal.TryParse(value, out d))
                                            //            score.Score = d;
                                            //        else
                                            //            score.Score = null;
                                            //        hasChanged = true;
                                            //    }
                                            //    break;

                                            //2015.1.27 Cloud新增
                                            case "成績":
                                                if ("" + score.Score != value)
                                                {
                                                    decimal d;
                                                    if (decimal.TryParse(value, out d))
                                                        score.Score = d;
                                                    else
                                                        score.Score = null;
                                                    hasChanged = true;
                                                }
                                                break;

                                            //2015.1.27 Cloud新增
                                            case "原始成績":
                                                if ("" + score.ScoreOrigin != value)
                                                {
                                                    decimal d;
                                                    if (decimal.TryParse(value, out d))
                                                        score.ScoreOrigin = d;
                                                    else
                                                        score.ScoreOrigin = null;
                                                    hasChanged = true;
                                                }
                                                break;

                                            //2015.1.27 Cloud新增
                                            case "補考成績":
                                                if ("" + score.ScoreMakeup != value)
                                                {
                                                    decimal d;
                                                    if (decimal.TryParse(value, out d))
                                                        score.ScoreMakeup = d;
                                                    else
                                                        score.ScoreMakeup = null;
                                                    hasChanged = true;
                                                }
                                                break;

                                            case "努力程度":
                                                if ("" + score.Effort != value)
                                                {
                                                    int i;
                                                    if (int.TryParse(value, out i))
                                                        score.Effort = i;
                                                    else
                                                        score.Effort = null;
                                                    hasChanged = true;
                                                }
                                                break;
                                            case "文字描述":
                                                if ("" + score.Text != value)
                                                {
                                                    score.Text = value;
                                                    hasChanged = true;
                                                }
                                                break;
                                            case "註記":
                                                if (score.Comment != value)
                                                {
                                                    score.Comment = value;
                                                    hasChanged = true;
                                                }
                                                break;
                                            case "刪除":
                                                if (value =="是")
                                                {
                                                    record.Subjects.Remove(subject);
                                                    hasChanged = true;                                                    
                                                }
                                                break;
                                        }
                                    }
                                    #endregion
                                }
                                else//加入新成績至已存在的學期
                                {
                                    if (!updatedNewSemesterScore.ContainsKey(info))
                                        updatedNewSemesterScore.Add(info, new List<RowData>());
                                    updatedNewSemesterScore[info].Add(data);
                                    hasChanged = true;
                                }
                                //真的有變更
                                if (hasChanged)
                                {
                                    #region 登錄有變更的學期
                                    if (!updatedSemester.Contains(info))
                                        updatedSemester.Add(info);
                                    #endregion
                                }
                            }
                        }
                    }
                    #endregion
                    //處理已登錄要更新的學期成績
                    #region 處理已登錄要更新的學期成績
                    foreach (SemesterInfo info in updatedSemester)
                    {
                        //Dictionary<int, Dictionary<int, string>> semeScoreID = (Dictionary<int, Dictionary<int, string>>)studentRec.Fields["SemesterSubjectScoreID"];
                        //string semesterScoreID = semeScoreID[sy][se];//從學期抓ID
                        //int gradeyear = semesterGradeYear[info];//抓年級
                        //XmlElement subjectScoreInfo = doc.CreateElement("SemesterSubjectScoreInfo");
                        #region 產生該學期科目成績的XML
                        //foreach (SemesterSubjectScoreInfo scoreInfo in semesterScoreDictionary[sy][se].Values)
                        //{
                        //    subjectScoreInfo.AppendChild(doc.ImportNode(scoreInfo.Detail, true));
                        //}

                        updateList.Add(semesterScoreDictionary[info]);

                        //if (updatedNewSemesterScore.ContainsKey(sy) && updatedNewSemesterScore[sy].ContainsKey(se))
                        if (updatedNewSemesterScore.ContainsKey(info))
                        {
                            foreach (RowData row in updatedNewSemesterScore[info])
                            {
                                //XmlElement newScore = doc.CreateElement("Subject");
                                K12.Data.SubjectScore subjectScore = new K12.Data.SubjectScore();

                                bool contain_deleted_words = false;

                                #region 建立newScore
                                //foreach (string field in new string[] { "領域", "科目", "權數", "節數", "分數評量", "努力程度", "文字描述", "註記" })
                                foreach (string field in new string[] { "領域", "科目", "權數", "節數", "成績", "原始成績", "補考成績", "努力程度", "文字描述", "註記","刪除" })
                                {
                                    if (e.ImportFields.Contains(field))
                                    {
                                        decimal d;

                                        #region 填入科目資訊
                                        string value = row[field];
                                        switch (field)
                                        {
                                            default:
                                                break;
                                            case "領域":
                                                subjectScore.Domain = value;
                                                break;
                                            case "科目":
                                                subjectScore.Subject = value;
                                                break;
                                            case "權數":
                                                subjectScore.Credit = decimal.Parse(value);
                                                break;
                                            case "節數":
                                                subjectScore.Period = decimal.Parse(value);
                                                break;
                                            //case "分數評量":
                                            //    decimal d;
                                            //    if (decimal.TryParse(value, out d))
                                            //        subjectScore.Score = d;
                                            //    else
                                            //        subjectScore.Score = null;
                                            //    break;

                                            //2015.1.27 Cloud新增
                                            case "成績":
                                                if (decimal.TryParse(value, out d))
                                                    subjectScore.Score = d;
                                                else
                                                    subjectScore.Score = null;
                                                break;
                                            //2015.1.27 Cloud新增
                                            case "原始成績":
                                                if (decimal.TryParse(value, out d))
                                                    subjectScore.ScoreOrigin = d;
                                                else
                                                    subjectScore.ScoreOrigin = null;
                                                break;
                                            //2015.1.27 Cloud新增
                                            case "補考成績":
                                                if (decimal.TryParse(value, out d))
                                                    subjectScore.ScoreMakeup = d;
                                                else
                                                    subjectScore.ScoreMakeup = null;
                                                break;

                                            case "努力程度":
                                                int i;
                                                if (int.TryParse(value, out i))
                                                    subjectScore.Effort = i;
                                                else
                                                    subjectScore.Effort = null;
                                                break;
                                            case "文字描述":
                                                subjectScore.Text = value;
                                                break;
                                            case "註記":
                                                subjectScore.Comment = value;
                                                break;
                                            case "刪除":
                                                if (value == "是")
                                                {
                                                    contain_deleted_words = true;
                                                }
                                                break;
                                        }
                                        #endregion
                                    }
                                }
                                #endregion
                                //subjectScoreInfo.AppendChild(newScore);

                                if (!contain_deleted_words) 
                                {
                                    JHSemesterScoreRecord record = semesterScoreDictionary[info];

                                    if (!record.Subjects.ContainsKey(subjectScore.Subject))
                                        record.Subjects.Add(subjectScore.Subject, subjectScore);
                                    else
                                        record.Subjects[subjectScore.Subject] = subjectScore;

                                    updateList.Add(record);
                                
                                }                                
                            }
                        }
                        #endregion
                        //updateList.Add(new SmartSchool.Feature.Score.EditScore.UpdateInfo(semesterScoreID, gradeyear, subjectScoreInfo));
                    }
                    #endregion
                    //處理新增成績學期
                    #region 處理新增成績學期
                    foreach (SemesterInfo info in insertNewSemesterScore.Keys)
                    {
                        //int gradeyear = semesterGradeYear[info];//抓年級
                        foreach (RowData row in insertNewSemesterScore[info])
                        {
                            K12.Data.SubjectScore subjectScore = new K12.Data.SubjectScore();


                            bool contain_deleted_words = false;

                            //foreach (string field in new string[] { "領域", "科目", "權數", "節數", "分數評量", "努力程度", "文字描述", "註記" })
                            foreach (string field in new string[] { "領域", "科目", "權數", "節數", "成績", "原始成績", "補考成績", "努力程度", "文字描述", "註記" ,"刪除"})
                            {
                                if (e.ImportFields.Contains(field))
                                {
                                    decimal d;

                                    string value = row[field];
                                    switch (field)
                                    {
                                        default: break;
                                        case "領域":
                                            subjectScore.Domain = value;
                                            break;
                                        case "科目":
                                            subjectScore.Subject = value;
                                            break;
                                        case "權數":
                                            subjectScore.Credit = decimal.Parse(value);
                                            break;
                                        case "節數":
                                            subjectScore.Period = decimal.Parse(value);
                                            break;
                                        //case "分數評量":
                                        //    decimal d;
                                        //    if (decimal.TryParse(value, out d))
                                        //        subjectScore.Score = d;
                                        //    else
                                        //        subjectScore.Score = null;
                                        //    break;

                                        //2015.1.27 Cloud新增
                                        case "成績":
                                            if (decimal.TryParse(value, out d))
                                                subjectScore.Score = d;
                                            else
                                                subjectScore.Score = null;
                                            break;
                                        //2015.1.27 Cloud新增
                                        case "原始成績":
                                            if (decimal.TryParse(value, out d))
                                                subjectScore.ScoreOrigin = d;
                                            else
                                                subjectScore.ScoreOrigin = null;
                                            break;
                                        //2015.1.27 Cloud新增
                                        case "補考成績":
                                            if (decimal.TryParse(value, out d))
                                                subjectScore.ScoreMakeup = d;
                                            else
                                                subjectScore.ScoreMakeup = null;
                                            break;

                                        case "努力程度":
                                            int i;
                                            if (int.TryParse(value, out i))
                                                subjectScore.Effort = i;
                                            else
                                                subjectScore.Effort = null;
                                            break;
                                        case "文字描述":
                                            subjectScore.Text = value;
                                            break;
                                        case "註記":
                                            subjectScore.Comment = value;
                                            break;
                                        case "刪除":
                                            if (value == "是")
                                            {
                                                contain_deleted_words = true;
                                            }                                            
                                            break;
                                    }
                                }
                            }
                            //subjectScoreInfo.AppendChild(newScore);
                            JHSemesterScoreRecord record = new JHSemesterScoreRecord();
                            record.SchoolYear = info.SchoolYear;
                            record.Semester = info.Semester;
                            record.RefStudentID = studentRec.ID;
                            //record.GradeYear = gradeyear;

                            if (!record.Subjects.ContainsKey(subjectScore.Subject))
                                record.Subjects.Add(subjectScore.Subject, subjectScore);
                            else
                                record.Subjects[subjectScore.Subject] = subjectScore;

                            if (!contain_deleted_words) 
                            {
                                insertList.Add(record);
                            }
                            
                        }
                        //insertList.Add(new SmartSchool.Feature.Score.AddScore.InsertInfo(studentRec.StudentID, "" + sy, "" + se, gradeyear, "", subjectScoreInfo));
                    }
                    #endregion
                }

                #endregion

                Dictionary<string, JHSemesterScoreRecord> iList = new Dictionary<string, JHSemesterScoreRecord>();
                Dictionary<string, JHSemesterScoreRecord> uList = new Dictionary<string, JHSemesterScoreRecord>();

                foreach (var record in insertList)
                {
                    string key = record.RefStudentID + "_" + record.SchoolYear + "_" + record.Semester;
                    if (!iList.ContainsKey(key))
                        iList.Add(key, new JHSemesterScoreRecord());
                    JHSemesterScoreRecord newRecord = iList[key];
                    newRecord.RefStudentID = record.RefStudentID;
                    newRecord.SchoolYear = record.SchoolYear;
                    newRecord.Semester = record.Semester;

                    foreach (var subject in record.Subjects.Keys)
                    {
                        if (!newRecord.Subjects.ContainsKey(subject))
                            newRecord.Subjects.Add(subject, record.Subjects[subject]);
                        if (newRecord.Subjects[subject].Text != null)
                        {
                            if (newRecord.Subjects[subject].Text.Contains("\b"))
                                newRecord.Subjects[subject].Text = newRecord.Subjects[subject].Text.Replace("\b", "");
                        }
                    }
                }

                foreach (var record in updateList)
                {
                    string key = record.RefStudentID + "_" + record.SchoolYear + "_" + record.Semester;
                    if (!uList.ContainsKey(key))
                        uList.Add(key, record);
                    JHSemesterScoreRecord newRecord = uList[key];
                    newRecord.RefStudentID = record.RefStudentID;
                    newRecord.SchoolYear = record.SchoolYear;
                    newRecord.Semester = record.Semester;
                    newRecord.ID = record.ID;

                    foreach (var subject in record.Subjects.Keys)
                    {
                        if (!newRecord.Subjects.ContainsKey(subject))
                            newRecord.Subjects.Add(subject, record.Subjects[subject]);
                        if (newRecord.Subjects[subject].Text != null)
                        {
                            if (newRecord.Subjects[subject].Text.Contains("\b"))
                                newRecord.Subjects[subject].Text = newRecord.Subjects[subject].Text.Replace("\b", "");
                        }
                    }
                }

                List<string> ids = new List<string>(id_Rows.Keys);
                Dictionary<string, JHSemesterScoreRecord> origs = new Dictionary<string, JHSemesterScoreRecord>();
                foreach (var record in JHSemesterScore.SelectByStudentIDs(ids))
                {
                    if (!origs.ContainsKey(record.ID))
                        origs.Add(record.ID, record);
                }
                foreach (var record in uList.Values)
                {
                    if (origs.ContainsKey(record.ID))
                    {
                        foreach (var domain in origs[record.ID].Domains.Keys)
                        {
                            if (!record.Domains.ContainsKey(domain))
                                record.Domains.Add(domain, origs[record.ID].Domains[domain]);
                        }
                    }
                }

                //FunctionSpliter<JHSemesterScoreRecord, int> splitInsert = new FunctionSpliter<JHSemesterScoreRecord, int>(200, 3);
                //splitInsert.Function = delegate(List<JHSemesterScoreRecord> part)
                //{
                //    JHSemesterScore.Insert(part);
                //    return null;
                //};
                //splitInsert.Execute(new List<JHSemesterScoreRecord>(iList.Values));

                //FunctionSpliter<JHSemesterScoreRecord, int> splitUpdate= new FunctionSpliter<JHSemesterScoreRecord, int>(200, 3);
                //splitUpdate.Function = delegate(List<JHSemesterScoreRecord> part)
                //{
                //    JHSemesterScore.Update(part);
                //    return null;
                //};
                //splitUpdate.Execute(new List<JHSemesterScoreRecord>(uList.Values));

                JHSemesterScore.Insert(iList.Values);
                JHSemesterScore.Update(uList.Values);

                FISCA.LogAgent.ApplicationLog.Log("成績系統.匯入匯出", "匯入學期科目成績", "總共匯入" + (insertList.Count + updateList.Count) + "筆學期科目成績。");
                #endregion
            };
            wizard.ImportComplete += delegate
            {
                MsgBox.Show("匯入完成");
            };
        }

        private void updateSemesterSubjectScore(object item)
        {
            List<List<JHSemesterScoreRecord>> updatePackages = (List<List<JHSemesterScoreRecord>>)item;
            foreach (List<JHSemesterScoreRecord> package in updatePackages)
            {
                JHSemesterScore.Update(package);
            }
        }

        private void insertSemesterSubjectScore(object item)
        {
            List<List<JHSemesterScoreRecord>> insertPackages = (List<List<JHSemesterScoreRecord>>)item;
            foreach (List<JHSemesterScoreRecord> package in insertPackages)
            {
                JHSemesterScore.Insert(package);
            }
        }
    }
}
