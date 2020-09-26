﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using JHSchool.Data;
using System.IO;
using HsinChu.JHEvaluation.Data;
using Aspose.Words;
using JHSchool.Evaluation.Calculation;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using FISCA.Data;
using Campus.ePaperCloud;
using HsinChuExamScore_JH.DAO;

namespace HsinChuExamScore_JH
{
    public partial class PrintForm : BaseForm, IFieldMergingCallback
    {

        private Boolean HasReferenceExam = false;
        private Boolean btnPrintClick = false;
        private Type _ScoreCompositions = typeof(EnumScoreComposition); // 取得 EnumScoreComposition  (總成績,定期評量,平時評量)
        string[] itemNames = new string[] { "平均", "總分" };
        Dictionary<string, string> itemTypesMapping = new Dictionary<string, string>() {
                        {"定期評量_定期/總計成績","定期成績"},
                        {"定期評量/總計成績","成績"}}; //因資料庫內 itemTypes用詞過長 轉換為較短並符合其他合併欄位命名方式的欄位
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        private Dictionary<string, List<string>> _ExamSubjects = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> _ExamSubjectFull = new Dictionary<string, List<string>>();
        List<string> _StudentIDList;
        List<string> UserDefineField; // 裝自訂欄位
        // 缺曠區間統計
        Dictionary<string, Dictionary<string, int>> _AttendanceDict = new Dictionary<string, Dictionary<string, int>>();

        ////領域加權排名資料 字典，存放、最後合併列印使用 (studentID,<MergeFieldkey,RankValue>)
        //Dictionary<string, Dictionary<string, string>> _DomainWeightRankDataDict = new Dictionary<string, Dictionary<string, string>>();

        ////科目排名資料 字典，存放、最後合併列印使用 (studentID,<subjectName +Key,RankValue>)
        //Dictionary<string, Dictionary<string, string>> _subjectRankDataDict = new Dictionary<string, Dictionary<string, string>>();

        ////科目排名資料 字典，存放、最後合併列印使用 (studentID,<domainName +Key,RankValue>)
        //Dictionary<string, Dictionary<string, string>> _domainRankDataDict = new Dictionary<string, Dictionary<string, string>>();

        //成績計算規則
        //key: ScoreCalcRuleID
        Dictionary<string, ScoreCalculator> calcCache = new Dictionary<string, ScoreCalculator>();
        //key: StudentID, val: ScoreCalcRuleID
        Dictionary<string, string> calcIDCache = new Dictionary<string, string>();

        //成績顯示位數
        int parseNumber = 0;

        private List<string> typeList = new List<string>();
        private List<string> absenceList = new List<string>();
        /// <summary>
        /// 勾選的科目
        /// </summary>
        private List<string> _SelSubjNameList = new List<string>();
        private List<string> _SelAttendanceList = new List<string>();

        private BackgroundWorker _bgWorkReport;
        private DocumentBuilder _builder;
        BackgroundWorker bkw;

        /// 錯誤訊息
        List<string> _ErrorList = new List<string>();

        // 領域錯誤訊息
        List<string> _ErrorDomainNameList = new List<string>();

        // 樣板內有科目名稱
        List<string> _TemplateSubjectNameList = new List<string>();

        // 存檔路徑
        string pathW = "";

        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();

        // 開始日期
        private DateTime _BeginDate;
        // 結束日期
        private DateTime _EndDate;

        // 成績校正日期字串
        private string _ScoreEditDate = "";

        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        private int _SelSchoolYear;
        private int _SelSemester;
        private int _SelRefSchoolYear;
        private int _SelRefSemester;
        private string _SelExamName = "";
        private string _SelExamID = "";
        private string _SelNotRankedFilter = "";


        private string _SelRefExamName = "";
        private string _SelRefExamID = "";

        ScoreMappingConfig _ScoreMappingConfig = new ScoreMappingConfig();

        private List<ExamRecord> _exams = new List<ExamRecord>();

        /// <summary>
        /// 裝有那些自訂欄位
        /// </summary>
        private List<string> _UserDefineFields; 
        private Dictionary<string, List<string>> _StudTagDict = new Dictionary<string, List<string>>();

        // 紀錄樣板設定
        List<DAO.UDT_ScoreConfig> _UDTConfigList;

        public PrintForm(List<string> StudIDList)
        {
            InitializeComponent();
            _StudentIDList = StudIDList;
            bkw = new BackgroundWorker();
            bkw.DoWork += new DoWorkEventHandler(bkw_DoWork);
            bkw.ProgressChanged += new ProgressChangedEventHandler(bkw_ProgressChanged);
            bkw.ProgressChanged += new ProgressChangedEventHandler(bkw_ProgressChanged2);
            bkw.WorkerReportsProgress = true;
            bkw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bkw_RunWorkerCompleted);

            _bgWorkReport = new BackgroundWorker();
            _bgWorkReport.DoWork += new DoWorkEventHandler(_bgWorkReport_DoWork);
            _bgWorkReport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorkReport_RunWorkerCompleted);
            _bgWorkReport.WorkerReportsProgress = true;
            _bgWorkReport.ProgressChanged += new ProgressChangedEventHandler(_bgWorkReport_ProgressChanged);
        }

        void bkw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            circularProgress1.Value = e.ProgressPercentage;
        }

        void bkw_ProgressChanged2(object sender, ProgressChangedEventArgs e)
        {
            circularProgress2.Value = e.ProgressPercentage;
        }

        void bkw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            EnbSelect();
            
            _DefalutSchoolYear = K12.Data.School.DefaultSchoolYear;
            _DefaultSemester = K12.Data.School.DefaultSemester;

            if (_Configure == null)
                _Configure = new Configure();

            cboConfigure.Items.Clear();
            foreach (var item in _ConfigureList)
            {
                cboConfigure.Items.Add(item);
            }
            cboConfigure.Items.Add(new Configure() { Name = "新增" });
            int i;

            if (int.TryParse(_DefalutSchoolYear, out i))
            {
                for (int j = 5; j > 0; j--)
                {
                    cboSchoolYear.Items.Add("" + (i - j));
                    cboRefSchoolYear.Items.Add("" + (i - j));
                }

                for (int j = 0; j < 3; j++)
                {
                    cboSchoolYear.Items.Add("" + (i + j));
                    cboRefSchoolYear.Items.Add("" + (i + j));
                }

            }

            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");
            cboRefSemester.Items.Add("1");
            cboRefSemester.Items.Add("2");
            cboExam.Items.Clear();
            cboRefExam.Items.Clear();

            foreach (ExamRecord exName in _exams)
            {
                cboExam.Items.Add(exName.Name);
                cboRefExam.Items.Add(exName.Name);
            }

            circularProgress1.Hide();
            if (_ConfigureList.Count > 0)
            {
                cboConfigure.SelectedIndex = 0;
            }
            else
            {
                cboConfigure.SelectedIndex = -1;
            }

            if (_Configure.PrintAttendanceList == null)
                _Configure.PrintAttendanceList = new List<string>();

            DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn();
            colName.HeaderText = "節次分類";
            colName.MinimumWidth = 70;
            colName.Name = "colName";
            colName.ReadOnly = true;
            colName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            colName.Width = 70;
            this.dgAttendanceData.Columns.Add(colName);

            List<string> colNameList = new List<string>();
            foreach (string absence in absenceList)
            {
                System.Windows.Forms.DataGridViewCheckBoxColumn newCol = new DataGridViewCheckBoxColumn();
                newCol.HeaderText = absence;
                newCol.Width = 55;
                newCol.ReadOnly = false;
                newCol.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
                newCol.Tag = absence;
                colNameList.Add(absence);
                newCol.ValueType = typeof(bool);
                this.dgAttendanceData.Columns.Add(newCol);
            }

            foreach (string str in typeList)
            {
                int rowIdx = dgAttendanceData.Rows.Add();
                dgAttendanceData.Rows[rowIdx].Tag = str;
                dgAttendanceData.Rows[rowIdx].Cells[0].Value = str;
                int colIdx = 1;

                foreach (string str1 in colNameList)
                {
                    string key = str + "_" + str1;
                    DataGridViewCheckBoxCell cell = new DataGridViewCheckBoxCell();
                    cell.Tag = key;
                    cell.Value = false;
                    if (_Configure.PrintAttendanceList.Contains(key))
                        cell.Value = true;

                    dgAttendanceData.Rows[rowIdx].Cells[colIdx] = cell;
                    colIdx++;
                }
            }

            string userSelectConfigName = "";
            // 檢查畫面上是否有使用者選的
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    userSelectConfigName = conf.Name;
                    break;
                }

            if (!string.IsNullOrEmpty(_Configure.SelSetConfigName))
                cboConfigure.Text = userSelectConfigName;


            btnSaveConfig.Enabled = btnPrint.Enabled = true;
        }

        void bkw_DoWork(object sender, DoWorkEventArgs e)
        {
            bkw.ReportProgress(1);

            //試別清單
            _exams.Clear();
            _exams = K12.Data.Exam.SelectAll();


            // 檢查預設樣板是否存在
            _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);

            // 沒有設定檔，建立預設設定檔
            if (_UDTConfigList.Count < 2)
            {
                bkw.ReportProgress(10);
                foreach (string name in Global.DefaultConfigNameList())
                {
                    Configure cn = new Configure();
                    cn.Name = name;
                    cn.SchoolYear = K12.Data.School.DefaultSchoolYear;
                    cn.Semester = K12.Data.School.DefaultSemester;
                    cn.RefSchoolYear = K12.Data.School.DefaultSchoolYear;
                    cn.RefSemester = K12.Data.School.DefaultSemester;
                    DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                    conf.Name = name;
                    conf.UDTTableName = Global._UDTTableName;
                    conf.ProjectName = Global._ProjectName;
                    conf.Type = Global._DefaultConfTypeName;
                    _UDTConfigList.Add(conf);

                    // 設預設樣板
                    switch (name)
                    {
                        case "領域成績單":
                            cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_領域成績單));
                            break;

                        case "科目成績單":
                            cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目成績單));
                            break;

                        case "科目及領域成績單_領域組距":
                            cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_領域組距));
                            break;
                        case "科目及領域成績單_科目組距":
                            cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_科目組距));
                            break;
                    }

                    if (cn.Template == null)
                        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹評量成績單樣板_固定排名_科目_領域__doc1));
                    cn.Encode();
                    cn.Save();
                }
                if (_UDTConfigList.Count > 0)
                    DAO.UDTTransfer.InsertConfigData(_UDTConfigList);
            }
            bkw.ReportProgress(20);
            // 取的設定資料
            _ConfigureList = _AccessHelper.Select<Configure>();

            bkw.ReportProgress(40);
            // 缺曠資料
            foreach (JHPeriodMappingInfo info in JHPeriodMapping.SelectAll())
            {
                if (!typeList.Contains(info.Type))
                    typeList.Add(info.Type);
            }

            bkw.ReportProgress(70);
            foreach (JHAbsenceMappingInfo info in JHAbsenceMapping.SelectAll())
            {
                if (!absenceList.Contains(info.Name))
                    absenceList.Add(info.Name);
            }
            bkw.ReportProgress(80);
            // 所有有學生類別
            _StudTagDict = Utility.GetStudentTagRefDict();


            bkw.ReportProgress(100);
        }

        void _bgWorkReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("評量成績報表產生中...", e.ProgressPercentage);
        }

        void _bgWorkReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                object[] objArray = (object[])e.Result;
                MemoryStream memoryStream = (MemoryStream)objArray[0];
                string reportNameW = "" + _SelSchoolYear + "學年度第" + _SelSemester + "學期" + _SelExamName + "" + objArray[1];
                ePaperCloud ePaperCloud = new ePaperCloud();
                ePaperCloud.upload_ePaper(_SelSchoolYear, _SelSemester, reportNameW, "", memoryStream, ePaperCloud.ViewerType.Student, ePaperCloud.FormatType.Docx);

                btnSaveConfig.Enabled = true;
                btnPrint.Enabled = true;

                if (_ErrorList.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    //sb.AppendLine("樣板內科目合併欄位不足，請新增：");
                    //sb.AppendLine(string.Join(",", _ErrorList.ToArray()));
                    sb.AppendLine("1.樣板內科目合併欄位不足，請檢查樣板。");
                    sb.AppendLine("2.如果使用只有領域樣板，請忽略此訊息。");
                    if (_ErrorDomainNameList.Count > 0)
                        sb.AppendLine(string.Join(",", _ErrorDomainNameList.ToArray()));

                    FISCA.Presentation.Controls.MsgBox.Show(sb.ToString(), "樣板內科目合併欄位不足", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                }

                FISCA.Presentation.MotherForm.SetStatusBarMessage("評量成績報表產生完成");
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
            }
        }

        void _bgWorkReport_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 資料讀取
            _bgWorkReport.ReportProgress(1);

            // 每次合併後放入，最後再合成一張
            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.新竹評量成績單樣板_固定排名_科目_領域__doc1));

            _ErrorList.Clear();
            _ErrorDomainNameList.Clear();
            _TemplateSubjectNameList.Clear();

            // 取得樣板內科目名稱
            foreach (string name in docTemplate.MailMerge.GetFieldNames())
            {
                if (name != null && name.Contains("科目名稱"))
                    _TemplateSubjectNameList.Add(name);
            }

            // 校名
            string SchoolName = K12.Data.School.ChineseName;
            // 校長
            string ChancellorChineseName = JHSchool.Data.JHSchoolInfo.ChancellorChineseName;
            // 教務主任
            string EduDirectorName = JHSchool.Data.JHSchoolInfo.EduDirectorName;

            // 班級
            Dictionary<string, ClassRecord> ClassDict = new Dictionary<string, ClassRecord>();
            foreach (ClassRecord cr in Class.SelectAll())
                ClassDict.Add(cr.ID, cr);
            //// 教師
            //Dictionary<string, TeacherRecord> TeacherDict = new Dictionary<string, TeacherRecord>();
            //foreach (TeacherRecord tr in Teacher.SelectAll())
            //{
            //    if (tr.Status == TeacherRecord.TeacherStatus.一般)
            //        TeacherDict.Add(tr.ID, tr);
            //}

            // 不排名學生ID
            List<string> notRankStudIDList = new List<string>();
            if (_StudTagDict.ContainsKey(_SelNotRankedFilter))
                notRankStudIDList.AddRange(_StudTagDict[_SelNotRankedFilter].ToArray());

            // 移除所選學生不排名
            foreach (string id in notRankStudIDList)
                _StudentIDList.Remove(id);

            // 所選學生資料
            List<StudentRecord> StudRecList = Student.SelectByIDs(_StudentIDList);


            // 因為讀取固定排名，這段不需要即時處理
            //// 班級年級區分,沒有年級不處理
            //Dictionary<int, List<StudentRecord>> studGradeDict = new Dictionary<int, List<StudentRecord>>();
            //List<string> studIDAllList = new List<string>();
            //foreach (StudentRecord studRec in Student.SelectAll())
            //{
            //    // 不排名學生ID
            //    if (notRankStudIDList.Contains(studRec.ID))
            //        continue;

            //    if (studRec.Status == StudentRecord.StudentStatus.一般)
            //    {
            //        if (ClassDict.ContainsKey(studRec.RefClassID))
            //        {
            //            if (ClassDict[studRec.RefClassID].GradeYear.HasValue)
            //            {
            //                int gr = ClassDict[studRec.RefClassID].GradeYear.Value;

            //                if (!studGradeDict.ContainsKey(gr))
            //                    studGradeDict.Add(gr, new List<StudentRecord>());

            //                studIDAllList.Add(studRec.ID);
            //                studGradeDict[gr].Add(studRec);
            //            }
            //        }

            //    }
            //}
            _bgWorkReport.ReportProgress(15);

            #region 取得學生成績計算規則
            ScoreCalculator defaultScoreCalculator = new ScoreCalculator(null);


            List<string> scoreCalcRuleIDList = new List<string>();
            foreach (StudentRecord student in StudRecList)
            {
                //calcCache.Add(student.ID, new ScoreCalculator(student.ScoreCalcRule));
                string calcID = string.Empty;
                if (!string.IsNullOrEmpty(student.OverrideScoreCalcRuleID))
                    calcID = student.OverrideScoreCalcRuleID;
                else if (student.Class != null && !string.IsNullOrEmpty(student.Class.RefScoreCalcRuleID))
                    calcID = student.Class.RefScoreCalcRuleID;

                if (!string.IsNullOrEmpty(calcID))
                    calcIDCache.Add(student.ID, calcID);
            }
            foreach (JHScoreCalcRuleRecord record in JHScoreCalcRule.SelectByIDs(calcIDCache.Values))
            {
                if (!calcCache.ContainsKey(record.ID))
                    calcCache.Add(record.ID, new ScoreCalculator(record));
            }

            #endregion


            Dictionary<string, JHCourseRecord> CourseDict = new Dictionary<string, JHCourseRecord>();
            // 1.課程資料   取得當 學年 度學期 課程資訊
            foreach (JHCourseRecord co in JHCourse.SelectBySchoolYearAndSemester(_SelSchoolYear, _SelSemester))
            {
                CourseDict.Add(co.ID, co);
            }


            // 取評量成績
            Dictionary<string, List<HC.JHSCETakeRecord>> Score1Dict = new Dictionary<string, List<HC.JHSCETakeRecord>>();

            // Scetake取得所選學生評量成績  這邊有點冗 但不想動太多原本的邏輯 所以先不改
            // 取得定期評量資訊
            List<string> sceTakes; // sceTakeID 清單
            if (!this.HasReferenceExam) //如果不包含參考試別
            {
                sceTakes = Utility.GetSCETakeIDsByExamID(_StudentIDList, _SelSchoolYear, _SelSemester, this._SelExamID);
            }
            else
            {
                sceTakes = Utility.GetSCETakeIDsByExamID(_StudentIDList, _SelSchoolYear, _SelSemester, this._SelExamID, this._SelRefExamID);
            }

            // 用sce_take 取得學生修課 試別 成績 
            List<JHSCETakeRecord> studSCETake = JHSCETake.SelectByIDs(sceTakes);

            // todo 1.取得成績
            foreach (JHSCETakeRecord record in studSCETake) // 取得成績後給
            {
                if (record.RefExamID == _SelExamID || record.RefExamID == this._SelRefExamID) //過濾試別 (如果修課記錄上識別==選擇試別或參考識別)
                {
                    if (!Score1Dict.ContainsKey(record.RefStudentID))
                        Score1Dict.Add(record.RefStudentID, new List<HC.JHSCETakeRecord>());

                    Score1Dict[record.RefStudentID].Add(new HC.JHSCETakeRecord(record));
                }
            }

            // TODO 
            // 取得這次該修課程
            Dictionary<string, Dictionary<string, DAO.SubjectDomainName>> StudCourseDict = Utility.GetStudentSCAttendCourse(_StudentIDList, CourseDict.Keys.ToList(), _SelExamID);
            //  Dictionary<string, Dictionary<string, DAO.SubjectDomainName>> StudCourseDict = Utility.GetStudentSCAttendCourse(_StudentIDList, CourseDict.Keys.ToList(), _SelRefExamID);
            // 取得評量設定比例
            Dictionary<string, decimal> ScorePercentageHSDict = Utility.GetScorePercentageHS();

            _bgWorkReport.ReportProgress(30);



            // 處理評量成績科目
            Dictionary<string, DAO.StudExamScore> studExamScoreDict = new Dictionary<string, DAO.StudExamScore>();

            foreach (string studID in _StudentIDList)
            {
                // 成績計算規則
                ScoreCalculator studentCalculator = defaultScoreCalculator;
                if (calcIDCache.ContainsKey(studID) && calcCache.ContainsKey(calcIDCache[studID]))
                    studentCalculator = calcCache[calcIDCache[studID]];

                if (Score1Dict.ContainsKey(studID))
                {
                    if (!studExamScoreDict.ContainsKey(studID))
                        studExamScoreDict.Add(studID, new DAO.StudExamScore(studentCalculator));
                    // 取得成績
                    foreach (HC.JHSCETakeRecord rec in Score1Dict[studID])
                    {
                        if ((rec.RefExamID == _SelExamID || rec.RefExamID == _SelRefExamID) && CourseDict.ContainsKey(rec.RefCourseID))
                        {
                            JHCourseRecord cr = CourseDict[rec.RefCourseID];

                            string SubjecName = cr.Subject;

                            // 勾選科目
                            if (_SelSubjNameList.Contains(SubjecName))
                            {
                                DAO.ExamSubjectScore examSubjectScore;
                                if (!studExamScoreDict[studID]._ExamSubjectScoreDict.ContainsKey(SubjecName))
                                {
                                    examSubjectScore = new DAO.ExamSubjectScore();
                                    studExamScoreDict[studID]._ExamSubjectScoreDict.Add(SubjecName, examSubjectScore);
                                }
                                else
                                {
                                    examSubjectScore = (ExamSubjectScore)studExamScoreDict[studID]._ExamSubjectScoreDict[SubjecName];
                                }

                                    examSubjectScore.DomainName = cr.Domain;
                                    examSubjectScore.SubjectName = SubjecName;

                                    if (rec.RefExamID == this._SelExamID) //如果是本次選擇類別
                                    {
                                        examSubjectScore.ScoreA = rec.AssignmentScore;
                                        examSubjectScore.ScoreF = rec.Score;
                                    if (examSubjectScore.ScoreT.HasValue)
                                        examSubjectScore.ScoreT = studentCalculator.ParseSubjectScore(examSubjectScore.ScoreT.Value);

                                    examSubjectScore.Text = rec.Text;
                                    examSubjectScore.Credit = cr.Credit;

                                }
                                    else if (rec.RefExamID == this._SelRefExamID) //處理參考試別
                                    {
                                        examSubjectScore.RefScoreA = rec.AssignmentScore;
                                        examSubjectScore.RefScoreF = rec.Score;
                                    }

                                    if (ScorePercentageHSDict.ContainsKey(cr.RefAssessmentSetupID))
                                    {
                                        // 取得定期，評量由100-定期
                                        // decimal f = ScorePercentageHSDict[cr.RefAssessmentSetupID] * 0.01M;
                                        // 計算所有成績
                                        examSubjectScore.GetTotalScore(this.HasReferenceExam, ScorePercentageHSDict[cr.RefAssessmentSetupID]);
                                        //decimal a = (100 - ScorePercentageHSDict[cr.RefAssessmentSetupID]) * 0.01M;
                                    }
                                    else
                                    {
                                        examSubjectScore.GetTotalScore(this.HasReferenceExam, 0.5M);
                                    }

                                    //examSubjectScore.ScoreT = examSubjectScore.ScoreA.Value * a + examSubjectScore.ScoreF.Value * f;
                                    //}
                                    //else
                                    //    examSubjectScore.ScoreT = examSubjectScore.ScoreA.Value * 0.5M + examSubjectScore.ScoreF.Value * 0.5M; // 沒有設定預設50,50

                                    // 原本
                                    //ess.ScoreT = (ess.ScoreA.Value + ess.ScoreF.Value) / 2;

                                    //if (examSubjectScore.ScoreA.HasValue && examSubjectScore.ScoreF.HasValue == false)
                                    //    examSubjectScore.ScoreT = examSubjectScore.ScoreA.Value;

                                    //if (examSubjectScore.ScoreA.HasValue == false && examSubjectScore.ScoreF.HasValue)
                                    //    examSubjectScore.ScoreT = examSubjectScore.ScoreF.Value;

                                    // 依照成績計算規則科目方式處理進位，只有總成績。
                                    // 平時
                                    //if(ess.ScoreA.HasValue)
                                    //    ess.ScoreA = studentCalculator.ParseSubjectScore(ess.ScoreA.Value);

                                    // 定期
                                    //if (ess.ScoreF.HasValue)
                                    //    ess.ScoreF = studentCalculator.ParseSubjectScore(ess.ScoreF.Value);

                                    // 進位 
                            }
                        }
                    }
                    // 計算領域成績
                    studExamScoreDict[studID].CalcSubjectToDomain("");
                }

                if (StudCourseDict.ContainsKey(studID))
                {
                    if (!studExamScoreDict.ContainsKey(studID))
                        studExamScoreDict.Add(studID, new DAO.StudExamScore(studentCalculator));
                    if (studExamScoreDict[studID]._ExamSubjectScoreDict == null)
                    {
                        studExamScoreDict[studID]._ExamSubjectScoreDict = new Dictionary<string, IScore>();
                        studExamScoreDict[studID]._ExamDomainScoreDict = new Dictionary<string, IScore>();
                        if (this.HasReferenceExam)
                        {
                            studExamScoreDict[studID]._RefExamSubjectScoreDict = new Dictionary<string, IScore>();
                        }
                    }
                    foreach (KeyValuePair<string, DAO.SubjectDomainName> data in StudCourseDict[studID])
                    {
                        // 沒有勾選不加入
                        if (!_SelSubjNameList.Contains(data.Key))
                            continue;

                        // 加入有修課沒有成績空科目
                        if (!studExamScoreDict[studID]._ExamSubjectScoreDict.ContainsKey(data.Key))
                        {
                            DAO.ExamSubjectScore ess = new DAO.ExamSubjectScore();
                            ess.SubjectName = data.Key;
                            ess.DomainName = data.Value.DomainName;
                            ess.Credit = data.Value.Credit;
                            studExamScoreDict[studID]._ExamSubjectScoreDict.Add(data.Key, ess);
                        }
                    }
                }

            }

            List<string> tmpDomainList = new List<string>();

            List<string> tmpSubjectList = new List<string>();
            // 成績排序
            foreach (string studID in studExamScoreDict.Keys)
            {
                tmpDomainList.Clear();
                tmpSubjectList.Clear();
                Dictionary<string, IScore> tmpSubjectScore = new Dictionary<string, IScore>();
                Dictionary<string, IScore> tmpDomainScore = new Dictionary<string, IScore>();
                Dictionary<string, IScore> refSubject = new Dictionary<string, IScore>();

                foreach (string dname in studExamScoreDict[studID]._ExamDomainScoreDict.Keys)
                {
                    tmpDomainList.Add(dname);
                    tmpDomainScore.Add(dname, studExamScoreDict[studID]._ExamDomainScoreDict[dname]);
                }

                foreach (string sname in studExamScoreDict[studID]._ExamSubjectScoreDict.Keys)
                {
                    tmpSubjectList.Add(sname);
                    tmpSubjectScore.Add(sname, studExamScoreDict[studID]._ExamSubjectScoreDict[sname]);
                }

                // 排序
                //tmpDomainList.Sort(new StringComparer("語文", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動", ""));

                tmpDomainList.Sort(new StringComparer(Utility.GetDominOrder().ToArray()));


                // tmpSubjectList.Sort(new StringComparer("國文", "英文", "數學", "理化", "生物", "社會", "物理", "化學", "歷史", "地理", "公民"));
                tmpSubjectList.Sort(new StringComparer(Utility.GetSubjectOrder().ToArray()));




                studExamScoreDict[studID]._ExamDomainScoreDict.Clear();
                studExamScoreDict[studID]._ExamSubjectScoreDict.Clear();

                foreach (string dname in tmpDomainList)
                {
                    if (tmpDomainScore.ContainsKey(dname))
                        studExamScoreDict[studID]._ExamDomainScoreDict.Add(dname, tmpDomainScore[dname]);
                }

                foreach (string sname in tmpSubjectList)
                {
                    if (tmpSubjectScore.ContainsKey(sname))
                        studExamScoreDict[studID]._ExamSubjectScoreDict.Add(sname, tmpSubjectScore[sname]);
                }
            }


            _bgWorkReport.ReportProgress(45);

            // 缺曠資料區間統計
            _AttendanceDict = Utility.GetAttendanceCountByDate(StudRecList, _BeginDate, _EndDate);

            // 獎懲資料
            Dictionary<string, Dictionary<string, int>> DisciplineCountDict = Utility.GetDisciplineCountByDate(_StudentIDList, _BeginDate, _EndDate);

            // 服務學習
            Dictionary<string, decimal> ServiceLearningDict = Utility.GetServiceLearningDetailByDate(_StudentIDList, _BeginDate, _EndDate);

            // 取得父母監護人資訊
            Dictionary<string, JHParentRecord> ParentRecordDict = new Dictionary<string, JHParentRecord>();
            foreach (JHParentRecord rec in JHParent.SelectByStudentIDs(_StudentIDList))
            {
                if (!ParentRecordDict.ContainsKey(rec.RefStudentID))
                    ParentRecordDict.Add(rec.RefStudentID, rec);
            }

            // 取得地址相關資訊
            Dictionary<string, JHAddressRecord> AddressRecordDict = new Dictionary<string, JHAddressRecord>();
            foreach (JHAddressRecord rec in JHAddress.SelectByStudentIDs(_StudentIDList))
            {
                if (!AddressRecordDict.ContainsKey(rec.RefStudentID))
                    AddressRecordDict.Add(rec.RefStudentID, rec);
            }

            //Jean 取得生自訂欄位資訊
            Dictionary<string, Dictionary<string, string>> UserDefineDict = Utility.GetStudUserDefineInfo(_StudentIDList);




            // 領域組距
            List<string> li2 = new List<string>();
            li2.Add("R100_u");
            li2.Add("R90_99");
            li2.Add("R80_89");
            li2.Add("R70_79");
            li2.Add("R60_69");
            li2.Add("R50_59");
            li2.Add("R40_49");
            li2.Add("R30_39");
            li2.Add("R20_29");
            li2.Add("R10_19");
            li2.Add("R0_9");


            #region 處理排名資料

            // 學生排名資料索引使用
            Dictionary<string, Dictionary<string, RankDataInfo>> StudentExamRankMatrixDict = Utility.GetStudentExamRankDict(_SelSchoolYear, _SelSemester, _SelExamID, _StudentIDList);

            //// debug write text file
            //using (StreamWriter sw = new StreamWriter(Application.StartupPath + "\\debug排名資料.txt", false))
            //{
            //    foreach (string sid in StudentExamRankMatrixDict.Keys)
            //    {
            //        foreach (string key in StudentExamRankMatrixDict[sid].Keys)
            //        {
            //            sw.WriteLine("sid:" + sid + ",Key:" + key + ",value:" + StudentExamRankMatrixDict[sid][key]);
            //        }
            //    }
            //}



            // 學生排名資料參考索引使用
            Dictionary<string, Dictionary<string, RankDataInfo>> StudentRefExamRankMatrixDict = Utility.GetStudentExamRankDict(_SelRefSchoolYear, _SelRefSemester, _SelRefExamID, _StudentIDList);
            #endregion

            #endregion

            _bgWorkReport.ReportProgress(60);



            // 學生筆學期歷程
            Dictionary<string, SemesterHistoryItem> StudShiDict = new Dictionary<string, SemesterHistoryItem>();
            // 取得學期歷程，給班級、座號、班導師使用，條件：所選學生
            List<SemesterHistoryRecord> SemesterHistoryRecordList = SemesterHistory.SelectByStudentIDs(_StudentIDList);
            foreach (SemesterHistoryRecord shr in SemesterHistoryRecordList)
            {
                bool thisTimeHaveSemesterHistoryItem = false;

                foreach (SemesterHistoryItem shi in shr.SemesterHistoryItems)
                {
                    if (shi.SchoolYear == _SelSchoolYear && shi.Semester == _SelSemester)
                    {
                        if (!StudShiDict.ContainsKey(shi.RefStudentID))
                            StudShiDict.Add(shi.RefStudentID, shi);
                        thisTimeHaveSemesterHistoryItem = true;
                    }

                }
                if (!thisTimeHaveSemesterHistoryItem)
                {
                    if (!StudShiDict.ContainsKey(shr.RefStudentID))
                    {
                        SemesterHistoryItem New_shi = new SemesterHistoryItem();

                        New_shi.ClassName = shr.Student.Class != null ? shr.Student.Class.Name : "";

                        New_shi.SeatNo = shr.Student.SeatNo;

                        New_shi.Teacher = shr.Student.Class != null ? shr.Student.Class.Teacher.Name : "";

                        StudShiDict.Add(shr.RefStudentID, New_shi);

                    }
                }

            }




            #region 處理合併 DataTable 相關資料
            // 儲存資料用 Data Table
            DataTable dt = new DataTable();
            Document doc = new Document();
            DataTable dtAtt = new DataTable();
            List<Document> docList = new List<Document>();

            List<string> scoreTypeList = new List<string>();
            List<string> rankTypeList = new List<string>();
            List<string> rankDSTypeList = new List<string>();
            List<string> socreItem1List = new List<string>();

            List<string> rankDataTypeList = new List<string>();


            //socreItem1List.Add("領域成績");
            //socreItem1List.Add("領域定期成績");
            //socreItem1List.Add("參考領域成績");
            //socreItem1List.Add("參考領域定期成績");

            socreItem1List.Add("科目成績");
            socreItem1List.Add("科目定期成績");
            socreItem1List.Add("參考科目成績");
            socreItem1List.Add("參考科目定期成績");


            scoreTypeList.Add("加權平均");
            scoreTypeList.Add("加權總分");
            scoreTypeList.Add("平均");
            scoreTypeList.Add("總分");


            rankTypeList.Add("班排名");
            rankTypeList.Add("年排名");
            rankTypeList.Add("類別1排名");
            rankTypeList.Add("類別2排名");

            rankDSTypeList.Add("班級_");
            rankDSTypeList.Add("年級_");
            rankDSTypeList.Add("類別1_");
            rankDSTypeList.Add("類別2_");

            rankDataTypeList.Add("名次");
            rankDataTypeList.Add("PR值");
            rankDataTypeList.Add("百分比");
            rankDataTypeList.Add("母體頂標");
            rankDataTypeList.Add("母體前標");
            rankDataTypeList.Add("母體平均");
            rankDataTypeList.Add("母體後標");
            rankDataTypeList.Add("母體底標");
            rankDataTypeList.Add("母體人數");

            List<string> domainLi = new List<string>();

            List<string> subjLi = new List<string>();

            subjLi.Add("科目名稱");
            subjLi.Add("科目權數");
            // todo : 新增 參考 識別定期評量
            subjLi.Add("參考試別科目定期評量");
            subjLi.Add("參考試別科目平時評量");
            subjLi.Add("參考試別科目總成績");

            subjLi.Add("科目定期評量");
            subjLi.Add("科目平時評量");
            subjLi.Add("科目總成績");
            subjLi.Add("科目文字評量");
            subjLi.Add("科目定期評量等第");
            subjLi.Add("科目平時評量等第");
            subjLi.Add("科目總成績等第");

            // 新增排名資訊 
            foreach (string rt in rankTypeList)
            {
                foreach (string rdt in rankDataTypeList)
                {
                    // 總成績
                    subjLi.Add("科目" + rt + rdt);
                    // 定期
                    subjLi.Add("科目定期" + rt + rdt);
                }
            }


            List<string> subjColList = new List<string>();
            foreach (string dName in Global.DomainNameList)
            {
                for (int i = 1; i <= 12; i++)
                {
                    foreach (string sName in subjLi)
                    {
                        string key = dName + "_" + sName + i;
                        subjColList.Add(key);
                    }
                }
            }

            // 填值
            foreach (StudentRecord StudRec in StudRecList)
            {

                dtAtt.Columns.Clear();
                dtAtt.Clear();
                dt.Clear();
                dt.Columns.Clear();

                dtAtt.Columns.Add("缺曠紀錄");
                DataRow rowT = dtAtt.NewRow();

                // 取得年級
                int grYear = 0;

                if (ClassDict.ContainsKey(StudRec.RefClassID))
                {
                    if (ClassDict[StudRec.RefClassID].GradeYear.HasValue)
                        grYear = ClassDict[StudRec.RefClassID].GradeYear.Value;
                }

                // 缺曠欄位
                foreach (var type in typeList)
                {
                    foreach (var absence in absenceList)
                    {
                        dt.Columns.Add(type + absence);
                    }
                }

                // 新增合併欄位
                dt.Columns.Add("系統編號");
                dt.Columns.Add("StudentID");
                dt.Columns.Add("學校名稱");
                dt.Columns.Add("學年度");
                dt.Columns.Add("學期");
                dt.Columns.Add("試別名稱");
                dt.Columns.Add("參考試別");
                dt.Columns.Add("班級");
                dt.Columns.Add("學號");
                dt.Columns.Add("座號");
                dt.Columns.Add("姓名");
                dt.Columns.Add("監護人姓名");
                dt.Columns.Add("父親姓名");
                dt.Columns.Add("母親姓名");
                dt.Columns.Add("戶籍地址");
                dt.Columns.Add("聯絡地址");
                dt.Columns.Add("其他地址");
                dt.Columns.Add("領域成績加權平均");
                dt.Columns.Add("領域定期成績加權平均");
                dt.Columns.Add("科目定期成績加權平均");
                dt.Columns.Add("科目平時成績加權平均");
                //dt.Columns.Add("科目總成績加權平均");
                dt.Columns.Add("科目成績加權平均");
                dt.Columns.Add("領域成績加權平均(不含彈性)");
                dt.Columns.Add("科目定期成績加權平均(不含彈性)");
                dt.Columns.Add("科目平時成績加權平均(不含彈性)");
                //dt.Columns.Add("科目總成績加權平均(不含彈性)");
                dt.Columns.Add("科目成績加權平均(不含彈性)");
                //dt.Columns.Add("科目成績平均(不含彈性)");
                dt.Columns.Add("領域成績加權總分");
                dt.Columns.Add("領域定期成績加權總分");
                dt.Columns.Add("科目定期成績加權總分");
                dt.Columns.Add("科目平時成績加權總分");
                dt.Columns.Add("科目成績加權總分");
                //dt.Columns.Add("科目成績總分");
                dt.Columns.Add("領域成績加權總分(不含彈性)");
                dt.Columns.Add("科目定期成績加權總分(不含彈性)");
                dt.Columns.Add("科目平時成績加權總分(不含彈性)"); 
               // dt.Columns.Add("科目總成績加權總分(不含彈性)"); 

                dt.Columns.Add("科目平時評量總分(不含彈性)");
                dt.Columns.Add("科目成績加權總分(不含彈性)"); 
               // dt.Columns.Add("科目成績總分(不含彈性)"); 
                dt.Columns.Add("領域成績總分");
                dt.Columns.Add("領域定期成績總分");
                dt.Columns.Add("領域成績平均");
                dt.Columns.Add("領域定期成績平均");

               

                #region 在dt加上科目相關header

                //增加科目相關欄位 by Jean *以上科目相關先註解
                string scoreTarget = "科目";
                //string[] ScoreCompositions = new string[] { "定期評量", "平時評量", "總成績" };
                string[] ScoreCaculateWays = new string[] { "平均", "平均(不含彈性)", "總分", "總分(不含彈性)" };

                foreach (EnumScoreComposition scoreComposition in Enum.GetValues(_ScoreCompositions)) // 為保持一致性 從列舉裡面取得用詞
                {
                    foreach (string scoreCaculateWay in ScoreCaculateWays)
                    {
                        dt.Columns.Add($"{scoreTarget}{scoreComposition}{scoreCaculateWay}");
                    }
                }


                #endregion

                foreach (string itemName in itemNames) //平均 總分 => 基本上這兩個 就是從科目來的
                {
                    foreach (string rankType in rankTypeList) // 年排名 班排名 類別1排名 類別2排名 
                    {
                        foreach (string itemType in itemTypesMapping.Keys)
                        {

                          

                        }
                    }
                }





                // 新增自訂欄位 依華商需求 增加自訂欄位
                foreach (string userDefineField in this._UserDefineFields)
                {
                    dt.Columns.Add(userDefineField + "-自訂欄位");
                }



                // 獎懲名稱
                foreach (string str in Global.GetDisciplineNameList())
                    dt.Columns.Add(str + "區間統計");

                dt.Columns.Add("缺曠紀錄");
                dt.Columns.Add("服務學習時數");
                dt.Columns.Add("校長");
                dt.Columns.Add("教務主任");
                dt.Columns.Add("班導師");
                dt.Columns.Add("區間開始日期");
                dt.Columns.Add("區間結束日期");
                dt.Columns.Add("成績校正日期");

                // todo 增加自訂欄位 全部抓出來








                // todo 
                // 新增科目成績欄位
                foreach (string colName in subjColList)
                    dt.Columns.Add(colName);

                //Global.DomainNameList.Sort(new StringComparer("語文", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"));
                Global.DomainNameList.Sort(new StringComparer(Utility.GetDominOrder().ToArray()));

                // 新增各領域成績與排名，與佳樺討論，個人評量各領域只提供從科目計算來的加權平均
                foreach (string dName in Global.DomainNameList)
                {
                    dt.Columns.Add(dName + "_領域權數");
                    dt.Columns.Add(dName + "_領域加權平均");
                    dt.Columns.Add(dName + "_領域定期加權平均");
                    dt.Columns.Add(dName + "_領域平時加權平均");
                    dt.Columns.Add(dName + "_領域加權平均等第");
                    dt.Columns.Add(dName + "_領域定期加權平均等第");
                    dt.Columns.Add(dName + "_領域平時加權平均等第");

                    // 領域五標
                    foreach (string rt in rankTypeList)
                    {
                        foreach (string rdt in rankDataTypeList)
                        {
                            dt.Columns.Add(dName + "_領域" + rt + rdt);
                            dt.Columns.Add(dName + "_領域定期" + rt + rdt);
                        }
                    }

                    // 領域組距
                    foreach (string rt in rankDSTypeList)
                    {
                        foreach (string li in li2)
                        {
                            dt.Columns.Add(rt + dName + "_" + li);
                            dt.Columns.Add(rt + dName + "F_" + li);
                        }
                    }

                    // 領域-科目1,2,3 組距
                    for (int subjIdx = 1; subjIdx <= 7; subjIdx++)
                    {
                        foreach (string rt in rankDSTypeList)
                        {
                            dt.Columns.Add(rt + dName + "領域_科目名稱" + subjIdx);
                            foreach (string li in li2)
                            {
                                dt.Columns.Add(rt + dName + "領域_科目" + subjIdx + "_" + li);
                                dt.Columns.Add(rt + dName + "領域_科目" + subjIdx + "F_" + li);
                            }
                        }
                    }


                }

                // todo

                // 加入科目用組距
                foreach (string rt in rankDSTypeList)
                {
                    for (int si = 1; si <= 30; si++)
                    {
                        dt.Columns.Add("s" + rt + "科目名稱" + si);
                        dt.Columns.Add("sf" + rt + "科目名稱" + si);

                        foreach (string li in li2)
                        {
                            dt.Columns.Add("s" + rt + "科目" + si + "_" + li);
                            dt.Columns.Add("sf" + rt + "科目" + si + "_" + li);
                        }

                    }
                }

                // 總領域加權排名欄位
                foreach (string st in scoreTypeList)
                {
                    // 總領域排名欄位
                    foreach (string rt in rankTypeList)
                    {
                        foreach (string si in socreItem1List)
                        {
                            foreach (string rdt in rankDataTypeList)
                            {
                                dt.Columns.Add(si + st + rt + rdt);
                            }

                            // 加組距
                            if (st.Contains("平均"))
                            {
                                foreach (string li in li2)
                                {
                                    dt.Columns.Add(si + st + rt + li);
                                }
                            }
                        }
                    }
                }


                dt.TableName = StudRec.ID;

                //// debug write text file
                using (StreamWriter sw = new StreamWriter(Application.StartupPath + "\\debug_TestSc.txt", false))
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        sw.WriteLine(dc.Caption);
                    }
                }


                // 填入對照資料
                DataRow row = dt.NewRow();


                row["系統編號"] = "系統編號{" + StudRec.ID + "}";
                row["StudentID"] = StudRec.ID;
                row["學校名稱"] = SchoolName;
                row["學年度"] = _SelSchoolYear;
                row["學期"] = _SelSemester;
                row["試別名稱"] = _SelExamName;
                row["參考試別"] = _SelRefExamName;

                if (ParentRecordDict.ContainsKey(StudRec.ID))
                {
                    row["監護人姓名"] = ParentRecordDict[StudRec.ID].CustodianName;
                    row["父親姓名"] = ParentRecordDict[StudRec.ID].FatherName;
                    row["母親姓名"] = ParentRecordDict[StudRec.ID].MotherName;
                }
                if (AddressRecordDict.ContainsKey(StudRec.ID))
                {
                    row["戶籍地址"] = AddressRecordDict[StudRec.ID].PermanentAddress;
                    row["聯絡地址"] = AddressRecordDict[StudRec.ID].MailingAddress;
                    row["其他地址"] = AddressRecordDict[StudRec.ID].Address1Address;
                }

                // 班級、座號、班導師 使用學期歷程內
                if (StudShiDict.ContainsKey(StudRec.ID))
                {
                    row["班級"] = StudShiDict[StudRec.ID].ClassName;
                    row["座號"] = StudShiDict[StudRec.ID].SeatNo;
                    row["班導師"] = StudShiDict[StudRec.ID].Teacher;
                }

                row["學號"] = StudRec.StudentNumber;
                row["姓名"] = StudRec.Name;


                // 新增自訂欄位資訊(動態產生)

                if (StudShiDict.ContainsKey(StudRec.ID))
                {
                    foreach (string fieldName in this._UserDefineFields)
                    {
                        if (UserDefineDict.ContainsKey(StudRec.ID) && UserDefineDict[StudRec.ID].ContainsKey(fieldName))
                        {
                            row[fieldName + "-自訂欄位"] = UserDefineDict[StudRec.ID][fieldName]; // 取得某學生下某自訂欄位之值
                        }
                        else
                        {
                            row[fieldName + "-自訂欄位"] = "";
                        }
                    }
                }


                // 傳入 ID當 Key
                // row["缺曠紀錄"] = StudRec.ID;
                rowT["缺曠紀錄"] = StudRec.ID;
                //缺曠套印
                foreach (var type in typeList)
                {
                    foreach (var absence in absenceList)
                    {
                        row[type + absence] = "0";
                    }
                }
                if (_AttendanceDict.ContainsKey(StudRec.ID))
                {
                    foreach (var absentKey in _AttendanceDict[StudRec.ID].Keys)
                    {
                        if (_AttendanceDict[StudRec.ID].ContainsKey(absentKey))
                        {
                            if (dt.Columns.Contains(absentKey))
                                row[absentKey] = _AttendanceDict[StudRec.ID][absentKey];
                        }

                    }
                }
                // 獎懲區間統計值
                if (DisciplineCountDict.ContainsKey(StudRec.ID))
                {
                    foreach (string str in Global.GetDisciplineNameList())
                    {
                        string key = str + "區間統計";
                        if (DisciplineCountDict[StudRec.ID].ContainsKey(str))
                            row[key] = DisciplineCountDict[StudRec.ID][str];
                    }
                }

                int rnkSubj = 1;
                // 處理成績套印
                if (studExamScoreDict.ContainsKey(StudRec.ID))
                {
                    // 科目
                    int subj = 1;
                    Dictionary<string, int> dNameDict = new Dictionary<string, int>();
                    foreach (string name in Global.DomainNameList)
                        dNameDict.Add(name, 1);

                    foreach (string dicKey in studExamScoreDict[StudRec.ID]._ExamSubjectScoreDict.Keys)
                    {
                        ExamSubjectScore examSubjScore = (ExamSubjectScore)studExamScoreDict[StudRec.ID]._ExamSubjectScoreDict[dicKey];

                        string ddname = examSubjScore.DomainName;
                        if (ddname == "")
                            ddname = "彈性課程";

                        // 當領域非固定領域，無法處理
                        if (!Global.DomainNameList.Contains(ddname))
                        {
                            string errMsg = ddname + "領域相關成績，無法處理。";
                            if (!_ErrorDomainNameList.Contains(errMsg))
                                _ErrorDomainNameList.Add(errMsg);
                            continue;
                        }

                        if (dNameDict.ContainsKey(ddname))
                        {
                            subj = dNameDict[ddname];
                            dNameDict[ddname]++;
                        }

                        //Jean 參考
                        //foreach (string item in subjLi)
                        //{
                        //    if (item.Contains("參考"))
                        //    {
                        //        string key = ddname + "_" + item + subj;

                        //        switch (item)
                        //        {
                        //            case "參考試別科目定期評量":
                        //                if (ess.ScoreF.HasValue)
                        //                    row[key] = studExamScoreDict[StudRec.ID].ess.;
                        //                break;
                        //            case "參考試別科目總成績":
                        //                if (ess.ScoreF.HasValue)
                        //                    row[key] = "test2";
                        //                //row[key] = _ScoreMappingConfig.ParseScoreName(ess.ScoreF.Value);

                        //                break;
                        //            case "參考試別科目平時評量":
                        //                if (ess.ScoreA.HasValue)
                        //                    row[key] = "test3";
                        //                //row[key] = ess.ScoreA.Value;
                        //                break;

                        //        }
                        //    }
                        //}

                        // TODO:  取得參考識別之資料

                        foreach (string item in subjLi)
                        {
                            string key = ddname + "_" + item + subj;

                            switch (item)
                            {
                                case "科目名稱":
                                    row[key] = examSubjScore.SubjectName;
                                    if (!_TemplateSubjectNameList.Contains(key))
                                    {
                                        if (!_ErrorList.Contains(key))
                                            _ErrorList.Add(key);
                                    }
                                    break;
                                case "科目權數":
                                    if (examSubjScore.Credit.HasValue)
                                        row[key] = examSubjScore.Credit.Value;
                                    break;
                                case "科目定期評量":
                                    if (examSubjScore.ScoreF.HasValue)
                                        row[key] = doParseTransfer(examSubjScore.ScoreF.Value);
                                    break;
                                case "科目定期評量等第":
                                    if (examSubjScore.ScoreF.HasValue)
                                        row[key] = _ScoreMappingConfig.ParseScoreName(examSubjScore.ScoreF.Value);
                                    break;
                                case "科目平時評量":
                                    if (examSubjScore.ScoreA.HasValue)
                                        row[key] = doParseTransfer(examSubjScore.ScoreA.Value);
                                    break;
                                case "科目平時評量等第":
                                    if (examSubjScore.ScoreA.HasValue)
                                        row[key] = _ScoreMappingConfig.ParseScoreName(examSubjScore.ScoreA.Value);
                                    break;
                                case "科目總成績":
                                    if (examSubjScore.ScoreT.HasValue)
                                        row[key] = doParseTransfer(examSubjScore.ScoreT.Value);
                                    break;
                                case "科目總成績等第":
                                    if (examSubjScore.ScoreT.HasValue)
                                        row[key] = _ScoreMappingConfig.ParseScoreName(examSubjScore.ScoreT.Value);
                                    break;
                                case "科目文字評量":
                                    row[key] = examSubjScore.Text;
                                    break;

                                case "參考試別科目定期評量":
                                    if(cboSchoolYear.Text != cboRefSchoolYear.Text || cboSemester.Text != cboRefSemester.Text)
                                    {
                                        row[key] = null;
                                        break;
                                    }
                                    else if (examSubjScore.ScoreF.HasValue)
                                    {
                                        if(examSubjScore.RefScoreF != null)
                                        {
                                            row[key] = doParseTransfer(examSubjScore.RefScoreF.Value);
                                        }
                                        else
                                        {
                                            //row[key] = examSubjScore.RefScoreF;
                                            row[key] = doParseTransfer(examSubjScore.RefScoreF.Value);
                                        }
                                    }
                                        
                                    break;

                                case "參考試別科目平時評量":
                                    if (cboSchoolYear.Text != cboRefSchoolYear.Text || cboSemester.Text != cboRefSemester.Text)
                                    {
                                        row[key] = null;
                                        break;
                                    }
                                    if (examSubjScore.RefScoreA.HasValue)
                                    {
                                        if (examSubjScore.RefScoreF != null)
                                        {
                                            row[key] = doParseTransfer(examSubjScore.RefScoreA.Value);
                                        }
                                        else
                                        {
                                            //row[key] = examSubjScore.RefScoreA;
                                            row[key] = doParseTransfer(examSubjScore.RefScoreA.Value);
                                        }
                                    }
   
                                    break;
                                case "參考試別科目總成績":
                                    if (cboSchoolYear.Text != cboRefSchoolYear.Text || cboSemester.Text != cboRefSemester.Text)
                                    {
                                        row[key] = null;
                                        break;
                                    }
                                    if (examSubjScore.RefScoreT.HasValue)
                                    {
                                        if (examSubjScore.RefScoreF != null)
                                        {
                                            row[key] = doParseTransfer(examSubjScore.RefScoreT.Value);
                                        }
                                        else
                                        {
                                            row[key] = doParseTransfer(examSubjScore.RefScoreT.Value);
                                        }
                                    }
                                    break;
                            }

                            // 處理參考試別成績 

                            // string refkey = ddname + "_" + item + subj;

                            //switch (item)
                            //{
                            //    case "參考試別科目定期評量":
                            //        if (examSubjScore.ScoreF.HasValue)
                            //            row[key] = examSubjScore.RefScoreF;
                            //        break;

                            //    case "參考試別科目平時評量":
                            //        if (examSubjScore.RefScoreA.HasValue)
                            //            row[key] = examSubjScore.RefScoreA;
                            //        //row[key] = ess.ScoreA.Value;
                            //        break;
                            //    case "參考試別科目總成績":
                            //        if (examSubjScore.RefScoreT.HasValue)
                            //            row[key] = examSubjScore.RefScoreT;
                            //        //row[key] = _ScoreMappingConfig.ParseScoreName(ess.ScoreF.Value);
                            //        break;
                            //}
                        }


                        if (StudentExamRankMatrixDict.ContainsKey(StudRec.ID))
                        {
                            foreach (string rt in rankTypeList)
                            {
                                string keyD = "定期評量/科目成績" + examSubjScore.SubjectName + rt;
                                if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD))
                                {
                                    // todo 
                                    row[ddname + "_科目" + rt + "名次" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].rank;
                                    row[ddname + "_科目" + rt + "PR值" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].pr;
                                    row[ddname + "_科目" + rt + "百分比" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].percentile;
                                    row[ddname + "_科目" + rt + "母體頂標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_25;
                                    row[ddname + "_科目" + rt + "母體前標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_50;
                                    row[ddname + "_科目" + rt + "母體平均" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg;
                                    row[ddname + "_科目" + rt + "母體後標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_50;
                                    row[ddname + "_科目" + rt + "母體底標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_25;
                                    row[ddname + "_科目" + rt + "母體人數" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].matrix_count;
                                }

                                keyD = "定期評量_定期/科目成績" + examSubjScore.SubjectName + rt;
                                if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD))
                                {
                                    row[ddname + "_科目定期" + rt + "名次" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].rank;
                                    row[ddname + "_科目定期" + rt + "PR值" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].pr;
                                    row[ddname + "_科目定期" + rt + "百分比" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].percentile;
                                    row[ddname + "_科目定期" + rt + "母體頂標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_25;
                                    row[ddname + "_科目定期" + rt + "母體前標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_50;
                                    row[ddname + "_科目定期" + rt + "母體平均" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg;
                                    row[ddname + "_科目定期" + rt + "母體後標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_50;
                                    row[ddname + "_科目定期" + rt + "母體底標" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_25;
                                    row[ddname + "_科目定期" + rt + "母體人數" + subj] = StudentExamRankMatrixDict[StudRec.ID][keyD].matrix_count;
                                }
                            }
                        }

                    }

                    //  Jean  
                    // 處理領域-科目成績
                    Dictionary<string, List<DAO.ExamSubjectScore>> DomainSubjectDict = new Dictionary<string, List<ExamSubjectScore>>();

                    foreach (DAO.ExamSubjectScore ess in studExamScoreDict[StudRec.ID]._ExamSubjectScoreDict.Values)
                    {
                        string dname = "";
                        if (ess.DomainName == "")
                            dname = "彈性課程";
                        else
                            dname = ess.DomainName;



                        if (!DomainSubjectDict.ContainsKey(dname))
                            DomainSubjectDict.Add(dname, new List<ExamSubjectScore>());

                        DomainSubjectDict[dname].Add(ess);
                    }
                    // 領域-科目
                    int subdIdx = 1;
                    foreach (string dname in DomainSubjectDict.Keys)
                    {
                        subdIdx = 1;
                        foreach (DAO.ExamSubjectScore ess in DomainSubjectDict[dname])
                        {
                            if (StudentExamRankMatrixDict.ContainsKey(StudRec.ID))
                            {
                                string rtSubjet = "";
                                // 處理單存科目組距
                                foreach (string rt in rankTypeList)
                                {
                                    if (rt == "班排名")
                                        rtSubjet = "班級_";

                                    if (rt == "年排名")
                                        rtSubjet = "年級_";

                                    if (rt == "類別1排名")
                                        rtSubjet = "類別1_";

                                    if (rt == "類別2排名")
                                        rtSubjet = "類別2_";

                                    string keyD1 = "定期評量/科目成績" + ess.SubjectName + rt;
                                    if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD1))
                                    {
                                        //班級_語文領域_科目1_R100_u
                                        //班級_語文領域_科目名稱1
                                        row[rtSubjet + dname + "領域_科目名稱" + subdIdx] = ess.SubjectName;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R100_u"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_gte100;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R90_99"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_90;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R80_89"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_80;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R70_79"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_70;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R60_69"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_60;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R50_59"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_50;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R40_49"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_40;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R30_39"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_30;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R20_29"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_20;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R10_19"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_10;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "_R0_9"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_lt10;
                                    }

                                    keyD1 = "定期評量_定期/科目成績" + ess.SubjectName + rt;
                                    if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD1))
                                    {
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R100_u"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_gte100;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R90_99"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_90;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R80_89"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_80;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R70_79"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_70;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R60_69"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_60;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R50_59"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_50;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R40_49"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_40;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R30_39"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_30;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R20_29"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_20;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R10_19"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_10;
                                        row[rtSubjet + dname + "領域_科目" + subdIdx + "F_R0_9"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_lt10;
                                    }
                                }
                            }

                            subdIdx++;
                        }
                    }


                    // 處理單獨科目組距
                    List<string> rankSubjectList = new List<string>();
                    foreach (DAO.ExamSubjectScore ess in studExamScoreDict[StudRec.ID]._ExamSubjectScoreDict.Values)
                    {
                        if (!rankSubjectList.Contains(ess.SubjectName))
                            rankSubjectList.Add(ess.SubjectName);
                    }

                    rankSubjectList.Sort(new StringComparer(Utility.GetSubjectOrder().ToArray()));


                    if (StudentExamRankMatrixDict.ContainsKey(StudRec.ID))
                    {
                        foreach (string subjName in rankSubjectList)
                        {
                            string rtSubjet = "";


                            // 處理單存科目組距
                            foreach (string rt in rankTypeList)
                            {
                                if (rt == "班排名")
                                    rtSubjet = "班級_";

                                if (rt == "年排名")
                                    rtSubjet = "年級_";

                                if (rt == "類別1排名")
                                    rtSubjet = "類別1_";

                                if (rt == "類別2排名")
                                    rtSubjet = "類別2_";

                                string keyD1 = "定期評量/科目成績" + subjName + rt;
                                if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD1))
                                {
                                    row["s" + rtSubjet + "科目名稱" + rnkSubj] = subjName;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R100_u"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_gte100;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R90_99"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_90;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R80_89"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_80;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R70_79"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_70;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R60_69"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_60;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R50_59"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_50;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R40_49"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_40;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R30_39"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_30;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R20_29"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_20;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R10_19"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_10;
                                    row["s" + rtSubjet + "科目" + rnkSubj + "_R0_9"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_lt10;
                                }

                                keyD1 = "定期評量_定期/科目成績" + subjName + rt;
                                if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD1))
                                {
                                    row["sf" + rtSubjet + "科目名稱" + rnkSubj] = subjName;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R100_u"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_gte100;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R90_99"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_90;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R80_89"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_80;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R70_79"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_70;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R60_69"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_60;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R50_59"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_50;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R40_49"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_40;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R30_39"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_30;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R20_29"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_20;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R10_19"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_10;
                                    row["sf" + rtSubjet + "科目" + rnkSubj + "_R0_9"] = StudentExamRankMatrixDict[StudRec.ID][keyD1].level_lt10;
                                }
                            }

                            string keyD2 = "定期評量/科目成績" + subjName + "班排名";
                            if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD2))
                            {
                                rnkSubj++;
                            }
                        }
                    }



                    // 領域
                    foreach (DAO.ExamDomainScore eds in studExamScoreDict[StudRec.ID]._ExamDomainScoreDict.Values)
                    {
                        // 當領域非固定領域，無法處理
                        if (!Global.DomainNameList.Contains(eds.DomainName))
                        {
                            string errMsg = eds.DomainName + "領域相關成績，無法處理。";
                            if (!_ErrorDomainNameList.Contains(errMsg))
                                _ErrorDomainNameList.Add(errMsg);
                            continue;
                        }

                        string key = eds.DomainName + "_領域加權平均";
                        string keya = eds.DomainName + "_領域平時加權平均";
                        string keyf = eds.DomainName + "_領域定期加權平均";
                        string keyc = eds.DomainName + "_領域權數";



                        // 總成績
                        if (eds.ScoreT.HasValue)
                        {
                            row[key] = doParseTransfer(eds.ScoreT.Value);
                            row[key + "等第"] = _ScoreMappingConfig.ParseScoreName(eds.ScoreT.Value);
     
                        }


                        // 學分
                        if (eds.Credit1.HasValue)
                            row[keyc] = eds.Credit1.Value;

                        // 平時加權
                        if (eds.ScoreA.HasValue)
                        {
                            row[keya] = doParseTransfer(eds.ScoreA.Value);
                            row[keya + "等第"] = _ScoreMappingConfig.ParseScoreName(eds.ScoreA.Value);
                        }


                        // 定期加權
                        if (eds.ScoreF.HasValue)
                        {
                            row[keyf] = doParseTransfer(eds.ScoreF.Value);
                            row[keyf + "等第"] = _ScoreMappingConfig.ParseScoreName(eds.ScoreF.Value);
                        }


                        if (StudentExamRankMatrixDict.ContainsKey(StudRec.ID))
                        {
                            foreach (string rt in rankTypeList)
                            {
                                string keyD = "定期評量/領域成績" + eds.DomainName + rt;
                                if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD))
                                {
                                    row[eds.DomainName + "_領域" + rt + "名次"] = StudentExamRankMatrixDict[StudRec.ID][keyD].rank;
                                    row[eds.DomainName + "_領域" + rt + "PR值"] = StudentExamRankMatrixDict[StudRec.ID][keyD].pr;
                                    row[eds.DomainName + "_領域" + rt + "百分比"] = StudentExamRankMatrixDict[StudRec.ID][keyD].percentile;
                                    row[eds.DomainName + "_領域" + rt + "母體頂標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_25;
                                    row[eds.DomainName + "_領域" + rt + "母體前標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_50;
                                    row[eds.DomainName + "_領域" + rt + "母體平均"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg;
                                    row[eds.DomainName + "_領域" + rt + "母體後標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_50;
                                    row[eds.DomainName + "_領域" + rt + "母體底標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_25;
                                    row[eds.DomainName + "_領域" + rt + "母體人數"] = StudentExamRankMatrixDict[StudRec.ID][keyD].matrix_count;

                                    string rowRt = "班級_";
                                    if (rt == "年排名")
                                        rowRt = "年級_";

                                    if (rt == "類別1排名")
                                        rowRt = "類別1_";

                                    if (rt == "類別2排名")
                                        rowRt = "類別2_";

                                    row[rowRt + eds.DomainName + "_R100_u"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_gte100;
                                    row[rowRt + eds.DomainName + "_R90_99"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_90;
                                    row[rowRt + eds.DomainName + "_R80_89"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_80;
                                    row[rowRt + eds.DomainName + "_R70_79"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_70;
                                    row[rowRt + eds.DomainName + "_R60_69"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_60;
                                    row[rowRt + eds.DomainName + "_R50_59"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_50;
                                    row[rowRt + eds.DomainName + "_R40_49"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_40;
                                    row[rowRt + eds.DomainName + "_R30_39"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_30;
                                    row[rowRt + eds.DomainName + "_R20_29"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_20;
                                    row[rowRt + eds.DomainName + "_R10_19"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_10;
                                    row[rowRt + eds.DomainName + "_R0_9"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_lt10;
                                }


                                keyD = "定期評量_定期/領域成績" + eds.DomainName + rt;
                                if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(keyD))
                                {

                                    row[eds.DomainName + "_領域定期" + rt + "名次"] = StudentExamRankMatrixDict[StudRec.ID][keyD].rank;
                                    row[eds.DomainName + "_領域定期" + rt + "PR值"] = StudentExamRankMatrixDict[StudRec.ID][keyD].pr;
                                    row[eds.DomainName + "_領域定期" + rt + "百分比"] = StudentExamRankMatrixDict[StudRec.ID][keyD].percentile;
                                    row[eds.DomainName + "_領域定期" + rt + "母體頂標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_25;
                                    row[eds.DomainName + "_領域定期" + rt + "母體前標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_top_50;
                                    row[eds.DomainName + "_領域定期" + rt + "母體平均"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg;
                                    row[eds.DomainName + "_領域定期" + rt + "母體後標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_50;
                                    row[eds.DomainName + "_領域定期" + rt + "母體底標"] = StudentExamRankMatrixDict[StudRec.ID][keyD].avg_bottom_25;
                                    row[eds.DomainName + "_領域定期" + rt + "母體人數"] = StudentExamRankMatrixDict[StudRec.ID][keyD].matrix_count;


                                    string rowRt = "班級_";
                                    if (rt == "年排名")
                                        rowRt = "年級_";

                                    if (rt == "類別1排名")
                                        rowRt = "類別1_";

                                    if (rt == "類別2排名")
                                        rowRt = "類別2_";

                                    row[rowRt + eds.DomainName + "F_R100_u"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_gte100;
                                    row[rowRt + eds.DomainName + "F_R90_99"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_90;
                                    row[rowRt + eds.DomainName + "F_R80_89"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_80;
                                    row[rowRt + eds.DomainName + "F_R70_79"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_70;
                                    row[rowRt + eds.DomainName + "F_R60_69"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_60;
                                    row[rowRt + eds.DomainName + "F_R50_59"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_50;
                                    row[rowRt + eds.DomainName + "F_R40_49"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_40;
                                    row[rowRt + eds.DomainName + "F_R30_39"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_30;
                                    row[rowRt + eds.DomainName + "F_R20_29"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_20;
                                    row[rowRt + eds.DomainName + "F_R10_19"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_10;
                                    row[rowRt + eds.DomainName + "F_R0_9"] = StudentExamRankMatrixDict[StudRec.ID][keyD].level_lt10;
                                }

                            }
                        }
                    }
                }

                // 加權平均
                if (studExamScoreDict.ContainsKey(StudRec.ID))
                {
                    if (studExamScoreDict[StudRec.ID].GetDomainWAvgScoreA(true).HasValue)
                        row["領域成績加權平均"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainWAvgScoreA(true).Value);

                    if (studExamScoreDict[StudRec.ID].GetDomainWAvgScoreF(true).HasValue)
                        row["領域定期成績加權平均"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainWAvgScoreF(true).Value);

                    //// Jean 增加領域算術平均 
                    //if (studExamScoreDict[StudRec.ID].GetDomainArithmeticMeanScoreA(true).HasValue)
                    //    row["領域成績算數平均"] = studExamScoreDict[StudRec.ID].GetDomainWAvgScoreA(true).Value;

                    //// Jean 增加領域算術平均 
                    //if (studExamScoreDict[StudRec.ID].GetDomainArithmeticMeanScoreF(true).HasValue)
                    //    row["領域定期成績算數平均"] = studExamScoreDict[StudRec.ID].GetDomainWAvgScoreF(true).Value;


                    //科目 加權平均 (含彈性)
                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreAA(true, "").HasValue)
                        row["科目平時成績加權平均"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreAA(true, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreAF(true, "").HasValue)
                        row["科目定期成績加權平均"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreAF(true, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreAT(true, "").HasValue)
                        row["科目成績加權平均"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreAT(true, "").Value);

                    // 科目 加權平均(不含彈性)
                    if (studExamScoreDict[StudRec.ID].GetDomainWAvgScoreA(false).HasValue)
                        row["領域成績加權平均(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainWAvgScoreA(false).Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreAA(false, "").HasValue)
                        row["科目平時成績加權平均(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreAA(false, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreAF(false, "").HasValue)
                        row["科目定期成績加權平均(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreAF(false, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreAT(false, "").HasValue)
                        //row["科目總成績加權平均(不含彈性)"] = studExamScoreDict[StudRec.ID].GetSubjectScoreAT(false,"").Value;
                        row["科目成績加權平均(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreAT(false, "").Value);
                }

                //  領域 加權總分 (含彈性)
                if (studExamScoreDict.ContainsKey(StudRec.ID))
                {
                    if (studExamScoreDict[StudRec.ID].GetDomainScoreS(true).HasValue)
                        row["領域成績加權總分"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainScoreS(true).Value);

                    if (studExamScoreDict[StudRec.ID].GetDomainScoreSF(true).HasValue)
                        row["領域定期成績加權總分"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainScoreSF(true).Value);
                    if (studExamScoreDict[StudRec.ID].GetDomainScoreS(false).HasValue)
                        row["領域成績加權總分(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainScoreS(false).Value);

                    // 科目 加權總分(不含彈性)
                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreSA(true, "").HasValue)
                        row["科目平時成績加權總分"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreSA(true, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreSF(true, "").HasValue)
                        row["科目定期成績加權總分"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreSF(true, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreST(true, "").HasValue)
                        row["科目成績加權總分"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreST(true, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreSA(false, "").HasValue)
                        row["科目平時成績加權總分(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreSA(false, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreSF(false, "").HasValue)
                        row["科目定期成績加權總分(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreSF(false, "").Value);

                    if (studExamScoreDict[StudRec.ID].GetSubjectScoreST(false, "").HasValue)
                        row["科目成績加權總分(不含彈性)"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetSubjectScoreST(false, "").Value);
                }


                // 領域
                // 算數總分
                // 算數平均
                if (studExamScoreDict.ContainsKey(StudRec.ID))
                {

                    if (studExamScoreDict[StudRec.ID].GetDomainScore_S(true).HasValue)
                        row["領域成績總分"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainScore_S(true).Value);

                    if (studExamScoreDict[StudRec.ID].GetDomainScore_SF(true).HasValue)
                        row["領域定期成績總分"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainScore_SF(true).Value);

                    if (studExamScoreDict[StudRec.ID].GetDomainScore_A(true).HasValue)
                        row["領域成績平均"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainScore_A(true).Value);

                    if (studExamScoreDict[StudRec.ID].GetDomainScore_F(true).HasValue)
                        row["領域定期成績平均"] = doParseTransfer(studExamScoreDict[StudRec.ID].GetDomainScore_F(true).Value);

                }

                // todo 科目 Jean  
                // (算數平均、算術總分) *(科目總成績 、科目定期成績)
                // 
                // 這邊以列舉內的項目 下去組合 保持 用詞一致性
                if (studExamScoreDict.ContainsKey(StudRec.ID))
                {
                    //增加科目相關欄位 by Jean *以上科目相關先註解
                    string _scoreTarget = "科目";
                    string[] _ScoreCaculateWays = new string[] { "平均", "平均(不含彈性)", "總分", "總分(不含彈性)" }; //這邊只增加算術相關成績 所以不從列舉中取得


                    // Type _ScoreCompositions = typeof(EnumScoreComposition);

                    Enum.TryParse(_scoreTarget, out EnumScoreType enumScoreType); // 將"科目"字串轉換為  enum 已符合下面

                    foreach (EnumScoreComposition scoreComposition in Enum.GetValues(_ScoreCompositions)) // 用列舉裡面的 (成績,定期評量,平時評量)
                    {
                        foreach (string scoreCaculateWay in _ScoreCaculateWays) // 用陣列裡面("平均", "平均(不含彈性)", "總分", "總分(不含彈性))
                        {
                            if (scoreCaculateWay.Contains("平均"))
                            {

                              decimal? ScoreArithmeticＭean = studExamScoreDict[StudRec.ID].GetScoreArithmeticＭean( Global.CheckIfContainFlex(scoreCaculateWay) // 確認是否包含彈性
                                                                                                                    , enumScoreType
                                                                                                                    , scoreComposition);

                                if (ScoreArithmeticＭean.HasValue)
                                    row[$"{_scoreTarget}{scoreComposition}{scoreCaculateWay}"] =  ScoreArithmeticＭean;
                            }
                            else if (scoreCaculateWay.Contains("總分"))
                            {

                                decimal? tScoreArithmeticTotal = studExamScoreDict[StudRec.ID].GetScoreArithmeticTotal( // 取得算數平均
                                                                                                                    Global.CheckIfContainFlex(scoreCaculateWay) // 確認是否包含彈性
                                                                                                                    , enumScoreType
                                                                                                                    , scoreComposition);
                                if (tScoreArithmeticTotal.HasValue)
                                    row[$"{_scoreTarget}{scoreComposition}{scoreCaculateWay}"] = doParseTransfer(
                                        studExamScoreDict[StudRec.ID].GetScoreArithmeticTotal( // 取得算數平均
                                           Global.CheckIfContainFlex(scoreCaculateWay) // 確認是否包含彈性
                                           , enumScoreType
                                           , scoreComposition).Value
                                           );
                            }
                        }
                    }
                }





                // 處理總計排名 // 領域成績  //刪除功能變數 因為固定排名的 總計成績是從科目直接計算來得 沒有所謂領域總計成績 Jean Dele
                if (StudentExamRankMatrixDict.ContainsKey(StudRec.ID))
                {
                    foreach (string st in scoreTypeList)
                    {
                        foreach (string rt in rankTypeList)
                        {
                            string key = "定期評量/總計成績" + st + rt;
                            if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(key))
                            {
                                row["科目成績" + st + rt + "名次"] = StudentExamRankMatrixDict[StudRec.ID][key].rank;
                                row["科目成績" + st + rt + "PR值"] = StudentExamRankMatrixDict[StudRec.ID][key].pr;
                                row["科目成績" + st + rt + "百分比"] = StudentExamRankMatrixDict[StudRec.ID][key].percentile;
                                row["科目成績" + st + rt + "母體頂標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_25;
                                row["科目成績" + st + rt + "母體前標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_50;
                                row["科目成績" + st + rt + "母體平均"] = StudentExamRankMatrixDict[StudRec.ID][key].avg;
                                row["科目成績" + st + rt + "母體後標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_50;
                                row["科目成績" + st + rt + "母體底標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_25;
                                row["科目成績" + st + rt + "母體人數"] = StudentExamRankMatrixDict[StudRec.ID][key].matrix_count;

                                if (st.Contains("平均"))
                                {
                                    row["科目成績" + st + rt + "R100_u"] = StudentExamRankMatrixDict[StudRec.ID][key].level_gte100;
                                    row["科目成績" + st + rt + "R90_99"] = StudentExamRankMatrixDict[StudRec.ID][key].level_90;
                                    row["科目成績" + st + rt + "R80_89"] = StudentExamRankMatrixDict[StudRec.ID][key].level_80;
                                    row["科目成績" + st + rt + "R70_79"] = StudentExamRankMatrixDict[StudRec.ID][key].level_70;
                                    row["科目成績" + st + rt + "R60_69"] = StudentExamRankMatrixDict[StudRec.ID][key].level_60;
                                    row["科目成績" + st + rt + "R50_59"] = StudentExamRankMatrixDict[StudRec.ID][key].level_50;
                                    row["科目成績" + st + rt + "R40_49"] = StudentExamRankMatrixDict[StudRec.ID][key].level_40;
                                    row["科目成績" + st + rt + "R30_39"] = StudentExamRankMatrixDict[StudRec.ID][key].level_30;
                                    row["科目成績" + st + rt + "R20_29"] = StudentExamRankMatrixDict[StudRec.ID][key].level_20;
                                    row["科目成績" + st + rt + "R10_19"] = StudentExamRankMatrixDict[StudRec.ID][key].level_10;
                                    row["科目成績" + st + rt + "R0_9"] = StudentExamRankMatrixDict[StudRec.ID][key].level_lt10;
                                }
                            }

                            key = "定期評量_定期/總計成績" + st + rt;
                            if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(key))
                            {
                                row["科目定期成績" + st + rt + "名次"] = StudentExamRankMatrixDict[StudRec.ID][key].rank;
                                row["科目定期成績" + st + rt + "PR值"] = StudentExamRankMatrixDict[StudRec.ID][key].pr;
                                row["科目定期成績" + st + rt + "百分比"] = StudentExamRankMatrixDict[StudRec.ID][key].percentile;
                                row["科目定期成績" + st + rt + "母體頂標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_25;
                                row["科目定期成績" + st + rt + "母體前標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_50;
                                row["科目定期成績" + st + rt + "母體平均"] = StudentExamRankMatrixDict[StudRec.ID][key].avg;
                                row["科目定期成績" + st + rt + "母體後標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_50;
                                row["科目定期成績" + st + rt + "母體底標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_25;
                                row["科目定期成績" + st + rt + "母體人數"] = StudentExamRankMatrixDict[StudRec.ID][key].matrix_count;

                                if (rt.Contains("平均"))
                                {
                                    row["科目定期成績" + st + rt + "R100_u"] = StudentExamRankMatrixDict[StudRec.ID][key].level_gte100;
                                    row["科目定期成績" + st + rt + "R90_99"] = StudentExamRankMatrixDict[StudRec.ID][key].level_90;
                                    row["科目定期成績" + st + rt + "R80_89"] = StudentExamRankMatrixDict[StudRec.ID][key].level_80;
                                    row["科目定期成績" + st + rt + "R70_79"] = StudentExamRankMatrixDict[StudRec.ID][key].level_70;
                                    row["科目定期成績" + st + rt + "R60_69"] = StudentExamRankMatrixDict[StudRec.ID][key].level_60;
                                    row["科目定期成績" + st + rt + "R50_59"] = StudentExamRankMatrixDict[StudRec.ID][key].level_50;
                                    row["科目定期成績" + st + rt + "R40_49"] = StudentExamRankMatrixDict[StudRec.ID][key].level_40;
                                    row["科目定期成績" + st + rt + "R30_39"] = StudentExamRankMatrixDict[StudRec.ID][key].level_30;
                                    row["科目定期成績" + st + rt + "R20_29"] = StudentExamRankMatrixDict[StudRec.ID][key].level_20;
                                    row["科目定期成績" + st + rt + "R10_19"] = StudentExamRankMatrixDict[StudRec.ID][key].level_10;
                                    row["科目定期成績" + st + rt + "R0_9"] = StudentExamRankMatrixDict[StudRec.ID][key].level_lt10;
                                }
                            }
                        }
                    }
                }


                // 處理科目算術平均   
                if (StudentExamRankMatrixDict.ContainsKey(StudRec.ID))
                {
                    //給科目 算術平均 及 算術總分用
                    //string[] itemTypes= new string[] { "定期評量_定期/總計成績", "定期評量/總計成績" };


                    // 2020/03/23 定期評量/總計成績 =>(定期加平時) Jean
                    string scoreType = (EnumScoreType.科目).ToString();  //科目 領域 參考科目

                    foreach (string itemName in itemNames) //平均 總分 => 基本上這兩個 就是從科目來的
                    {
                        foreach (string rankType in rankTypeList) // 年排名 班排名 類別1排名 類別2排名 
                        {
                            foreach (string itemType in itemTypesMapping.Keys)
                            {
                                string key = itemType + itemName + rankType;

                                if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(key))
                                {
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}名次"] = StudentExamRankMatrixDict[StudRec.ID][key].rank;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}PR值"] = StudentExamRankMatrixDict[StudRec.ID][key].pr;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}百分比"] = StudentExamRankMatrixDict[StudRec.ID][key].percentile;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}母體頂標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_25;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}母體前標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_50;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}母體平均"] = StudentExamRankMatrixDict[StudRec.ID][key].avg;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}母體後標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_50;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}母體底標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_25;
                                    row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}母體人數"] = StudentExamRankMatrixDict[StudRec.ID][key].matrix_count;


                                    if (itemName.Contains("平均"))
                                    {
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R100_u"] = StudentExamRankMatrixDict[StudRec.ID][key].level_gte100;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R90_99"] = StudentExamRankMatrixDict[StudRec.ID][key].level_90;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R80_89"] = StudentExamRankMatrixDict[StudRec.ID][key].level_80;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R70_79"] = StudentExamRankMatrixDict[StudRec.ID][key].level_70;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R60_69"] = StudentExamRankMatrixDict[StudRec.ID][key].level_60;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R50_59"] = StudentExamRankMatrixDict[StudRec.ID][key].level_50;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R40_49"] = StudentExamRankMatrixDict[StudRec.ID][key].level_40;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R30_39"] = StudentExamRankMatrixDict[StudRec.ID][key].level_30;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R20_29"] = StudentExamRankMatrixDict[StudRec.ID][key].level_20;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R10_19"] = StudentExamRankMatrixDict[StudRec.ID][key].level_10;
                                        row[$"{scoreType}{itemTypesMapping[itemType]}{itemName}{rankType}R0_9"] = StudentExamRankMatrixDict[StudRec.ID][key].level_lt10;
                                    }
                                }

                            }

                            //key = "定期評量_定期/總計成績" + itemName + rankType;
                            //if (StudentExamRankMatrixDict[StudRec.ID].ContainsKey(key))
                            //{
                            //    row["領域定期成績" + itemName + rankType + "名次"] = StudentExamRankMatrixDict[StudRec.ID][key].rank;
                            //    row["領域定期成績" + itemName + rankType + "PR值"] = StudentExamRankMatrixDict[StudRec.ID][key].pr;
                            //    row["領域定期成績" + itemName + rankType + "百分比"] = StudentExamRankMatrixDict[StudRec.ID][key].percentile;
                            //    row["領域定期成績" + itemName + rankType + "母體頂標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_25;
                            //    row["領域定期成績" + itemName + rankType + "母體前標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_top_50;
                            //    row["領域定期成績" + itemName + rankType + "母體平均"] = StudentExamRankMatrixDict[StudRec.ID][key].avg;
                            //    row["領域定期成績" + itemName + rankType + "母體後標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_50;
                            //    row["領域定期成績" + itemName + rankType + "母體底標"] = StudentExamRankMatrixDict[StudRec.ID][key].avg_bottom_25;

                            //    if (rankType.Contains("平均"))
                            //    {
                            //        row["領域定期成績" + itemName + rankType + "R100_u"] = StudentExamRankMatrixDict[StudRec.ID][key].level_gte100;
                            //        row["領域定期成績" + itemName + rankType + "R90_99"] = StudentExamRankMatrixDict[StudRec.ID][key].level_90;
                            //        row["領域定期成績" + itemName + rankType + "R80_89"] = StudentExamRankMatrixDict[StudRec.ID][key].level_80;
                            //        row["領域定期成績" + itemName + rankType + "R70_79"] = StudentExamRankMatrixDict[StudRec.ID][key].level_70;
                            //        row["領域定期成績" + itemName + rankType + "R60_69"] = StudentExamRankMatrixDict[StudRec.ID][key].level_60;
                            //        row["領域定期成績" + itemName + rankType + "R50_59"] = StudentExamRankMatrixDict[StudRec.ID][key].level_50;
                            //        row["領域定期成績" + itemName + rankType + "R40_49"] = StudentExamRankMatrixDict[StudRec.ID][key].level_40;
                            //        row["領域定期成績" + itemName + rankType + "R30_39"] = StudentExamRankMatrixDict[StudRec.ID][key].level_30;
                            //        row["領域定期成績" + itemName + rankType + "R20_29"] = StudentExamRankMatrixDict[StudRec.ID][key].level_20;
                            //        row["領域定期成績" + itemName + rankType + "R10_19"] = StudentExamRankMatrixDict[StudRec.ID][key].level_10;
                            //        row["領域定期成績" + itemName + rankType + "R0_9"] = StudentExamRankMatrixDict[StudRec.ID][key].level_lt10;
                            //    }
                            //}
                        }
                    }
                }


                // 處理參考總計成績排名
                if (StudentRefExamRankMatrixDict.ContainsKey(StudRec.ID))
                {
                    foreach (string st in scoreTypeList) //加權平均 加權總分
                    {
                        foreach (string rt in rankTypeList)
                        {
                            string key = "定期評量/總計成績" + st + rt;
                            if (StudentRefExamRankMatrixDict[StudRec.ID].ContainsKey(key))
                            {
                                row["參考科目成績" + st + rt + "名次"] = StudentRefExamRankMatrixDict[StudRec.ID][key].rank;
                                row["參考科目成績" + st + rt + "PR值"] = StudentRefExamRankMatrixDict[StudRec.ID][key].pr;
                                row["參考科目成績" + st + rt + "百分比"] = StudentRefExamRankMatrixDict[StudRec.ID][key].percentile;
                                row["參考科目成績" + st + rt + "母體頂標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_top_25;
                                row["參考科目成績" + st + rt + "母體前標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_top_50;
                                row["參考科目成績" + st + rt + "母體平均"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg;
                                row["參考科目成績" + st + rt + "母體後標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_bottom_50;
                                row["參考科目成績" + st + rt + "母體底標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_bottom_25;
                                row["參考科目成績" + st + rt + "母體人數"] = StudentRefExamRankMatrixDict[StudRec.ID][key].matrix_count;

                                if (st.Contains("平均"))
                                {
                                    row["參考科目成績" + st + rt + "R100_u"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_gte100;
                                    row["參考科目成績" + st + rt + "R90_99"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_90;
                                    row["參考科目成績" + st + rt + "R80_89"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_80;
                                    row["參考科目成績" + st + rt + "R70_79"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_70;
                                    row["參考科目成績" + st + rt + "R60_69"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_60;
                                    row["參考科目成績" + st + rt + "R50_59"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_50;
                                    row["參考科目成績" + st + rt + "R40_49"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_40;
                                    row["參考科目成績" + st + rt + "R30_39"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_30;
                                    row["參考科目成績" + st + rt + "R20_29"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_20;
                                    row["參考科目成績" + st + rt + "R10_19"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_10;
                                    row["參考科目成績" + st + rt + "R0_9"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_lt10;
                                }
                            }

                            key = "定期評量_定期/總計成績" + st + rt;
                            if (StudentRefExamRankMatrixDict[StudRec.ID].ContainsKey(key))
                            {
                                row["參考科目定期成績" + st + rt + "名次"] = StudentRefExamRankMatrixDict[StudRec.ID][key].rank;
                                row["參考科目定期成績" + st + rt + "PR值"] = StudentRefExamRankMatrixDict[StudRec.ID][key].pr;
                                row["參考科目定期成績" + st + rt + "百分比"] = StudentRefExamRankMatrixDict[StudRec.ID][key].percentile;
                                row["參考科目定期成績" + st + rt + "母體頂標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_top_25;
                                row["參考科目定期成績" + st + rt + "母體前標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_top_50;
                                row["參考科目定期成績" + st + rt + "母體平均"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg;
                                row["參考科目定期成績" + st + rt + "母體後標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_bottom_50;
                                row["參考科目定期成績" + st + rt + "母體底標"] = StudentRefExamRankMatrixDict[StudRec.ID][key].avg_bottom_25;
                                row["參考科目定期成績" + st + rt + "母體人數"] = StudentRefExamRankMatrixDict[StudRec.ID][key].matrix_count;

                                if (st.Contains("平均"))
                                {
                                    row["參考科目定期成績" + st + rt + "R100_u"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_gte100;
                                    row["參考科目定期成績" + st + rt + "R90_99"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_90;
                                    row["參考科目定期成績" + st + rt + "R80_89"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_80;
                                    row["參考科目定期成績" + st + rt + "R70_79"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_70;
                                    row["參考科目定期成績" + st + rt + "R60_69"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_60;
                                    row["參考科目定期成績" + st + rt + "R50_59"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_50;
                                    row["參考科目定期成績" + st + rt + "R40_49"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_40;
                                    row["參考科目定期成績" + st + rt + "R30_39"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_30;
                                    row["參考科目定期成績" + st + rt + "R20_29"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_20;
                                    row["參考科目定期成績" + st + rt + "R10_19"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_10;
                                    row["參考科目定期成績" + st + rt + "R0_9"] = StudentRefExamRankMatrixDict[StudRec.ID][key].level_lt10;
                                }
                            }
                        }
                    }
                }

                //新增參考科目  






                row["服務學習時數"] = "";
                if (ServiceLearningDict.ContainsKey(StudRec.ID))
                    row["服務學習時數"] = ServiceLearningDict[StudRec.ID];

                row["校長"] = ChancellorChineseName;
                row["教務主任"] = EduDirectorName;
                row["區間開始日期"] = _BeginDate.ToShortDateString();
                row["區間結束日期"] = _EndDate.ToShortDateString();
                row["成績校正日期"] = _ScoreEditDate;

                dt.Rows.Add(row);
                dtAtt.Rows.Add(rowT);

                //dt.TableName = "test";
                //dt.WriteXml(Application.StartupPath + "\\testData.xml");


                // 處理固定欄位對應
                Document doc1 = new Document();
                doc1.Sections.Clear();

                // 處理動態處理(缺曠)
                Document docAtt = new Document();
                docAtt.Sections.Clear();
                docAtt.Sections.Add(docAtt.ImportNode(docTemplate.Sections[0], true));

                //_builder = new DocumentBuilder(docAtt);
                //                docAtt.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
                docAtt.MailMerge.FieldMergingCallback = this;
                try
                {
                    docAtt.MailMerge.Execute(dtAtt);
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine(ex.Message);
                }

                doc1.Sections.Add(doc1.ImportNode(docAtt.Sections[0], true));
                doc1.MailMerge.Execute(dt);

                //foreach (DataColumn cols in dt.Columns)
                //{
                //    if (doc1.MailMerge.GetFieldNames().Contains(cols.ColumnName)) 
                //    {
                    
                //    }
                //} 
               
                doc1.MailMerge.RemoveEmptyParagraphs = true;
               doc1.MailMerge.DeleteFields();
                docList.Add(doc1);

            }

            _bgWorkReport.ReportProgress(80);
            //// debug 用           
            //string ssStr = Application.StartupPath + "\\dt_debug.xml";
            //dt.WriteXml(ssStr);


            #endregion

            #region Word 合併列印

            doc.Sections.Clear();
            foreach (Document doc1 in docList)
                doc.Sections.Add(doc.ImportNode(doc1.Sections[0], true));

            string reportNameW = "新竹評量成績單";
            pathW = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", "");
            if (!Directory.Exists(pathW))
                Directory.CreateDirectory(pathW);
            pathW = Path.Combine(pathW, reportNameW + ".doc");

            if (File.Exists(pathW))
            {
                int i = 1;
                while (true)
                {
                    string newPathW = Path.GetDirectoryName(pathW) + "\\" + Path.GetFileNameWithoutExtension(pathW) + (i++) + Path.GetExtension(pathW);
                    if (!File.Exists(newPathW))
                    {
                        pathW = newPathW;
                        break;
                    }
                }
            }

            try
            {
                //2019/4/19 俊緯更新 完成[190417-01][03] 電子報表及推播功能，聽恩正建議將Save功能移到Complete執行
                MemoryStream memoryStream = new MemoryStream();
                doc.Save(memoryStream, SaveFormat.Docx);
                e.Result = new object[] { memoryStream, reportNameW };

            }
            catch (Exception exow)
            {
                throw exow;
            }
            doc = null;
            docList.Clear();

            GC.Collect();
            #endregion
            _bgWorkReport.ReportProgress(100);
        }

        //void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        //{

        //}


        // 載入學生所屬學年度學習的試別，科目，並排序
        private void LoadExamSubject()
        {
            // 取得該學年度學期所有學生的試別修課科目
            _SelSchoolYear = _SelSemester = 0;
            int ss, sc;
            if (int.TryParse(cboSchoolYear.Text, out ss))
                _SelSchoolYear = ss;

            if (int.TryParse(cboSemester.Text, out sc))
                _SelSemester = sc;

            _ExamSubjectFull = Utility.GetExamSubjecList(_StudentIDList, _SelSchoolYear, _SelSemester);

            foreach (var list in _ExamSubjectFull.Values)
            {
                #region 排序
                //list.Sort(new StringComparer("國文"
                //                , "英文"
                //                , "數學"
                //                , "理化"
                //                , "生物"
                //                , "社會"
                //                , "物理"
                //                , "化學"
                //                , "歷史"
                //                , "地理"
                //                , "公民"));


                list.Sort(new StringComparer(Utility.GetSubjectOrder().ToArray()));



                #endregion
            }
        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            DisSelect();
            _SelSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _SelSemester = int.Parse(K12.Data.School.DefaultSemester);
            _SelRefSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _SelRefSemester = int.Parse(K12.Data.School.DefaultSemester);
            _ScoreMappingConfig.LoadData();
            //Jean todo 
            _UserDefineFields = Global.GetUserDefineFields(); // 載入自訂欄位

            //加入成績顯示小數位數設定
            for (int i = 0; i <= 5; i++)
                cboParseNumber.Items.Add(i);


            bkw.RunWorkerAsync();
        }

        private void LoadSubject()
        {
            lvSubject.Items.Clear();
            string ExamID = "";
            foreach (ExamRecord ex in _exams)
            {
                if (ex.Name == cboExam.Text)
                {
                    ExamID = ex.ID;
                    break;
                }
            }

            if (_ExamSubjectFull.ContainsKey(ExamID))
            {
                foreach (string subjName in _ExamSubjectFull[ExamID])
                    lvSubject.Items.Add(subjName);
            }
        }


        private void lnkCopyConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = _Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;
                conf.ExamRecord = _Configure.ExamRecord;
                conf.RefExamRecord = _Configure.RefExamRecord;
                conf.PrintSubjectList.AddRange(_Configure.PrintSubjectList);
                conf.SchoolYear = _Configure.SchoolYear;
                conf.Semester = _Configure.Semester;
                conf.RefSchoolYear = _Configure.RefSchoolYear;
                conf.RefSemester = _Configure.RefSemester;
                conf.SubjectLimit = _Configure.SubjectLimit;
                conf.Template = _Configure.Template;
                conf.BeginDate = _Configure.BeginDate;
                conf.EndDate = _Configure.EndDate;
                conf.ParseNumber = _Configure.ParseNumber;
                conf.ScoreEditDate = _Configure.ScoreEditDate;
                if (conf.PrintAttendanceList == null)
                    conf.PrintAttendanceList = new List<string>();
                conf.PrintAttendanceList.AddRange(_Configure.PrintAttendanceList);
                conf.Encode();
                conf.Save();
                _ConfigureList.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }

        public Configure _Configure { get; private set; }

        private void lnkDelConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;

            // 檢查是否是預設設定檔名稱，如果是無法刪除
            if (Global.DefaultConfigNameList().Contains(_Configure.Name))
            {
                FISCA.Presentation.Controls.MsgBox.Show("系統預設設定檔案無法刪除");
                return;
            }

            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _ConfigureList.Remove(_Configure);
                if (_Configure.UID != "")
                {
                    _Configure.Deleted = true;
                    _Configure.Save();
                }
                var conf = _Configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 停用可選功能
        private void DisSelect()
        {

            cboConfigure.Enabled = false;
            cboExam.Enabled = false;
            cboRefExam.Enabled = false;
            cboSchoolYear.Enabled = false;
            cboSemester.Enabled = false;
            cboRefSchoolYear.Enabled = false;
            cboRefSemester.Enabled = false;
            btnSaveConfig.Enabled = false;
            btnPrint.Enabled = false;
        }

        // 啟用可選功能
        private void EnbSelect()
        {
            cboConfigure.Enabled = true;
            cboExam.Enabled = true;
            cboRefExam.Enabled = true;
            cboSchoolYear.Enabled = true;
            cboSemester.Enabled = true;
            cboRefSchoolYear.Enabled = true;
            cboRefSemester.Enabled = true;
            btnSaveConfig.Enabled = true;
            btnPrint.Enabled = true;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrintClick = true;
            if (dtBegin.IsEmpty || dtEnd.IsEmpty)
            {
                FISCA.Presentation.Controls.MsgBox.Show("日期區間必須輸入!");
                return;
            }

            if (dtBegin.Value > dtEnd.Value)
            {
                FISCA.Presentation.Controls.MsgBox.Show("開始日期必須小於或等於結束日期!!");
                return;
            }

            int sc, ss, rsc, rss;
            if (int.TryParse(cboSchoolYear.Text, out sc))
            {
                _SelSchoolYear = sc;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學年度必填!");
                return;
            }

            if (int.TryParse(cboSemester.Text, out ss))
            {
                _SelSemester = ss;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學期必填!");
                return;
            }

            if (string.IsNullOrEmpty(cboExam.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇試別!");
                return;
            }
            else
            {

                bool isEr = true;
                foreach (ExamRecord ex in _exams)
                    if (ex.Name == cboExam.Text)
                    {
                        _SelExamID = ex.ID;
                        _SelExamName = ex.Name;
                        isEr = false;
                        break;
                    }

                if (isEr)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("試別錯誤，請重新選擇!");
                    return;
                }
            }

            if (!string.IsNullOrEmpty(cboRefExam.Text))
            {
                this.HasReferenceExam = true;
                bool isEr = true;
                foreach (ExamRecord ex in _exams)
                    if (ex.Name == cboRefExam.Text)
                    {
                        _SelRefExamID = ex.ID;
                        _SelRefExamName = ex.Name;
                        isEr = false;
                        break;
                    }

                if (int.TryParse(cboRefSchoolYear.Text, out rsc))
                {
                    _SelRefSchoolYear = rsc;
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("需填參考學年度!");
                    return;
                }

                if (int.TryParse(cboRefSemester.Text, out rss))
                {
                    _SelRefSemester = rss;
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("需填參考學期!");
                    return;
                }
                if (isEr)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("參考試別錯誤，請重新選擇!");
                    return;
                }
            }
            //_SelNotRankedFilter = cboNotRankedFilter.Text;
            _SelSubjNameList.Clear();

            Global.parseNumebr = parseNumber;

            SaveTemplate(null, null);

            // 使用者勾選科目
            foreach (string name in _Configure.PrintSubjectList)
                _SelSubjNameList.Add(name);

            // 缺曠
            foreach (string name in _Configure.PrintAttendanceList)
                _SelAttendanceList.Add(name);


            _BeginDate = dtBegin.Value;
            _EndDate = dtEnd.Value;

            if (dtScoreEdit.IsEmpty)
                _ScoreEditDate = "";
            else
                _ScoreEditDate = dtScoreEdit.Value.ToShortDateString();

            btnSaveConfig.Enabled = false;
            btnSaveConfig.Enabled = false;

            SetDomainList();

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
            // 執行報表
            _bgWorkReport.RunWorkerAsync();
        }

        // 儲存樣板
        private void SaveTemplate(object sender, EventArgs e)
        {
            DisSelect();
            if (_Configure == null) return;
            _Configure.SchoolYear = cboSchoolYear.Text;
            _Configure.Semester = cboSemester.Text;
            _Configure.RefSchoolYear = cboRefSchoolYear.Text;
            _Configure.RefSemester = cboRefSemester.Text;
            _Configure.SelSetConfigName = cboConfigure.Text;
            _Configure.NotRankedTagNameFilter = _SelNotRankedFilter;
            //_Configure.ParseNumber = cboParseNumber.Text;
            _Configure.ParseNumber = Int32.Parse(cboParseNumber.Text);
            parseNumber = _Configure.ParseNumber;
            foreach (ExamRecord exm in _exams)
            {
                if (exm.Name == cboRefExam.Text)
                {
                    _Configure.RefExamRecord = exm;
                    //break;
                }

                if (exm.Name == cboExam.Text)
                {
                    _Configure.ExamRecord = exm;
                    //break;
                }
            }
 
            // 科目
            foreach (ListViewItem item in lvSubject.Items)
            {
                if (item.Checked)
                {
                    if (!_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Add(item.Text);
                }
                else
                {
                    if (_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Remove(item.Text);
                }
            }

            if (_Configure.PrintAttendanceList == null)
                _Configure.PrintAttendanceList = new List<string>();

            _Configure.PrintAttendanceList.Clear();
            // 儲存缺曠選項
            foreach (DataGridViewRow drv in dgAttendanceData.Rows)
            {
                foreach (DataGridViewCell cell in drv.Cells)
                {
                    bool bl;
                    if (bool.TryParse(cell.Value.ToString(), out bl))
                    {
                        if (bl)
                            _Configure.PrintAttendanceList.Add(cell.Tag.ToString());
                    }
                }
            }


            foreach (ListViewItem item in lvSubject.Items)
            {
                if (item.Checked)
                {
                    if (!_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Add(item.Text);
                }
                else
                {
                    if (_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Remove(item.Text);
                }
            }
    
            // 儲存開始與結束日期
            _Configure.BeginDate = dtBegin.Value.ToShortDateString();
            _Configure.EndDate = dtEnd.Value.ToShortDateString();
            if (dtScoreEdit.IsEmpty)
                _Configure.ScoreEditDate = "";
            else
                _Configure.ScoreEditDate = dtScoreEdit.Value.ToShortDateString();

            _Configure.Encode();
            _Configure.Save();

            #region 樣板設定檔記錄用

            // 記錄使用這選的專案            
            List<DAO.UDT_ScoreConfig> uList = new List<DAO.UDT_ScoreConfig>();
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    conf.Name = cboConfigure.Text;
                    uList.Add(conf);
                    break;
                }

            if (uList.Count > 0)
            {
                DAO.UDTTransfer.UpdateConfigData(uList);
            }
            else
            {
                // 新增
                List<DAO.UDT_ScoreConfig> iList = new List<DAO.UDT_ScoreConfig>();
                DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                conf.Name = cboConfigure.Text;
                conf.ProjectName = Global._ProjectName;
                conf.Type = Global._UserConfTypeName;
                conf.UDTTableName = Global._UDTTableName;
                iList.Add(conf);
                DAO.UDTTransfer.InsertConfigData(iList);
            }
            #endregion
            if(btnPrintClick == false)
            {
                MessageBox.Show("已完成樣板儲存", "儲存樣板", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if(cboSchoolYear.Text != cboRefSchoolYear.Text || cboSemester.Text != cboRefSemester.Text)
                {
                    MessageBox.Show("跨學年度學期之參考試別只會呈排名資料，不會呈現評量成績。", "列印提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnPrintClick = false;
                }
                    btnPrintClick = false;
            }
            

            EnbSelect();
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisSelect();
            LoadExamSubject();
            LoadSubject();
            EnbSelect();
        }

        private void cboSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisSelect();
            LoadExamSubject();
            LoadSubject();
            EnbSelect();
        }

        private void cboExam_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSubject();
        }

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = btnPrint.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    _Configure = new Configure();
                    _Configure.Name = dialog.ConfigName;
                    _Configure.Template = dialog.Template;
                    _Configure.SubjectLimit = dialog.SubjectLimit;
                    _Configure.SchoolYear = _DefalutSchoolYear;
                    _Configure.Semester = _DefaultSemester;
                    _Configure.RefSchoolYear = _DefalutSchoolYear;
                    _Configure.RefSemester = _DefaultSemester;
                    if (_Configure.PrintAttendanceList == null)
                        _Configure.PrintAttendanceList = new List<string>();
                    if (_Configure.PrintSubjectList == null)
                        _Configure.PrintSubjectList = new List<string>();

                    if (cboExam.Items.Count > 0)
                    {
                        string exName = cboExam.Items[0].ToString();
                        foreach (ExamRecord rec in _exams)
                        {
                            if (exName == rec.Name)
                            {
                                _Configure.ExamRecord = rec;
                                break;
                            }
                        }
                    }
                    if (cboRefExam.Items.Count > 0)
                    {
                        string exName = cboRefExam.Items[0].ToString();
                        foreach (ExamRecord rec in _exams)
                        {
                            if (exName == rec.Name)
                            {
                                _Configure.RefExamRecord = rec;
                                break;
                            }
                        }

                    }
                    _Configure.ParseNumber = parseNumber;
                    _ConfigureList.Add(_Configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, _Configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;
                    _Configure.Encode();
                    _Configure.Save();
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnSaveConfig.Enabled = btnPrint.Enabled = true;
                    _Configure = _ConfigureList[cboConfigure.SelectedIndex];
                    if (_Configure.Template == null)
                        _Configure.Decode();
                    if (!cboSchoolYear.Items.Contains(_Configure.SchoolYear))
                        cboSchoolYear.Items.Add(_Configure.SchoolYear);
                    cboSchoolYear.Text = _Configure.SchoolYear;
                    cboSemester.Text = _Configure.Semester;
                    cboRefSchoolYear.Text = _Configure.RefSchoolYear;
                    cboRefSemester.Text = _Configure.RefSemester;
                    if (_Configure.ExamRecord != null)
                    {
                        int idx = 0;
                        foreach (string sitm in cboExam.Items)
                        {
                            if (sitm == _Configure.ExamRecord.Name)
                            {
                                cboExam.SelectedIndex = idx;
                                break;
                            }
                            idx++;
                        }
                    }
                    else
                    {
                        cboExam.Text = "";
                    }
                    if (_Configure.RefExamRecord != null)
                    {
                        if (cboRefExam.Text == "") // 初始化一下,因為如果cboRefExam.Text 為空字串時(別的樣板可能會這樣帶過來), .SelectedIndex 時.Text也都會是空字串
                        {
                            cboRefExam.SelectedIndex = 0;
                        }
                            int idx = 0;
                        foreach (string sitm in cboRefExam.Items)
                        {
                            if (sitm == _Configure.RefExamRecord.Name)
                            {
                                cboRefExam.SelectedIndex = idx;
                                break;
                            }
                            idx++;
                        }
                    }
                    else
                    {
                        cboRefExam.Text = "";
                    }
                    //Global.parseNumebr = parseNumber = _Configure.ParseNumber;
                    //Global.parseNumebr = parseNumber = int.Parse(_Configure.ParseNumber);
                    parseNumber = _Configure.ParseNumber;
                    cboParseNumber.Text = _Configure.ParseNumber.ToString();
                    // 解析科目
                    foreach (ListViewItem lvi in lvSubject.Items)
                    {
                        if (_Configure.PrintSubjectList.Contains(lvi.Text))
                        {
                            lvi.Checked = true;
                        }
                    }

                    // 解析缺曠
                    foreach (DataGridViewRow drv in dgAttendanceData.Rows)
                    {
                        foreach (DataGridViewCell cell in drv.Cells)
                        {
                            if (cell.Tag == null)
                                continue;

                            string key = cell.Tag.ToString();
                            cell.Value = false;
                            if (_Configure.PrintAttendanceList.Contains(key))
                            {
                                cell.Value = true;
                            }
                        }
                    }


                    // 開始與結束日期
                    DateTime dtb, dte, dtee;
                    if (DateTime.TryParse(_Configure.BeginDate, out dtb))
                        dtBegin.Value = dtb;
                    else
                        dtBegin.Value = DateTime.Now;

                    if (DateTime.TryParse(_Configure.EndDate, out dte))
                        dtEnd.Value = dte;
                    else
                        dtEnd.Value = DateTime.Now;

                    // 成績校正日期
                    if (DateTime.TryParse(_Configure.ScoreEditDate, out dtee))
                        dtScoreEdit.Value = dtee;
                    else
                        dtScoreEdit.IsEmpty = true;

                }
                else
                {
                    _Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;
                    cboExam.SelectedIndex = -1;

                    // 開始與結束日期沒有預設值時給當天
                    dtBegin.Value = dtEnd.Value = DateTime.Now;
                }
            }
        }

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "新竹評量成績單樣板(" + _Configure.Name + ").doc";

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                _Configure.Template.Save(stream, Aspose.Words.SaveFormat.Doc);

                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.新竹評量成績單樣板_固定排名_科目_領域__doc1, 0, Properties.Resources.新竹評量成績單樣板_固定排名_科目_領域__doc1.Length);
                        stream.Flush();
                        stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkViewTemplate.Enabled = true;
            #endregion
        }

        private void lnkChangeTemplate_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            lnkChangeTemplate.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    _Configure.Template = new Aspose.Words.Document(dialog.FileName);
                    List<string> fields = new List<string>(_Configure.Template.MailMerge.GetFieldNames());
                    _Configure.SubjectLimit = 0;
                    while (fields.Contains("科目名稱" + (_Configure.SubjectLimit + 1)))
                    {
                        _Configure.SubjectLimit++;
                    }

                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
            lnkChangeTemplate.Enabled = true;
        }

        private void lnkViewMapColumns_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {

            int sc, ss;
            if (int.TryParse(cboSchoolYear.Text, out sc))
            {
                _SelSchoolYear = sc;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學年度必填!");
                return;
            }

            if (int.TryParse(cboSemester.Text, out ss))
            {
                _SelSemester = ss;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學期必填!");
                return;
            }

            if (string.IsNullOrEmpty(cboExam.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇試別!");
                return;
            }
            else
            {
                bool isEr = true;
                foreach (ExamRecord ex in _exams)
                    if (ex.Name == cboExam.Text)
                    {
                        _SelExamID = ex.ID;
                        _SelExamName = ex.Name;
                        isEr = false;
                        break;
                    }

                if (isEr)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("試別錯誤，請重新選擇!");
                    return;
                }
            }

            if (!string.IsNullOrEmpty(cboRefExam.Text))
            {
                bool isEr = true;
                foreach (ExamRecord ex in _exams)
                    if (ex.Name == cboRefExam.Text)
                    {
                        _SelRefExamID = ex.ID;
                        _SelRefExamName = ex.Name;
                        isEr = false;
                        break;
                    }

                if (isEr)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("參考試別錯誤，請重新選擇!");
                    return;
                }
            }
            lnkViewMapColumns.Enabled = false;

            SetDomainList();

            Global.ExportMappingFieldWord(); // 產生合併欄位 
            lnkViewMapColumns.Enabled = true;
        }

        private void SetDomainList()
        {
            Global._SelSchoolYear = _SelSchoolYear;
            Global._SelSemester = _SelSemester;
            Global._SelStudentIDList = _StudentIDList;
            Global._SelExamID = _SelExamID;
            Global._SelRefsExamID = _SelRefExamID;
            Global.SetDomainList();
        }

        private void chkSubjSelAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in lvSubject.Items)
            {
                lvi.Checked = chkSubjSelAll.Checked;
            }
        }

        private void chkAttendSelAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow drv in dgAttendanceData.Rows)
            {
                foreach (DataGridViewCell cell in drv.Cells)
                {
                    if (cell.ColumnIndex != 0)
                        cell.Value = chkAttendSelAll.Checked;
                }
            }
        }




        void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
        {
            if (e.FieldName == "缺曠紀錄")
            {
                var _builder = new DocumentBuilder(e.Document);
                if (_builder.MoveToMergeField(e.FieldName))
                {
                    string sid = e.FieldValue.ToString();

                    Dictionary<string, int> dataDict = new Dictionary<string, int>();
                    List<string> colNameList = new List<string>();
                    if (_AttendanceDict.ContainsKey(sid))
                        dataDict = _AttendanceDict[sid];
                    //dataDict.Keys

                    foreach (string name in _SelAttendanceList)
                        colNameList.Add(name.Replace("_", ""));

                    //colNameList.Sort();
                    int colCount = colNameList.Count;

                    if (colCount > 0)
                    {
                        Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                        cell.CellFormat.LeftPadding = 0;
                        cell.CellFormat.RightPadding = 0;
                        double width = cell.CellFormat.Width;
                        int columnCount = colCount;
                        double miniUnitWitdh = width / (double)columnCount;

                        Table table = _builder.StartTable();

                        //(table.ParentNode.ParentNode as Row).RowFormat.LeftIndent = 0;
                        double p = _builder.RowFormat.LeftIndent;
                        _builder.RowFormat.HeightRule = HeightRule.Exactly;
                        _builder.RowFormat.Height = 18.0;
                        _builder.RowFormat.LeftIndent = 0;

                        // 缺曠名稱
                        foreach (string name in colNameList)
                        {
                            Cell c1 = _builder.InsertCell();
                            c1.CellFormat.Width = miniUnitWitdh;
                            c1.CellFormat.WrapText = true;
                            _builder.Write(name);
                        }
                        _builder.EndRow();

                        // 缺曠統計
                        foreach (string name in colNameList)
                        {
                            Cell c1 = _builder.InsertCell();
                            c1.CellFormat.Width = miniUnitWitdh;
                            c1.CellFormat.WrapText = true;
                            if (dataDict.ContainsKey(name))
                                _builder.Write(dataDict[name].ToString());
                            else
                                _builder.Write("0");
                        }
                        _builder.EndRow();

                        _builder.EndTable();

                        //去除表格四邊的線
                        foreach (Cell c in table.FirstRow.Cells)
                            c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                        foreach (Cell c in table.LastRow.Cells)
                            c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                        foreach (Row r in table.Rows)
                        {
                            r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                            r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                        }

                        _builder.RowFormat.LeftIndent = p;
                    }
                }
            }

        }

        private decimal doParseTransfer(decimal score)
        {
            var transferScore = Math.Round(score, parseNumber, MidpointRounding.AwayFromZero);
            return transferScore;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            throw new NotImplementedException();
        }

        private void tabControlPanel1_Click(object sender, EventArgs e)
        {

        }

        private void cboRefExam_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
