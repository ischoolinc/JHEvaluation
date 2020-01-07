using Aspose.Words;
using FISCA.Data;
using FISCA.Presentation.Controls;
using HsinChuSemesterScoreFixed_JH.DAO;
using JHSchool.Data;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using JHSchool.Behavior.BusinessLogic;
using Campus.ePaperCloud;

namespace HsinChuSemesterScoreFixed_JH
{
    public partial class PrintForm : BaseForm
    {
        FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        List<string> _StudentIDList;
        Dictionary<string, DataRow> _StudentRowDict = new Dictionary<string, DataRow>();
        DataTable dt = new DataTable();
        private List<string> typeList = new List<string>();
        List<string> _文字描述 = new List<string>();

        BackgroundWorker _bgWorkReport;
        DocumentBuilder _builder;
        BackgroundWorker _bgWorkerLoadData;

        List<StudentRecord> _Students = new List<StudentRecord>();

        // 錯誤訊息
        List<string> _ErrorList = new List<string>();

        // 領域錯誤訊息
        List<string> _ErrorDomainNameList = new List<string>();

        // 樣板內有科目名稱
        List<string> _TemplateSubjectNameList = new List<string>();

        // 存檔路徑
        string pathW = "";

        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();

        /// <summary>
        /// 日常生活表現名稱對照使用
        /// </summary>
        private static Dictionary<string, string> _DLBehaviorConfigNameDict = new Dictionary<string, string>();

        //假別設定
        Dictionary<string, List<string>> allowAbsentDic = new Dictionary<string, List<string>>();


        /// <summary>
        /// 日常生活表現子項目名稱,呼叫GetDLBehaviorConfigNameDict 一同取得
        /// </summary>
        private static Dictionary<string, List<string>> _DLBehaviorConfigItemNameDict = new Dictionary<string, List<string>>();


        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        private int _SelSchoolYear;
        private int _SelSemester;

        ScoreMappingConfig _ScoreMappingConfig = new ScoreMappingConfig();

        // 紀錄樣板設定
        List<DAO.UDT_ScoreConfig> _UDTConfigList;

        public PrintForm(List<string> StudIDList)
        {
            InitializeComponent();
            _bgWorkerLoadData = new BackgroundWorker();
            _bgWorkReport = new BackgroundWorker();
            _bgWorkerLoadData.DoWork += _bgWorkerLoadData_DoWork;
            _bgWorkerLoadData.RunWorkerCompleted += _bgWorkerLoadData_RunWorkerCompleted;
            _bgWorkerLoadData.WorkerReportsProgress = true;
            _bgWorkerLoadData.ProgressChanged += _bgWorkerLoadData_ProgressChanged;
            _bgWorkReport.DoWork += _bgWorkReport_DoWork;
            _bgWorkReport.RunWorkerCompleted += _bgWorkReport_RunWorkerCompleted;
            _bgWorkReport.WorkerReportsProgress = true;
            _bgWorkReport.ProgressChanged += _bgWorkReport_ProgressChanged;

            _StudentIDList = StudIDList;
            _Students = Student.SelectByIDs(StudIDList);
        }

        private void _bgWorkReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績報表產生中...", e.ProgressPercentage);
        }

        private void _bgWorkReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
                return;
            }
            Document doc = e.Result as Document;

            string reportName = _SelSchoolYear + "學年度第" + _SelSemester + "學期學期成績通知單";
            MemoryStream memoryStream = new MemoryStream();
            doc.Save(memoryStream, SaveFormat.Doc);
            ePaperCloud ePaperCloud = new ePaperCloud();
            ePaperCloud.upload_ePaper(_SelSchoolYear, _SelSemester, reportName, "", memoryStream, ePaperCloud.ViewerType.Student, ePaperCloud.FormatType.Docx);

            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void _bgWorkReport_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorkReport.ReportProgress(1);
            dt.Clear();
            _DLBehaviorConfigNameDict = GetDLBehaviorConfigNameDict();
            List<string> plist = K12.Data.PeriodMapping.SelectAll().Select(x => x.Type).Distinct().ToList();
            List<string> alist = K12.Data.AbsenceMapping.SelectAll().Select(x => x.Name).ToList();


            // 建立合併欄位
            dt.Columns.Add("列印日期");
            dt.Columns.Add("學校名稱");
            dt.Columns.Add("學年度");
            dt.Columns.Add("學期");
            dt.Columns.Add("系統編號");
            dt.Columns.Add("姓名");
            dt.Columns.Add("英文姓名");
            dt.Columns.Add("班級");
            dt.Columns.Add("座號");
            dt.Columns.Add("學號");
            dt.Columns.Add("大功");
            dt.Columns.Add("小功");
            dt.Columns.Add("嘉獎");
            dt.Columns.Add("大過");
            dt.Columns.Add("小過");
            dt.Columns.Add("警告");
            dt.Columns.Add("上課天數");
            dt.Columns.Add("學習領域成績");
            dt.Columns.Add("學習領域原始成績");
            dt.Columns.Add("課程學習成績");
            dt.Columns.Add("課程學習原始成績");
            dt.Columns.Add("班導師");
            dt.Columns.Add("教務主任");
            dt.Columns.Add("校長");
            dt.Columns.Add("服務學習時數");
            dt.Columns.Add("文字描述");

            //科目欄位
            for (int i = 1; i <= Global.SupportSubjectCount; i++)
            {
                dt.Columns.Add("科目名稱" + i);
                dt.Columns.Add("科目領域" + i);
                dt.Columns.Add("科目節數" + i);
                dt.Columns.Add("科目權數" + i);
                dt.Columns.Add("科目成績" + i);
                dt.Columns.Add("科目原始成績" + i);
                dt.Columns.Add("科目補考成績" + i);
                dt.Columns.Add("科目成績等第" + i);
                dt.Columns.Add("科目原始成績等第" + i);
                dt.Columns.Add("科目補考成績等第" + i);
                dt.Columns.Add("科目需補考標示" + i);
                dt.Columns.Add("科目補考成績標示" + i);
                dt.Columns.Add("科目不及格標示" + i);
            }

            //領域欄位
            for (int i = 1; i <= Global.SupportDomainCount; i++)
            {
                dt.Columns.Add("領域名稱" + i);
                dt.Columns.Add("領域節數" + i);
                dt.Columns.Add("領域權數" + i);
                dt.Columns.Add("領域成績" + i);
                dt.Columns.Add("領域原始成績" + i);
                dt.Columns.Add("領域補考成績" + i);
                dt.Columns.Add("領域成績等第" + i);
                dt.Columns.Add("領域原始成績等第" + i);
                dt.Columns.Add("領域補考成績等第" + i);
                dt.Columns.Add("領域需補考標示" + i);
                dt.Columns.Add("領域補考成績標示" + i);
                dt.Columns.Add("領域不及格標示" + i);

            }


            //學期科目及領域成績
            List<JHSemesterScoreRecord> SemesterScoreRecordList = JHSemesterScore.SelectBySchoolYearAndSemester(_StudentIDList, _SelSchoolYear, _SelSemester);

            // 領域名稱
            List<string> DomainNameList = new List<string>();

            foreach (JHSemesterScoreRecord SemsScore in SemesterScoreRecordList)
            {
                foreach (string dn in SemsScore.Domains.Keys)
                {
                    if (!DomainNameList.Contains(dn))
                        DomainNameList.Add(dn);
                }
            }
            DomainNameList.Sort(new StringComparer("語文", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"));
            DomainNameList.Add("彈性課程");

            _bgWorkReport.ReportProgress(20);

            // 指定領域合併
            foreach (string dName in DomainNameList)
            {
                dt.Columns.Add(dName + "領域");
                dt.Columns.Add(dName + "節數");
                dt.Columns.Add(dName + "權數");
                dt.Columns.Add(dName + "成績");
                dt.Columns.Add(dName + "原始成績");
                dt.Columns.Add(dName + "補考成績");
                dt.Columns.Add(dName + "領域成績等第");
                dt.Columns.Add(dName + "原始成績等第");
                dt.Columns.Add(dName + "補考成績等第");
                dt.Columns.Add(dName + "領域需補考標示");
                dt.Columns.Add(dName + "領域補考成績標示");
                dt.Columns.Add(dName + "領域不及格標示");

            }

            foreach (string dName in DomainNameList)
            {
                for (int j = 1; j <= 12; j++)
                {
                    dt.Columns.Add(dName + "_科目名稱" + j);
                    dt.Columns.Add(dName + "_科目領域" + j);
                    dt.Columns.Add(dName + "_科目節數" + j);
                    dt.Columns.Add(dName + "_科目權數" + j);
                    dt.Columns.Add(dName + "_科目成績" + j);
                    dt.Columns.Add(dName + "_科目原始成績" + j);
                    dt.Columns.Add(dName + "_科目補考成績" + j);
                    dt.Columns.Add(dName + "_科目成績等第" + j);
                    dt.Columns.Add(dName + "_科目原始成績等第" + j);
                    dt.Columns.Add(dName + "_科目補考成績等第" + j);
                    dt.Columns.Add(dName + "_科目需補考標示" + j);
                    dt.Columns.Add(dName + "_科目補考成績標示" + j);
                    dt.Columns.Add(dName + "_科目不及格標示" + j);
                }
            }

            // 缺曠欄位
            foreach(string aa in alist)
            {
                foreach (string pp in plist)
                {
                    string key = pp + "_" + aa;
                    if (!dt.Columns.Contains(key))
                        dt.Columns.Add(key);
                }
            }

            //日常生活表現欄位
            foreach (string key in Global.DLBehaviorRef.Keys)
            {
                dt.Columns.Add(key + "_Name");
                dt.Columns.Add(key + "_Description");
            }

            //日常生活表現子項目欄位
            foreach (string key in _DLBehaviorConfigNameDict.Keys)
            {
                int itemIndex = 0;

                if (_DLBehaviorConfigItemNameDict.ContainsKey(key))
                {
                    foreach (string item in _DLBehaviorConfigItemNameDict[key])
                    {
                        itemIndex++;
                        dt.Columns.Add(key + "_Item_Name" + itemIndex);
                        dt.Columns.Add(key + "_Item_Degree" + itemIndex);
                        dt.Columns.Add(key + "_Item_Description" + itemIndex);
                    }
                }
            }

            // 

            List<string> rank_typeList = new List<string>();
            rank_typeList.Add("班排名");
            rank_typeList.Add("年排名");
            rank_typeList.Add("類別1排名");
            rank_typeList.Add("類別2排名");

            List<string> item_typeList = new List<string>();
            item_typeList.Add("科目成績");
            item_typeList.Add("科目成績(原始)");
            item_typeList.Add("領域成績");
            item_typeList.Add("領域成績(原始)");
            item_typeList.Add("總計成績");
            item_typeList.Add("總計成績(原始)");

            //item_name：
            //科目成績：科目名稱
            //領域成績:領域名稱
            //總計成績：學習領域總成績、課程學習總成績

            // 排名、母數、pr、五標、組距
            List<string> r2List = new List<string>();
            r2List.Add("rank");
            r2List.Add("matrix_count");
            r2List.Add("pr");
            r2List.Add("percentile");
            r2List.Add("avg_top_25");
            r2List.Add("avg_top_50");
            r2List.Add("avg");
            r2List.Add("avg_bottom_50");
            r2List.Add("avg_bottom_25");
            r2List.Add("level_gte100");
            r2List.Add("level_90");
            r2List.Add("level_80");
            r2List.Add("level_70");
            r2List.Add("level_60");
            r2List.Add("level_50");
            r2List.Add("level_40");
            r2List.Add("level_30");
            r2List.Add("level_20");
            r2List.Add("level_10");
            r2List.Add("level_lt10");

            // 科目排
            for (int i = 1; i <= Global.SupportSubjectCount; i++)
            {
                foreach (string rn in rank_typeList)
                {
                    foreach (string r2 in r2List)
                    {
                        dt.Columns.Add("科目成績" + i + "_" + rn + "_" + r2);
                        dt.Columns.Add("科目成績(原始)" + i + "_" + rn + "_" + r2);
                    }
                }
            }

            // 領域排
            for (int i = 1; i <= Global.SupportDomainCount; i++)
            {
                foreach (string rn in rank_typeList)
                {
                    foreach (string r2 in r2List)
                    {
                        dt.Columns.Add("領域成績" + i + "_" + rn + "_" + r2);
                        dt.Columns.Add("領域成績(原始)" + i + "_" + rn + "_" + r2);
                    }
                }
            }

            // 總計成績
            foreach (string rn in rank_typeList)
            {
                foreach (string r2 in r2List)
                {
                    dt.Columns.Add("總計成績_學習領域總成績" + "_" + rn + "_" + r2);
                    dt.Columns.Add("總計成績(原始)_學習領域總成績" + "_" + rn + "_" + r2);
                    dt.Columns.Add("總計成績_課程學習總成績" + "_" + rn + "_" + r2);
                    dt.Columns.Add("總計成績(原始)_課程學習總成績" + "_" + rn + "_" + r2);
                }
            }


            List<string> classIDs = _Students.Select(x => x.RefClassID).Distinct().ToList();
            //班級 catch
            Dictionary<string, ClassRecord> classDic = new Dictionary<string, ClassRecord>();
            foreach (ClassRecord cr in K12.Data.Class.SelectByIDs(classIDs))
            {
                if (!classDic.ContainsKey(cr.ID))
                    classDic.Add(cr.ID, cr);
            }

            string printDateTime = SelectTime();
            string schoolName = K12.Data.School.ChineseName;
            string 校長 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText;
            string 教務主任 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("EduDirectorName").InnerText;


            _bgWorkReport.ReportProgress(40);


            //基本資料
            foreach (StudentRecord student in _Students)
            {
                DataRow row = dt.NewRow();
                ClassRecord myClass = classDic.ContainsKey(student.RefClassID) ? classDic[student.RefClassID] : new ClassRecord();
                TeacherRecord myTeacher = myClass.Teacher != null ? myClass.Teacher : new TeacherRecord();

                row["列印日期"] = printDateTime;
                row["學校名稱"] = schoolName;
                row["學年度"] = _SelSchoolYear;
                row["學期"] = _SelSemester;
                row["系統編號"] = "系統編號{" + student.ID + "}";
                row["姓名"] = student.Name;
                row["英文姓名"] = student.EnglishName;
                row["班級"] = myClass.Name + "";
                row["班導師"] = myTeacher.Name + "";
                row["座號"] = student.SeatNo + "";
                row["學號"] = student.StudentNumber;

                row["校長"] = 校長;
                row["教務主任"] = 教務主任;


                dt.Rows.Add(row);

                _StudentRowDict.Add(student.ID, row);
            }


            //上課天數
            foreach (SemesterHistoryRecord shr in K12.Data.SemesterHistory.SelectByStudents(_Students))
            {
                DataRow row = _StudentRowDict[shr.RefStudentID];

                foreach (SemesterHistoryItem shi in shr.SemesterHistoryItems)
                {
                    if (shi.SchoolYear == _SelSchoolYear && shi.Semester == _SelSemester)
                        row["上課天數"] = shi.SchoolDayCount + "";
                }
            }


            //學期科目及領域成績
            foreach (JHSemesterScoreRecord jsr in SemesterScoreRecordList)
            {
                DataRow row = _StudentRowDict[jsr.RefStudentID];
                _文字描述.Clear();

                //學習領域成績

                row["學習領域成績"] = jsr.LearnDomainScore.HasValue ? jsr.LearnDomainScore.Value + "" : string.Empty;
                row["課程學習成績"] = jsr.CourseLearnScore.HasValue ? jsr.CourseLearnScore.Value + "" : string.Empty;


                row["學習領域原始成績"] = jsr.LearnDomainScoreOrigin.HasValue ? jsr.LearnDomainScoreOrigin.Value + "" : string.Empty;
                row["課程學習原始成績"] = jsr.CourseLearnScoreOrigin.HasValue ? jsr.CourseLearnScoreOrigin.Value + "" : string.Empty;

                // 收集領域科目成績給領域科目對照時使用
                Dictionary<string, DomainScore> DomainScoreDict = new Dictionary<string, DomainScore>();
                Dictionary<string, List<SubjectScore>> DomainSubjScoreDict = new Dictionary<string, List<SubjectScore>>();

                // 取得學期成績排名、五標、分數區間
                Dictionary<string, Dictionary<string, DataRow>> SemsScoreRankMatrixDataDict = Utility.GetSemsScoreRankMatrixData(_Configure.SchoolYear, _Configure.Semester, _StudentIDList);


                _bgWorkReport.ReportProgress(60);

                #region 科目成績照領域排序
                var jsSubjects = new List<SubjectScore>(jsr.Subjects.Values);
                var domainList = new Dictionary<string, int>();
                int num = 100;
                foreach (string dn in DomainNameList)
                {
                    domainList.Add(dn, num);
                    num--;
                }

                jsSubjects.Sort(delegate (SubjectScore r1, SubjectScore r2)
                {
                    decimal rank1 = 0;
                    decimal rank2 = 0;

                    if (r1.Credit != null)
                        rank1 += r1.Credit.Value;
                    if (r2.Credit != null)
                        rank2 += r2.Credit.Value;

                    if (domainList.ContainsKey(r1.Domain))
                        rank1 += domainList[r1.Domain];
                    if (domainList.ContainsKey(r2.Domain))
                        rank2 += domainList[r2.Domain];

                    if (rank1 == rank2)
                        return r2.Subject.CompareTo(r1.Subject);
                    else
                        return rank2.CompareTo(rank1);
                });
                #endregion
                //科目成績
                int count = 0;
                foreach (SubjectScore subj in jsSubjects)
                {
                    string ssNmae = subj.Domain;
                    if (string.IsNullOrEmpty(ssNmae))
                        ssNmae = "彈性課程";
                    if (!DomainSubjScoreDict.ContainsKey(ssNmae))
                        DomainSubjScoreDict.Add(ssNmae, new List<SubjectScore>());

                    DomainSubjScoreDict[ssNmae].Add(subj);

                    count++;

                    //超過就讓它爆炸
                    if (count > Global.SupportSubjectCount)
                        throw new Exception("超過支援列印科目數量: " + Global.SupportSubjectCount);

                    row["科目名稱" + count] = subj.Subject;
                    row["科目領域" + count] = string.IsNullOrWhiteSpace(subj.Domain) ? "彈性課程" : subj.Domain;
                    row["科目節數" + count] = subj.Period + "";
                    row["科目權數" + count] = subj.Credit + "";
                    row["科目成績" + count] = subj.Score.HasValue ? subj.Score.Value + "" : string.Empty;
                    row["科目原始成績" + count] = subj.ScoreOrigin.HasValue ? subj.ScoreOrigin.Value + "" : string.Empty;
                    row["科目補考成績" + count] = subj.ScoreMakeup.HasValue ? subj.ScoreMakeup.Value + "" : string.Empty;
                    row["科目成績等第" + count] = subj.Score.HasValue ? _ScoreMappingConfig.ParseScoreName(subj.Score.Value) : string.Empty;
                    row["科目原始成績等第" + count] = subj.ScoreOrigin.HasValue ? _ScoreMappingConfig.ParseScoreName(subj.ScoreOrigin.Value) : string.Empty;
                    row["科目補考成績等第" + count] = subj.ScoreMakeup.HasValue ? _ScoreMappingConfig.ParseScoreName(subj.ScoreMakeup.Value) : string.Empty;

                    if (subj.ScoreMakeup.HasValue)
                        row["科目補考成績標示" + count] = _Configure.ReScoreMark;

                    if (subj.Score.HasValue && subj.Score.Value < 60)
                    {
                        row["科目需補考標示" + count] = _Configure.NeeedReScoreMark;
                        row["科目不及格標示" + count] = _Configure.FailScoreMark;
                    }

                }


                // 處理領域科目並列               
                foreach (string dName in DomainNameList)
                {
                    if (DomainSubjScoreDict.ContainsKey(dName))
                    {
                        int si = 1;
                        foreach (SubjectScore ss in DomainSubjScoreDict[dName])
                        {
                            row[dName + "_科目名稱" + si] = ss.Subject;
                            row[dName + "_科目領域" + si] = ss.Domain;
                            row[dName + "_科目節數" + si] = ss.Period + "";
                            row[dName + "_科目權數" + si] = ss.Credit + "";
                            row[dName + "_科目成績" + si] = ss.Score.HasValue ? ss.Score.Value + "" : string.Empty;
                            row[dName + "_科目原始成績" + si] = ss.ScoreMakeup.HasValue ? ss.ScoreMakeup.Value + "" : string.Empty;
                            row[dName + "_科目補考成績" + si] = ss.ScoreMakeup.HasValue ? ss.ScoreMakeup.Value + "" : string.Empty;
                            row[dName + "_科目成績等第" + si] = ss.Score.HasValue ? _ScoreMappingConfig.ParseScoreName(ss.Score.Value) : string.Empty;
                            row[dName + "_科目原始成績等第" + si] = ss.ScoreOrigin.HasValue ? _ScoreMappingConfig.ParseScoreName(ss.ScoreOrigin.Value) : string.Empty;
                            row[dName + "_科目補考成績等第" + si] = ss.ScoreMakeup.HasValue ? _ScoreMappingConfig.ParseScoreName(ss.ScoreMakeup.Value) : string.Empty;

                            if (ss.Score.HasValue && ss.Score.Value < 60)
                            {
                                row[dName + "_科目需補考標示" + si] = _Configure.NeeedReScoreMark;
                                row[dName + "_科目不及格標示" + si] = _Configure.FailScoreMark;
                            }

                            if (ss.ScoreMakeup.HasValue)
                                row["科目補考成績標示" + count] = _Configure.ReScoreMark;

                            si++;
                        }
                    }
                }


                count = 0;
                foreach (DomainScore domain in jsr.Domains.Values)
                {
                    if (!DomainScoreDict.ContainsKey(domain.Domain))
                        DomainScoreDict.Add(domain.Domain, domain);

                    count++;

                    //超過就讓它爆炸
                    if (count > Global.SupportDomainCount)
                        throw new Exception("超過支援列印領域數量: " + Global.SupportDomainCount);

                    row["領域名稱" + count] = domain.Domain;
                    row["領域節數" + count] = domain.Period + "";
                    row["領域權數" + count] = domain.Credit + "";
                    row["領域成績" + count] = domain.Score.HasValue ? domain.Score.Value + "" : string.Empty;
                    row["領域原始成績" + count] = domain.ScoreOrigin.HasValue ? domain.ScoreOrigin.Value + "" : string.Empty;
                    row["領域補考成績" + count] = domain.ScoreMakeup.HasValue ? domain.ScoreMakeup.Value + "" : string.Empty;
                    row["領域成績等第" + count] = domain.Score.HasValue ? _ScoreMappingConfig.ParseScoreName(domain.Score.Value) : string.Empty;
                    row["領域原始成績等第" + count] = domain.ScoreOrigin.HasValue ? _ScoreMappingConfig.ParseScoreName(domain.ScoreOrigin.Value) : string.Empty;
                    row["領域補考成績等第" + count] = domain.ScoreMakeup.HasValue ? _ScoreMappingConfig.ParseScoreName(domain.ScoreMakeup.Value) : string.Empty;

                    if (domain.Score.HasValue && domain.Score.Value < 60)
                    {
                        row["領域需補考標示" + count] = _Configure.NeeedReScoreMark;
                        row["領域不及格標示" + count] = _Configure.FailScoreMark;
                    }

                    if (domain.ScoreMakeup.HasValue)
                        row["領域補考成績標示" + count] = _Configure.ReScoreMark;

                    if (!string.IsNullOrWhiteSpace(domain.Text))
                        _文字描述.Add(domain.Domain + " : " + domain.Text);
                }

                // 處理指定領域
                foreach (string dName in DomainNameList)
                {
                    if (DomainScoreDict.ContainsKey(dName))
                    {
                        DomainScore domain = DomainScoreDict[dName];

                        row[dName + "領域"] = domain.Domain;
                        row[dName + "節數"] = domain.Period + "";
                        row[dName + "權數"] = domain.Credit + "";
                        row[dName + "成績"] = domain.Score.HasValue ? domain.Score.Value + "" : string.Empty;
                        row[dName + "原始成績"] = domain.ScoreOrigin.HasValue ? domain.ScoreOrigin.Value + "" : string.Empty;
                        row[dName + "補考成績"] = domain.ScoreMakeup.HasValue ? domain.ScoreMakeup.Value + "" : string.Empty;
                        row[dName + "領域成績等第"] = domain.Score.HasValue ? _ScoreMappingConfig.ParseScoreName(domain.Score.Value) : string.Empty;
                        row[dName + "原始成績等第"] = domain.ScoreOrigin.HasValue ? _ScoreMappingConfig.ParseScoreName(domain.ScoreOrigin.Value) : string.Empty;
                        row[dName + "補考成績等第"] = domain.ScoreMakeup.HasValue ? _ScoreMappingConfig.ParseScoreName(domain.ScoreMakeup.Value) : string.Empty;

                        if (domain.Score.HasValue && domain.Score.Value < 60)
                        {
                            row[dName + "領域需補考標示"] = _Configure.NeeedReScoreMark;
                            row[dName + "領域不及格標示"] = _Configure.FailScoreMark;
                        }

                        if (domain.ScoreMakeup.HasValue)
                            row[dName + "領域補考成績標示"] = _Configure.ReScoreMark;

                    }
                }


                row["文字描述"] = string.Join(Environment.NewLine, _文字描述);
            }

            _bgWorkReport.ReportProgress(80);
            //預設學年度學期物件
            JHSchool.Behavior.BusinessLogic.SchoolYearSemester sysm = new JHSchool.Behavior.BusinessLogic.SchoolYearSemester(_SelSchoolYear, _SelSemester);

            //AutoSummary
            foreach (AutoSummaryRecord asr in AutoSummary.Select(_Students.Select(x => x.ID), new JHSchool.Behavior.BusinessLogic.SchoolYearSemester[] { sysm }))
            {
                DataRow row = _StudentRowDict[asr.RefStudentID];

                //缺曠
                foreach (AbsenceCountRecord acr in asr.AbsenceCounts)
                {
                    string key = Global.GetKey(acr.PeriodType, acr.Name);

                    if (dt.Columns.Contains(key))
                    {
                        int count = 0;
                        int.TryParse(row[key] + "", out count);

                        count += acr.Count;
                        row[key] = count;
                    }
                }

                //獎懲
                row["大功"] = asr.MeritA;
                row["小功"] = asr.MeritB;
                row["嘉獎"] = asr.MeritC;
                row["大過"] = asr.DemeritA;
                row["小過"] = asr.DemeritB;
                row["警告"] = asr.DemeritC;

                //日常生活表現
                JHMoralScoreRecord msr = asr.MoralScore;
                XmlElement textScore = (msr != null && msr.TextScore != null) ? msr.TextScore : K12.Data.XmlHelper.LoadXml("<TextScore/>");

                foreach (string key in Global.DLBehaviorRef.Keys)
                    SetDLBehaviorData(key, Global.DLBehaviorRef[key], textScore, row);
            }

            _bgWorkReport.ReportProgress(100);

            #region 將 DataTable 內合併欄位產生出來
            StreamWriter sw = new StreamWriter(Application.StartupPath + "\\學期成績單合併欄位.txt");
            foreach (DataColumn dc in dt.Columns)
                sw.WriteLine(dc.Caption);

            sw.Close();
            #endregion

            Document doc = _Configure.Template;
            doc.MailMerge.Execute(dt);
            doc.MailMerge.DeleteFields();
            e.Result = doc;
        }


        /// <summary>
        /// 取得日常生活表現設定名稱
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> GetDLBehaviorConfigNameDict()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            try
            {
                _DLBehaviorConfigItemNameDict.Clear();

                // 包含新竹與高雄
                K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];
                if (!string.IsNullOrEmpty(cd["DailyBehavior"]))
                {
                    string key = "日常行為表現";
                    //日常行為表現
                    XElement e1 = XElement.Parse(cd["DailyBehavior"]);
                    string name = e1.Attribute("Name").Value;
                    retVal.Add(key, name);

                    // 日常生活表現子項目
                    List<string> items = ParseItems(e1);
                    if (items.Count > 0)
                        _DLBehaviorConfigItemNameDict.Add(key, items);
                }

                if (!string.IsNullOrEmpty(cd["GroupActivity"]))
                {
                    string key = "團體活動表現";
                    //團體活動表現
                    XElement e4 = XElement.Parse(cd["GroupActivity"]);
                    string name = e4.Attribute("Name").Value;
                    retVal.Add(key, name);

                    // 團體活動表現
                    List<string> items = ParseItems(e4);
                    if (items.Count > 0)
                        _DLBehaviorConfigItemNameDict.Add(key, items);

                }

                if (!string.IsNullOrEmpty(cd["PublicService"]))
                {
                    string key = "公共服務表現";
                    //公共服務表現
                    XElement e5 = XElement.Parse(cd["PublicService"]);
                    string name = e5.Attribute("Name").Value;
                    retVal.Add(key, name);
                    List<string> items = ParseItems(e5);
                    if (items.Count > 0)
                        _DLBehaviorConfigItemNameDict.Add(key, items);

                }

                if (!string.IsNullOrEmpty(cd["SchoolSpecial"]))
                {
                    string key = "校內外特殊表現";
                    //校內外特殊表現,新竹沒有子項目，高雄有子項目
                    XElement e6 = XElement.Parse(cd["SchoolSpecial"]);
                    string name = e6.Attribute("Name").Value;
                    retVal.Add(key, name);
                    List<string> items = ParseItems(e6);
                    if (items.Count > 0)
                        _DLBehaviorConfigItemNameDict.Add(key, items);
                }

                if (!string.IsNullOrEmpty(cd["OtherRecommend"]))
                {
                    //其他表現
                    XElement e2 = XElement.Parse(cd["OtherRecommend"]);
                    string name = e2.Attribute("Name").Value;
                    retVal.Add("其他表現", name);
                }

                if (!string.IsNullOrEmpty(cd["DailyLifeRecommend"]))
                {
                    //日常生活表現具體建議
                    XElement e3 = XElement.Parse(cd["DailyLifeRecommend"]);
                    string name = e3.Attribute("Name").Value;
                    retVal.Add("具體建議", name);  // 高雄
                    retVal.Add("綜合評語", name);  // 新竹
                }
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("日常生活表現設定檔解析失敗!" + ex.Message);
            }

            return retVal;
        }

        /// <summary>
        /// XML 內解析子項目名稱
        /// </summary>
        /// <param name="elm"></param>
        /// <returns></returns>
        private static List<string> ParseItems(XElement elm)
        {
            List<string> retVal = new List<string>();

            foreach (XElement subElm in elm.Elements("Item"))
            {
                // 因為社團功能，所以要將"社團活動" 字不放入
                string name = subElm.Attribute("Name").Value;
                if (name != "社團活動")
                    retVal.Add(name);
            }
            return retVal;
        }

        /// <summary>
        /// 填寫日常生活表現資料
        /// </summary>
        /// <param name="name"></param>
        /// <param name="path"></param>
        /// <param name="textScore"></param>
        /// <param name="row"></param>

        private static void SetDLBehaviorData(string name, string path, XmlElement textScore, DataRow row)
        {
            row[name + "_Name"] = _DLBehaviorConfigNameDict.ContainsKey(name) ? _DLBehaviorConfigNameDict[name] : string.Empty;

            if (_DLBehaviorConfigItemNameDict.ContainsKey(name))
            {
                int index = 0;
                foreach (string itemName in _DLBehaviorConfigItemNameDict[name])
                {
                    foreach (XmlElement item in textScore.SelectNodes(path))
                    {
                        if (itemName == item.GetAttribute("Name"))
                        {
                            index++;
                            row[name + "_Item_Name" + index] = itemName;
                            row[name + "_Item_Degree" + index] = item.GetAttribute("Degree");
                            row[name + "_Item_Description" + index] = item.GetAttribute("Description");
                        }
                    }
                }
            }
            else if (_DLBehaviorConfigNameDict.ContainsKey(name))
            {
                string value = _DLBehaviorConfigNameDict[name];

                foreach (XmlElement item in textScore.SelectNodes(path))
                {
                    if (value == item.GetAttribute("Name"))
                        row[name + "_Description"] = item.GetAttribute("Description");
                }
            }
        }


        private static string SelectTime() //取得Server的時間
        {
            QueryHelper qh = new QueryHelper();
            DataTable dtable = qh.Select("select now()"); //取得時間
            DateTime dt = DateTime.Now;
            DateTime.TryParse("" + dtable.Rows[0][0], out dt); //Parse資料
            string ComputerSendTime = dt.ToString("yyyy/MM/dd"); //最後時間

            return ComputerSendTime;
        }



        private void _bgWorkerLoadData_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績報表資料讀取中...", e.ProgressPercentage);
        }

        private void _bgWorkerLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_Configure == null)
                _Configure = new Configure();

            _DefalutSchoolYear = K12.Data.School.DefaultSchoolYear;
            _DefaultSemester = K12.Data.School.DefaultSemester;

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
                }

                for (int j = 0; j < 3; j++)
                {
                    cboSchoolYear.Items.Add("" + (i + j));
                }

            }

            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");


            string userSelectConfigName = "";
            // 檢查畫面上是否有使用者選的
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    userSelectConfigName = conf.Name;
                    break;
                }

            if (!string.IsNullOrEmpty(userSelectConfigName))
                cboConfigure.Text = userSelectConfigName;

            EnableSelectItem(true);
        }

        private void _bgWorkerLoadData_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorkerLoadData.ReportProgress(1);
            // 檢查預設樣板是否存在
            _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);

            // 沒有設定檔，建立預設設定檔
            if (_UDTConfigList.Count < 2)
            {
                _bgWorkerLoadData.ReportProgress(20);
                foreach (string name in Global.DefaultConfigNameList())
                {
                    Configure cn = new Configure();
                    cn.Name = name;
                    cn.SchoolYear = K12.Data.School.DefaultSchoolYear;
                    cn.Semester = K12.Data.School.DefaultSemester;
                    DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                    conf.Name = name;
                    conf.UDTTableName = Global._UDTTableName;
                    conf.ProjectName = Global._ProjectName;
                    conf.Type = Global._DefaultConfTypeName;
                    _UDTConfigList.Add(conf);

                    // 設預設樣板
                    switch (name)
                    {
                        //case "領域成績單":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_領域成績單));
                        //    break;

                        //case "科目成績單":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目成績單));
                        //    break;

                        //case "科目及領域成績單_領域組距":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_領域組距));
                        //    break;
                        //case "科目及領域成績單_科目組距":
                        //    cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_科目組距));
                        //    break;
                    }

                    if (cn.Template == null)
                        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹學期成績單樣板_固定排名_科目_領域));
                    cn.Encode();
                    cn.Save();
                }
                if (_UDTConfigList.Count > 0)
                    DAO.UDTTransfer.InsertConfigData(_UDTConfigList);
            }
            _bgWorkerLoadData.ReportProgress(70);
            // 取的設定資料
            _ConfigureList = _AccessHelper.Select<Configure>();

            _bgWorkerLoadData.ReportProgress(100);
        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            EnableSelectItem(false);
            _SelSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _SelSemester = int.Parse(K12.Data.School.DefaultSemester);
            _ScoreMappingConfig.LoadData();
            _bgWorkerLoadData.RunWorkerAsync();

        }


        // 啟用項目
        private void EnableSelectItem(bool enable)
        {
            cboConfigure.Enabled = enable;
            cboSchoolYear.Enabled = enable;
            cboSemester.Enabled = enable;
            lnkViewTemplate.Enabled = enable;
            lnkChangeTemplate.Enabled = enable;
            lnkViewMapColumns.Enabled = enable;
            btnSaveConfig.Enabled = enable;
            btnPrint.Enabled = enable;
            txtFailScoreMark.Enabled = enable;
            txtNeeedReScoreMark.Enabled = enable;
            txtReScoreMark.Enabled = enable;
        }



        private void lnkCopyConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = _Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;
                conf.PrintSubjectList.AddRange(_Configure.PrintSubjectList);
                conf.SchoolYear = _Configure.SchoolYear;
                conf.Semester = _Configure.Semester;
                conf.SubjectLimit = _Configure.SubjectLimit;
                conf.Template = _Configure.Template;
                conf.NeeedReScoreMark = _Configure.NeeedReScoreMark;
                conf.ReScoreMark = _Configure.ReScoreMark;
                conf.FailScoreMark = _Configure.FailScoreMark;

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


        private void btnPrint_Click(object sender, EventArgs e)
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


            SaveTemplate(null, null);

            btnSaveConfig.Enabled = false;


            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
            // 執行報表
            _bgWorkReport.RunWorkerAsync();
        }

        // 儲存樣板
        private void SaveTemplate(object sender, EventArgs e)
        {
            if (_Configure == null) return;
            _Configure.SchoolYear = cboSchoolYear.Text;
            _Configure.Semester = cboSemester.Text;
            _Configure.SelSetConfigName = cboConfigure.Text;
            _Configure.NeeedReScoreMark = txtNeeedReScoreMark.Text;
            _Configure.ReScoreMark = txtReScoreMark.Text;
            _Configure.FailScoreMark = txtFailScoreMark.Text;

            if (_Configure.PrintAttendanceList == null)
                _Configure.PrintAttendanceList = new List<string>();

            _Configure.PrintAttendanceList.Clear();

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
                    if (_Configure.PrintAttendanceList == null)
                        _Configure.PrintAttendanceList = new List<string>();
                    if (_Configure.PrintSubjectList == null)
                        _Configure.PrintSubjectList = new List<string>();


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
                    txtFailScoreMark.Text = _Configure.FailScoreMark;
                    txtReScoreMark.Text = _Configure.ReScoreMark;
                    txtNeeedReScoreMark.Text = _Configure.NeeedReScoreMark;
                }
                else
                {
                    _Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;

                }
            }
        }

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "新竹學期成績單樣板(" + _Configure.Name + ")";

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
                        stream.Write(Properties.Resources.新竹學期成績單樣板_固定排名_科目_領域, 0, Properties.Resources.新竹學期成績單樣板_固定排名_科目_領域.Length);
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

            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
        }


    }
}
