﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using HsinChu.ClassExamScoreAvgComparison.Model;
using HsinChu.JHEvaluation.Data;
using JHSchool.Evaluation.Ranking;
using K12.Data;
using FISCA.Data;
using HsinChu.ClassExamScoreAvgComparison.DAL;
using System.Xml;

namespace HsinChu.ClassExamScoreAvgComparison
{
    public partial class MainForm : BaseForm
    {
        private Config _config;
        private List<JHClassRecord> _classes;
        private List<ClassExamScoreData> _data;
        private BackgroundWorker _worker;
        private List<JHExamRecord> _exams;
        //private Dictionary<string, List<string>> _ecMapping;
        List<string> subjectNameList = DALTransfer.GetSubjectList();
        private List<string> _courseList;

        private Dictionary<string, JHCourseRecord> _courseDict;

        private int _runningSchoolYear;
        private int _runningSemester;

        public  Dictionary<string, DAL.ScoreMap> ScoreTextMap = new Dictionary<string, DAL.ScoreMap>();
        public  Dictionary<decimal, DAL.ScoreMap> ScoreValueMap = new Dictionary<decimal, DAL.ScoreMap>();
        public static void Run()
        {
            new MainForm().ShowDialog();
        }

        public MainForm()
        {
            //GetSubjectList();
            InitializeComponent();
            InitializeSemester();
            this.Text = Global.ReportName;
            this.MinimumSize = this.Size;
            this.MaximumSize = this.Size;

            _config = new Config(Global.ReportName);

            _data = new List<ClassExamScoreData>();
            _courseDict = new Dictionary<string, JHCourseRecord>();
            _exams = new List<JHExamRecord>();
            //_ecMapping = new Dictionary<string, List<string>>();
            _courseList = new List<string>();

            cbExam.DisplayMember = "Name";
            cbSource.Items.Add("定期");
            cbSource.Items.Add("定期加平時");
            cbSource.SelectedIndex = 0;

            cbExam.Items.Add("");
            _exams = JHExam.SelectAll();
            foreach (var exam in _exams)
                cbExam.Items.Add(exam);
            cbExam.SelectedIndex = 0;

            _worker = new BackgroundWorker();
            _worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                JHExamRecord exam = e.Argument as JHExamRecord;
                #region 取得試別
                //_ecMapping.Clear();
                //_exams = JHExam.SelectAll();
                //List<string> examIDs = new List<string>();
                //foreach (JHExamRecord exam in _exams)
                //{
                //    examIDs.Add(exam.ID);
                //_ecMapping.Add(exam.ID, new List<string>());
                //}
                #endregion

                #region 取得課程
                _courseDict.Clear();
                List<JHCourseRecord> courseList = JHCourse.SelectBySchoolYearAndSemester(_runningSchoolYear, _runningSemester);
                List<string> courseIDs = new List<string>();
                foreach (JHCourseRecord course in courseList)
                {
                    courseIDs.Add(course.ID);
                    _courseDict.Add(course.ID, course);
                }
                #endregion

                #region 取得評量成績
                //StudentID -> ClassExamScoreData
                Dictionary<string, ClassExamScoreData> scMapping = new Dictionary<string, ClassExamScoreData>();
                List<string> ids = new List<string>();
                _classes = JHClass.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);

                // 排序
                if (_classes.Count > 1)
                    _classes = DAL.DALTransfer.ClassRecordSortByDisplayOrder(_classes);

                // TODO: 這邊要排序
                //List<K12.Data.ClassRecord> c = new List<K12.Data.ClassRecord>(_classes);
                //c.Sort();
                //_classes = new List<JHClassRecord>(c);
                //((List<K12.Data.ClassRecord>)_classes).Sort();

                _data.Clear();
                foreach (JHClassRecord cla in _classes)
                {
                    ClassExamScoreData classData = new ClassExamScoreData(cla);
                    foreach (JHStudentRecord stu in classData.Students)
                        scMapping.Add(stu.ID, classData);
                    _data.Add(classData);
                }

                //foreach (string examID in examIDs)
                //{

                _courseList.Clear();
                if (courseIDs.Count > 0)
                {
                    // TODO: JHSCETake 需要提供 SelectBy 課程IDs and 試別IDs 嗎？
                    foreach (JHSCETakeRecord sce in JHSCETake.SelectByCourseAndExam(courseIDs, exam.ID))
                    {
                        // TODO: 下面前兩個判斷應該可以拿掉
                        //if (!examIDs.Contains(sce.RefExamID)) continue; //試別無效
                        //if (!courseIDs.Contains(sce.RefCourseID)) continue; //課程無效
                        if (!scMapping.ContainsKey(sce.RefStudentID)) continue; //學生編號無效
                        if (string.IsNullOrEmpty(_courseDict[sce.RefCourseID].RefAssessmentSetupID)) continue; //課程無評量設定

                        if (!_courseList.Contains(sce.RefCourseID))
                            _courseList.Add(sce.RefCourseID);
                        //if (!_ecMapping[sce.RefExamID].Contains(sce.RefCourseID))
                        //    _ecMapping[sce.RefExamID].Add(sce.RefCourseID);

                        ClassExamScoreData classData = scMapping[sce.RefStudentID];
                        classData.AddScore(sce);
                    }
                }
                //}
                #endregion

                ScoreTextMap.Clear();
                ScoreValueMap.Clear();
                #region 取得評量成績缺考暨免試設定
                Framework.ConfigData cd = JHSchool.School.Configuration["評量成績缺考暨免試設定"];
                if (!string.IsNullOrEmpty(cd["評量成績缺考暨免試設定"]))
                {
                    XmlElement element = Framework.XmlHelper.LoadXml(cd["評量成績缺考暨免試設定"]);

                    foreach (XmlElement each in element.SelectNodes("Setting"))
                    {
                        var UseText = each.SelectSingleNode("UseText").InnerText;
                        var AllowCalculation = bool.Parse(each.SelectSingleNode("AllowCalculation").InnerText);
                        decimal Score;
                        decimal? NullableScore = null;
                        bool result = decimal.TryParse(each.SelectSingleNode("Score").InnerText, out Score);
                        if (result)
                        {
                            NullableScore = Score;
                        }
                        var Active = bool.Parse(each.SelectSingleNode("Active").InnerText);
                        var UseValue = decimal.Parse(each.SelectSingleNode("UseValue").InnerText);

                        if (Active)
                        {
                            if (!ScoreTextMap.ContainsKey(UseText))
                            {
                                ScoreTextMap.Add(UseText, new ScoreMap
                                {
                                    UseText = UseText,  //「缺」或「免」
                                    AllowCalculation = AllowCalculation,  //是否計算成績
                                    Score = NullableScore,  //計算成績時，應以多少分來計算
                                    Active = Active, //此設定是否啟用
                                    UseValue = UseValue, //代表「缺」或「免」的負數
                                });
                            }
                            if (!ScoreValueMap.ContainsKey(UseValue))
                            {
                                ScoreValueMap.Add(UseValue, new ScoreMap
                                {
                                    UseText = UseText,
                                    AllowCalculation = AllowCalculation,
                                    Score = NullableScore,
                                    Active = Active,
                                    UseValue = UseValue,
                                });
                            }
                        }
                    }
                }

                #endregion
            };
            _worker.RunWorkerCompleted += delegate
            {
                string running = _runningSchoolYear + "_" + _runningSemester;
                string current = (int)cboSchoolYear.SelectedItem + "_" + (int)cboSemester.SelectedItem;
                if (running != current)
                {
                    if (!_worker.IsBusy)
                    {
                        _runningSchoolYear = (int)cboSchoolYear.SelectedItem;
                        _runningSemester = (int)cboSemester.SelectedItem;

                        RunWorker();
                    }
                }
                else
                {
                    FillData();
                    ControlEnabled = true;
                }
            };

            RunWorker();
        }

        private void RunWorker()
        {
            if (cbExam.SelectedItem != null && cbExam.SelectedItem is JHExamRecord)
            {
                ControlEnabled = false;
                _worker.RunWorkerAsync(cbExam.SelectedItem as JHExamRecord);
            }
            else
            {
                lvSubject.Items.Clear();
                lvDomain.Items.Clear();
                
                ControlEnabled = true;
                btnPrint.Enabled = false;
            }
        }

        private void InitializeSemester()
        {
            try
            {
                cboSchoolYear.SuspendLayout();
                cboSemester.SuspendLayout();
                int schoolYear = int.Parse(School.DefaultSchoolYear);
                int semester = int.Parse(School.DefaultSemester);
                for (int i = -2; i <= 2; i++)
                    cboSchoolYear.Items.Add(schoolYear + i);
                cboSemester.Items.Add(1);
                cboSemester.Items.Add(2);

                cboSchoolYear.SelectedIndex = 2;
                cboSemester.SelectedIndex = semester - 1;
                cboSchoolYear.ResumeLayout();
                cboSemester.ResumeLayout();

                _runningSchoolYear = (int)cboSchoolYear.SelectedItem;
                _runningSemester = (int)cboSemester.SelectedItem;

                cboSchoolYear.SelectedIndexChanged += new EventHandler(Semester_SelectedIndexChanged);
                cboSemester.SelectedIndexChanged += new EventHandler(Semester_SelectedIndexChanged);
            }
            catch (Exception ex)
            {

            }
        }

        private void Semester_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_worker.IsBusy)
            {
                //cboSchoolYear.Enabled = false;
                //cboSemester.Enabled = false;
                //cbExam.Enabled = false;
                //cbSource.Enabled = false;
                //btnPrint.Enabled = false;
                //pictureBox1.Visible = true;
                //pictureBox2.Visible = true;

                _runningSchoolYear = (int)cboSchoolYear.SelectedItem;
                _runningSemester = (int)cboSemester.SelectedItem;
                RunWorker();
            }
        }

        private bool ControlEnabled
        {
            set
            {
                cboSchoolYear.Enabled = value;
                cboSemester.Enabled = value;
                cbExam.Enabled = value;
                cbSource.Enabled = value;
                btnPrint.Enabled = value;
                pictureBox1.Visible = !value;
                pictureBox2.Visible = !value;
            }
        }

        private void lnConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            _config.Load();
            new ConfigForm(_config).ShowDialog();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            #region 進行驗證
            if (!ValidItem())
            {
                MsgBox.Show("請選擇要列印的" + gpSubject.Text + "或" + gpDomain.Text);
                return;
            }
            #endregion
            Report._UserSelectCount = 0;

            Global.UserSelectSchoolYear = cboSchoolYear.Text;
            Global.UserSelectSemester = cboSemester.Text;

            #region 取得使用者選取的課程(科目)編號
            List<string> courseIDs = new List<string>();
            foreach (ListViewItem item in lvSubject.Items)
            {
                if (item.Checked)
                {
                    Report._UserSelectCount++;
                    List<string> list = item.Tag as List<string>;
                    courseIDs.AddRange(list);
                }
            }
            #endregion

            #region 取得使用者選取的領域
            List<string> domains = new List<string>();
            foreach (ListViewItem item in lvDomain.Items)
            {
                Report._UserSelectCount++;
                if (item.Checked)
                    domains.Add(item.Text);
            }
            #endregion

            btnPrint.Enabled = false;
            pictureBox1.Visible = true;
            pictureBox2.Visible = true;

            #region 轉換成 CourseScore、計算分數、排名
            JHExamRecord exam = cbExam.SelectedItem as JHExamRecord;
            //ScoreType type = GetScoreType();

            List<string> asIDs = new List<string>();
            Dictionary<string, string> asMapping = new Dictionary<string, string>();
            Dictionary<string, JHAEIncludeRecord> aeDict = new Dictionary<string, JHAEIncludeRecord>();
            foreach (string courseID in courseIDs)
            {
                string asID = _courseDict[courseID].RefAssessmentSetupID;
                asIDs.Add(asID);
                asMapping.Add(courseID, asID);
            }

            foreach (JHAEIncludeRecord record in JHAEInclude.SelectByAssessmentSetupIDs(asIDs))
            {
                if (record.RefExamID != exam.ID) continue;
                aeDict.Add(record.RefAssessmentSetupID, record);
            }

            //ComputeScore computer = new ComputeScore(_courseDict);
            //List<ComputeMethod> methods = GetMethods(_config);
            //RankMethod rankMethod = GetRankMethod(_config);

            Rank rank = new Rank();
            // TODO: 不確定排名是否接續
            rank.Sequence = false;

            // 取得評量比例
            Global.ScorePercentageHSDict = Global.GetScorePercentageHS();

            foreach (var ced in _data)
            {
                //轉成 CourseScore
                ced.ConvertToCourseScores(courseIDs, exam.ID);
                //RankData rd = new RankData();

                foreach (string studentID in ced.Rows.Keys)
                {
                    StudentRow row = ced.Rows[studentID];

                    //計算單一評量成績
                    foreach (CourseScore courseScore in row.CourseScoreList)
                    {
                        string asID = asMapping[courseScore.CourseID];
                        if (aeDict.ContainsKey(asID))
                            courseScore.CalculateScore(new HC.JHAEIncludeRecord(aeDict[asID]), "" + cbSource.SelectedItem, ScoreValueMap);
                    }
                }

                //排序班級課程ID
                ced.SortCourseIDs(courseIDs);
            }
            #endregion

            #region 產生報表
            //Report report = new Report(_data, _courseDict, exam, methods);
            Report report = new Report(_data, _courseDict, exam, domains, ScoreValueMap, cbSource.SelectedItem.ToString());
            report.GenerateCompleted += new EventHandler(report_GenerateCompleted);
            report.GenerateError += new EventHandler(report_GenerateError);
            report.Generate();
            #endregion
        }

        private void report_GenerateError(object sender, EventArgs e)
        {
            btnPrint.Enabled = true;
            pictureBox1.Visible = false;
            pictureBox2.Visible = false;
        }

        private void report_GenerateCompleted(object sender, EventArgs e)
        {
            btnPrint.Enabled = true;
            pictureBox1.Visible = false;
            pictureBox2.Visible = false;
        }

        private RankMethod GetRankMethod(Config _config)
        {
            if (_config.RankMethod == "加權平均") return RankMethod.加權平均;
            else if (_config.RankMethod == "加權總分") return RankMethod.加權總分;
            else if (_config.RankMethod == "合計總分") return RankMethod.合計總分;
            else if (_config.RankMethod == "算術平均") return RankMethod.算術平均;
            else
            {
                throw new Exception("無效的排名依據");
            }
        }

        private List<ComputeMethod> GetMethods(Config _config)
        {
            List<ComputeMethod> methods = new List<ComputeMethod>();
            foreach (string item in _config.PrintItems)
            {
                if (item == "加權平均") methods.Add(ComputeMethod.加權平均);
                else if (item == "加權總分") methods.Add(ComputeMethod.加權總分);
                else if (item == "合計總分") methods.Add(ComputeMethod.合計總分);
                else if (item == "算術平均") methods.Add(ComputeMethod.算術平均);
            }
            return methods;
        }

        /// <summary>
        /// 取得使用者選擇的分數類型
        /// </summary>
        /// <returns></returns>
        //private ScoreType GetScoreType()
        //{
        //    string type = "" + cbScore.SelectedItem;
        //    if (type == "定期") return ScoreType.定期;
        //    else if (type == "定期加平時") return ScoreType.定期加平時;
        //    else
        //    {
        //        throw new Exception("無效的分數類型");
        //    }
        //}

        /// <summary>
        /// 驗證，至少要選一個項目
        /// </summary>
        /// <returns></returns>
        private bool ValidItem()
        {
            foreach (ListViewItem item in lvSubject.Items)
            {
                if (item.Checked)
                    return true;
            }
            foreach (ListViewItem item in lvDomain.Items)
            {
                if (item.Checked)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 驗證設定
        /// </summary>
        /// <returns></returns>
        private bool ValidConfig()
        {
            return _config.HasValue;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FillData()
        {
            lvSubject.Items.Clear();
            lvDomain.Items.Clear();
            lvSubject.SuspendLayout();
            lvDomain.SuspendLayout();

            Dictionary<string, ListViewItem> subjectItems = new Dictionary<string, ListViewItem>();
            Dictionary<string, ListViewItem> domainItems = new Dictionary<string, ListViewItem>();

            //string id = (cb.SelectedItem as JHExamRecord).ID;
            foreach (string courseID in _courseList)
            {
                JHCourseRecord course = _courseDict[courseID];

                ListViewItem subject_item;
                if (!subjectItems.ContainsKey(course.Subject))
                {
                    subject_item = new ListViewItem(course.Subject);
                    List<string> ids = new List<string>();
                    subject_item.Tag = ids;
                    subjectItems.Add(course.Subject, subject_item);
                }
                subject_item = subjectItems[course.Subject];
                (subject_item.Tag as List<string>).Add(course.ID);

                if (!domainItems.ContainsKey(course.Domain))
                {
                    ListViewItem domain_item = new ListViewItem(course.Domain);
                    domainItems.Add(course.Domain, domain_item);
                }
            }
            List<ListViewItem> itemList = new List<ListViewItem>(subjectItems.Values);
            itemList.Sort(ItemSort);
            foreach (ListViewItem item in itemList)
            {
                (item.Tag as List<string>).Sort();
                lvSubject.Items.Add(item);
            }
            List<ListViewItem> itemList2 = new List<ListViewItem>(domainItems.Values);
            itemList2.Sort(ItemSortDomain);
            foreach (ListViewItem item in itemList2)
                lvDomain.Items.Add(item);
            lvSubject.ResumeLayout();
            lvDomain.ResumeLayout();
        }

        private void cbo_SelectedIndexChanged(object sender, EventArgs e)
        {
            RunWorker();

            //ComboBox cb = sender as ComboBox;

            //if (cb.SelectedItem != null)
            //{

            //}
        }
        /// <summary>
        /// 依照科目資料管理排序
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        private int ItemSort(ListViewItem x, ListViewItem y)
        {
            //return JHSchool.Evaluation.Subject.CompareSubjectOrdinal(x.Text, y.Text);
            //List<string> list = new List<string>(new string[] { "國語文", "國文", "英語文", "英文", "英語", "語文", "數學", "歷史", "公民", "地理", "社會", "藝術與人文", "理化", "生物", "自然與生活科技", "健康與體育", "綜合活動" });
            int ix = subjectNameList.IndexOf(x.Text);
            int iy = subjectNameList.IndexOf(y.Text);

            if (ix >= 0 && iy >= 0)
                return ix.CompareTo(iy);
            else if (ix >= 0)
                return -1;
            else if (iy >= 0)
                return 1;
            else
                return x.Text.CompareTo(y.Text);
        }
        /// <summary>
        /// 領域排序
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        private int ItemSortDomain(ListViewItem x, ListViewItem y)
        {
            //return JHSchool.Evaluation.Subject.CompareSubjectOrdinal(x.Text, y.Text);
            List<string> list = new List<string>(new string[] { "國語文", "英語文",  "語文", "數學",  "社會", "自然科學", "自然與生活科技", "藝術", "藝術與人文", "健康與體育", "綜合活動", "科技" });
            int ix = list.IndexOf(x.Text);
            int iy = list.IndexOf(y.Text);

            if (ix >= 0 && iy >= 0)
                return ix.CompareTo(iy);
            else if (ix >= 0)
                return -1;
            else if (iy >= 0)
                return 1;
            else
                return x.Text.CompareTo(y.Text);
        }

    }
}
