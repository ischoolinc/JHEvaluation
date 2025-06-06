﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using System.Threading;
using K12.Data;
using Campus.Rating;
using SubjectOrder = JHSchool.Evaluation.Subject;

namespace JHEvaluation.ClassSemesterScoreReport
{
    /// <summary>
    /// 此表單不適合繼承。
    /// </summary>
    public partial class MainForm : BaseForm
    {
        /// <summary>
        /// 報表設定。
        /// </summary>
        private ReportPreference Perference = null;

        /// <summary>
        /// 可選擇的科目清單。
        /// 用於第二執行緒收集，主執行緒填入 ListView用途。
        /// </summary>
        private List<string> Subjects = new List<string>();

        /// <summary>
        /// 用於科目對照領域。
        /// </summary>
        private Dictionary<string, string> SubjectDomainMap = new Dictionary<string, string>();

        /// <summary>
        /// 代表使用者選擇的學期資訊。
        /// </summary>
        private SemesterSelector Semester;

        /// <summary>
        /// 要列印的班級清單。
        /// </summary>
        private List<JHClassRecord> SelectedClasses;

        /// <summary>
        /// 因為要算年排名，所以要全校的學生(一般狀態)。
        /// </summary>
        private List<ReportStudent> AllStudents = new List<ReportStudent>();

        /// <summary>
        /// 下載目前學期資料的 BackgroundWorker。
        /// </summary>
        private BackgroundWorker MasterWorker = new BackgroundWorker();
        private bool WorkPadding = false; //是否有工作 Padding。

        List<string> _tmpUserSelSubjectList = new List<string>();
        List<string> _tmpUserSelDomainList = new List<string>();

        public static void Run(List<string> classes)
        {
            new MainForm(classes).ShowDialog();
        }

        public MainForm(List<string> classes)
        {
            InitializeComponent();

            //準備相關學生、班級資料。
            ReportStudent.SetClassMapping(JHClass.SelectAll());
            SelectedClasses = JHClass.SelectByIDs(classes); //要列印成績單的班級清單。
            AllStudents = JHStudent.SelectAll().ToReportStudent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

            cbxScoreType.Items.Add("原始成績");
            cbxScoreType.Items.Add("原始補考擇優");
            cbxScoreType.DropDownStyle = ComboBoxStyle.DropDownList;

            if (Perference == null)
                Perference = new ReportPreference();
            cbxScoreType.Text = Perference.UserSelScoreType;
            txtReMark.Text = Perference.ReScoreMark;

            Utilities.SetSemesterDefaultItems(cboSchoolYear, cboSemester); //顯示學年度學期選項。

            Semester = new SemesterSelector(cboSchoolYear, cboSemester);
            Semester.SemesterChanged += new EventHandler(Semester_SemesterChanged);

            MasterWorker.DoWork += new DoWorkEventHandler(MasterWorker_DoWork);
            MasterWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(MasterWorker_RunWorkerCompleted);

            ////報表設定。
            //Perference = new ReportPreference();

            FillCurrentSemesterData();
        }

        private void FillCurrentSemesterData()
        {
            Loading = true;
            Utilities.DisableControls(this);
            lvSubject.Items.Clear();
            lvDomain.Items.Clear();

            if (MasterWorker.IsBusy)
                WorkPadding = true;
            else
                MasterWorker.RunWorkerAsync();
        }

        private void MasterWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (WorkPadding)
            {
                MasterWorker.RunWorkerAsync();
                WorkPadding = false;
                return;
            }

            Loading = false;
            Utilities.EnableControls(this);

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
                return;
            }

            lvDomain.FillItems(Utilities.DomainNames, "領域");
            //lvSubject.FillItems(Subjects, "科目");
            lvSubject.FillSubjectItems(Subjects, "科目", SubjectDomainMap);
            
            // 回填使用者勾選
            foreach (ListViewItem lvi in lvSubject.Items)
                if (_tmpUserSelSubjectList.Contains(lvi.Text))
                    lvi.Checked = true;

            foreach (ListViewItem lvi in lvDomain.Items)
                if (_tmpUserSelDomainList.Contains(lvi.Text))
                    lvi.Checked = true;
            btnPrint.Enabled = true;
        }

        private void MasterWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //1.Goup By 可選擇的科目清單。
            //2.寫入成績資料到 ReportStudent 上。



            int schoolYear = Semester.SelectedSchoolYear;
            int semester = Semester.SelectedSemester;
            FunctionSpliter<string, JHSemesterScoreRecord> selectData = new FunctionSpliter<string, JHSemesterScoreRecord>(1000, 5);
            selectData.Function = delegate (List<string> ps)
            {
                return JHSemesterScore.SelectBySchoolYearAndSemester(ps, schoolYear, semester);
            };
            List<JHSemesterScoreRecord> semsScores = selectData.Execute(AllStudents.ToKeys());

            GroupBySubjects(semsScores);

            //先把學生身上的成績、排名相關資料清掉。
            foreach (ReportStudent each in AllStudents)
            {
                each.Clear();
                each.Scores.Add(Utilities.SubjectToken, new ScoreCollection()); //科目成績。
                each.Scores.Add(Utilities.DomainToken, new ScoreCollection());  //領域成績。
                each.Scores.Add(Utilities.SummaryToken, new ScoreCollection());   //運算後的成績。
            }

            //將成績填到學生身上。
            Dictionary<string, ReportStudent> dicAllStudent = AllStudents.ToDictionary();
            foreach (JHSemesterScoreRecord eachScore in semsScores)
            {
                //如果找不到該學生，跳到下一筆。
                if (!dicAllStudent.ContainsKey(eachScore.RefStudentID)) continue;

                ReportStudent student = dicAllStudent[eachScore.RefStudentID];

                //科目成績。
                foreach (SubjectScore each in eachScore.Subjects.Values)
                {
                    // 初始執
                    decimal ss = -1;

                    if (Perference.UserSelScoreType == "原始成績")
                    {
                        if (each.ScoreOrigin.HasValue)
                            ss = each.ScoreOrigin.Value;
                    }

                    if (Perference.UserSelScoreType == "原始補考擇優")
                    {
                        // 成績
                        if (each.Score.HasValue && each.Score.Value > ss)
                            ss = each.Score.Value;
                    }

                    //                    if (!each.Score.HasValue) continue; //沒有分數不處理。

                    if (ss == -1)
                        continue;

                    if (!each.Credit.HasValue || each.Credit.Value < 0) continue;  //沒有節數不處理。 //2021-07 要求權重0也要印出

                    if (!student.Scores[Utilities.SubjectToken].Contains(each.Subject))
                    {
                        student.Scores[Utilities.SubjectToken].Add(each.Subject, ss, each.Credit.Value);

                        if (Perference.UserSelScoreType == "原始補考擇優" && each.ScoreMakeup.HasValue)
                        {
                            student.Scores[Utilities.DomainToken].AddReExam(each.Subject, each.ScoreMakeup.Value);
                        }
                    }
                }

                //領域成績。
                foreach (DomainScore each in eachScore.Domains.Values)
                {
                    decimal dd = -1;
                    if (Perference.UserSelScoreType == "原始成績")
                    {
                        if (each.ScoreOrigin.HasValue)
                            dd = each.ScoreOrigin.Value;
                    }

                    if (Perference.UserSelScoreType == "原始補考擇優")
                    {
                        if (each.Score.HasValue && each.Score.Value > dd)
                            dd = each.Score.Value;
                    }

                    //if (!each.Score.HasValue) continue;
                    if (dd == -1)
                        continue;

                    if (!each.Credit.HasValue || each.Credit.Value < 0) continue; //2021-07 要求權重0也要印出

                    if (!student.Scores[Utilities.DomainToken].Contains(each.Domain))
                    {
                        student.Scores[Utilities.DomainToken].Add(each.Domain, dd, each.Credit.Value);

                        if (Perference.UserSelScoreType == "原始補考擇優" && each.ScoreMakeup.HasValue)
                            student.Scores[Utilities.DomainToken].AddReExam(each.Domain, each.ScoreMakeup.Value);
                    }
                }

                //運算後成績是在使用者按下列印時才計算。
                //因為需要依據使用者選擇的科目進行計算。
            }
        }

        /// <summary>
        /// Group By 科目清單，儲存於 Subjects  變數中。
        /// (限制在已選擇的學生有修的才算)。
        /// </summary>
        /// <param name="semsScores"></param>
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)]
        private void GroupBySubjects(List<JHSemesterScoreRecord> semsScores)
        {
            SubjectDomainMap = new Dictionary<string, string>();
            Dictionary<string, string> groupBySubject = new Dictionary<string, string>();

            //Group By 可選擇的科目清單。
            Dictionary<string, ReportStudent> dicSelectedStudents = JHStudent.SelectByClassIDs(
                SelectedClasses.ToKeys()).ToReportStudent().ToDictionary();

            foreach (JHSemesterScoreRecord eachScore in semsScores)
            {
                //如果成績不屬於已選擇的學生之中，就跳到下一筆。
                if (!dicSelectedStudents.ContainsKey(eachScore.RefStudentID)) continue;

                foreach (SubjectScore each in eachScore.Subjects.Values)
                {
                    if (!groupBySubject.ContainsKey(each.Subject))
                    {
                        groupBySubject.Add(each.Subject, null);
                        SubjectDomainMap.Add(each.Subject, each.Domain);
                    }
                }
            }

            //所有可選擇的科目。
            Subjects = new List<string>(groupBySubject.Keys);
            Subjects.Sort(SubjectOrder.CompareSubjectOrdinal);
        }

        /// <summary>
        /// 當使用者選的學年度學期變更時。
        /// </summary>
        private void Semester_SemesterChanged(object sender, EventArgs e)
        {
            FillCurrentSemesterData();
        }

        private void lnConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (new ConfigForm().ShowDialog() == DialogResult.OK)
                Perference = new ReportPreference(); //重新 New 一次就會重新讀取設定。
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            MsgBox.Show("請注意: 本成績單包含所有學生學期成績及排名，僅供校內教師參考使用，請勿公佈或發放與學生。", "注意", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            List<string> selectedSubjects = lvSubject.GetSelectedItems();
            List<string> selectedSubjectsM = lvSubject.GetSelectedItemsM();
            List<string> selectedDomains = lvDomain.GetSelectedItems();
            List<string> selectedDomainsM = lvDomain.GetSelectedItemsM();

            #region 檢查選擇的科目、領域是否合理。
            //if ((selectedSubjects.Count + selectedDomains.Count + Perference.PrintItems.Count) > Report.ScoreHeaderCount)
            //int ScoreHeaderCount = (Perference.PaperSize == "B4") ? 32 : 18;
            int ScoreHeaderCount = 18;

            if (Perference.PaperSize == "B4")
                ScoreHeaderCount = 32;

            if (Perference.PaperSize == "A3")
                ScoreHeaderCount = 38;
            
            if ((selectedSubjects.Count + selectedDomains.Count + Perference.PrintItems.Count) > ScoreHeaderCount)
            {
                MsgBox.Show("選擇的成績項目超過，無法列印報表。");
                return;
            }
            #endregion

            Perference.ReScoreMark = txtReMark.Text;
            Perference.UserSelScoreType = cbxScoreType.Text;
            Perference.Save();

            #region 重設計算資料
            foreach (ReportStudent each in AllStudents)
            {
                each.Scores[Utilities.SummaryToken].Clear();
                each.Places.Clear();
                each.Places.NS("班排名").Clear();
                each.Places.NS("年排名").Clear();
            }

            #endregion




            #region 計算各類成績。
            ScoreCalculatorP calculator = new ScoreCalculatorP(2, Utilities.SummaryToken);
            calculator.Subjects = selectedSubjects;
            calculator.Domains = selectedDomains;

            if (checkBoxX1.Checked)
            {
                //20210721 要求彈性課程領域不列入計算
                //2022-02-19 高雄專案[1110205]要求改為「只將八大領域列入計算」
                calculator.Subjects = selectedSubjectsM;
                calculator.Domains = selectedDomainsM;
            }

            foreach (ReportStudent each in AllStudents)
                calculator.CalculateScore(each);
            #endregion

            #region 計算排名。
            List<RatingScope<ReportStudent>> classScopes = AllStudents.ToClassScopes();
            List<RatingScope<ReportStudent>> gyScopes = AllStudents.ToGradeYearScopes();

            List<ScoreParser> parsers = new List<ScoreParser>(new ScoreParser[]{
                new ScoreParser("加權平均", Utilities.SummaryToken),
                new ScoreParser("加權總分", Utilities.SummaryToken),
                new ScoreParser("合計總分", Utilities.SummaryToken),
                new ScoreParser("算術平均", Utilities.SummaryToken),
            });

            foreach (ScoreParser parser in parsers)
            {
                foreach (RatingScope<ReportStudent> Scope in classScopes)
                    Scope.Rank(parser, PlaceOptions.Unsequence);

                foreach (RatingScope<ReportStudent> Scope in gyScopes)
                    Scope.Rank(parser, PlaceOptions.Unsequence);
            }

            #endregion

            #region 輸出到 Excel
            Report report = new Report();
            report.Perference = Perference;
            report.SchoolName = School.ChineseName;
            report.SchoolYear = Semester.SelectedSchoolYear.ToString();
            report.Semester = Semester.SelectedSemester.ToString();
            report.Classes = SelectedClasses;
            report.AllStudents = AllStudents.ToDictionary();
            report.SelectedSubject = selectedSubjects;
            report.SelectedDomain = selectedDomains;
            report.SubjectDomainMap = SubjectDomainMap;

            Utilities.DisableControls(this);
            BackgroundWorker OutputWork = new BackgroundWorker();
            OutputWork.DoWork += delegate (object sender1, DoWorkEventArgs e1)
            {
                e1.Result = report.Output();
            };
            OutputWork.RunWorkerCompleted += delegate (object sender1, RunWorkerCompletedEventArgs e1)
            {
                if (e1.Error != null)
                    MsgBox.Show(e1.Error.Message);
                else
                    Utilities.Save(e1.Result as Aspose.Cells.Workbook, "班級學期成績單.xls");

                Utilities.EnableControls(this);
            };
            OutputWork.RunWorkerAsync(report);

            #endregion
        }

        private bool Loading
        {
            set
            {
                LoadingSubject.Visible = value;
                LoadingDomain.Visible = value;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void cbxScoreType_SelectedIndexChanged(object sender, EventArgs e)
        {
            Loading = true;
            btnPrint.Enabled = false;
            // 存取使用成績類型
            Perference.UserSelScoreType = cbxScoreType.Text;
            Perference.Save();

            _tmpUserSelDomainList.Clear();
            _tmpUserSelSubjectList.Clear();

            // 記錄讀取前
            foreach (ListViewItem lvi in lvSubject.Items)
            {
                if (lvi.Checked)
                    _tmpUserSelSubjectList.Add(lvi.Text);
            }

            foreach (ListViewItem lvi in lvDomain.Items)
            {
                if (lvi.Checked)
                    _tmpUserSelDomainList.Add(lvi.Text);
            }

            if (MasterWorker.IsBusy)
                WorkPadding = true;
            else
                MasterWorker.RunWorkerAsync();
        }
    }
}
