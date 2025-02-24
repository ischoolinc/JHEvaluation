﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using JHSchool.Evaluation.Editor;
using Framework;
using System.Xml;
using DevComponents.DotNetBar.Controls;
using JHSchool.Evaluation.Calculation;
using JHSchool.Data;
using JHSchool;
using JHSchool.Evaluation;

namespace HsinChu.JHEvaluation.StudentExtendControls.SemesterScoreItemRelated
{
    public partial class QuickInputSemesterScoreForm : FISCA.Presentation.Controls.BaseForm
    {
        private List<JHSemesterScoreRecord> _semesterScoreRecordList;
        private JHStudentRecord _student;
        //private bool _gradeYearIsValid = false;
        private bool _semesterIsValid = false;
        private TextBoxManager _manager;
        private bool _filled = false;
        private ScoreCalculator _calculator;
        private SemesterHistoryUtility _util;

        private bool inputed = false;

        private PeriodCredit _learnDomainPC = new PeriodCredit();
        private PeriodCredit _elasticPC = new PeriodCredit();
        private decimal _learnDomainTotal = 0, _elasticTotal = 0;

        //努力程度對照表
        private Dictionary<decimal, int> _effortDict;
        private List<decimal> _effortScoreList;

        private QuickInputSemesterScoreForm()
        {
            InitializeComponent();
            InitializeComboBoxes();
            InitializeTextBoxManager();
            InitializeEffortDegreeMapping();
            InitializeDomainComboBox();
        }

        private void InitializeDomainComboBox()
        {
            chsDomain.Items.Clear();
            foreach (string domain in JHSchool.Evaluation.Subject.Domains)
                chsDomain.Items.Add(domain);
        }

        private void InitializeEffortDegreeMapping()
        {
            _effortDict = new Dictionary<decimal, int>();
            _effortScoreList = new List<decimal>();

            ConfigData cd = School.Configuration["努力程度對照表"];
            if (!string.IsNullOrEmpty(cd["xml"]))
            {
                XmlElement element = XmlHelper.LoadXml(cd["xml"]);

                foreach (XmlElement each in element.SelectNodes("Effort"))
                {
                    int code = int.Parse(each.GetAttribute("Code"));
                    decimal score;
                    if (!decimal.TryParse(each.GetAttribute("Score"), out score))
                        score = 0;

                    if (!_effortDict.ContainsKey(score))
                        _effortDict.Add(score, code);
                }

                _effortScoreList.AddRange(_effortDict.Keys);
                _effortScoreList.Sort(delegate (decimal a, decimal b)
                {
                    return b.CompareTo(a);
                });
            }
        }

        private void InitializeTextBoxManager()
        {
            _manager = new TextBoxManager(
                textBoxPC1, textBoxScore1, textBoxText1, // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。
                textBoxPC2, textBoxScore2, textBoxText2,
                textBoxPC3, textBoxScore3, textBoxText3,
                textBoxPC4, textBoxScore4, textBoxText4,
                textBoxPC5, textBoxScore5, textBoxText5,
                textBoxPC6, textBoxScore6, textBoxText6,
                textBoxPC7, textBoxScore7, textBoxText7,
                textBoxPC8, textBoxScore8, textBoxText8
                );
            _manager.AddEffortTextBoxMapping(textBoxScore1, textBoxEffort1); // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。
            _manager.AddEffortTextBoxMapping(textBoxScore2, textBoxEffort2);
            _manager.AddEffortTextBoxMapping(textBoxScore3, textBoxEffort3);
            _manager.AddEffortTextBoxMapping(textBoxScore4, textBoxEffort4);
            _manager.AddEffortTextBoxMapping(textBoxScore5, textBoxEffort5);
            _manager.AddEffortTextBoxMapping(textBoxScore6, textBoxEffort6);
            _manager.AddEffortTextBoxMapping(textBoxScore7, textBoxEffort7);
            _manager.AddEffortTextBoxMapping(textBoxScore8, textBoxEffort8);
        }

        private void InitializeComboBoxes()
        {
            try
            {
                cboSchoolYear.Enabled = cboSemester.Enabled = false;
                int schoolYear;
                if (!int.TryParse(School.DefaultSchoolYear, out schoolYear))
                    schoolYear = 97;

                for (int i = -2; i <= 2; i++)
                    cboSchoolYear.Items.Add(schoolYear + i);

                foreach (string item in new string[] { "1", "2" })
                    cboSemester.Items.Add(item);

                cboSchoolYear.Enabled = cboSemester.Enabled = true;
            }
            catch (Exception ex)
            {

            }
        }

        public QuickInputSemesterScoreForm(JHStudentRecord student)
            : this()
        {
            _student = student;
            JHScoreCalcRuleRecord record = student.ScoreCalcRule;

            _calculator = new ScoreCalculator(record);
            _util = new SemesterHistoryUtility(JHSemesterHistory.SelectByStudentID(student.ID));
            _semesterScoreRecordList = JHSemesterScore.SelectByStudentID(_student.ID);

            errorProvider.SetError(cboSchoolYear, "無效的學年度");
            errorProvider.SetError(cboSemester, "無效的學期");
        }

        private void SchoolYearAndSemester_TextChanged(object sender, EventArgs e)
        {
            if (ValidSemester())
            {
                _semesterIsValid = true;
                SemesterChanged(sender);
            }
            else
                _semesterIsValid = false;

            ValidSemesterAndGradeYear();
        }

        private void ValidSemesterAndGradeYear()
        {
            bool valid = _semesterIsValid;
            btnSave.Enabled = panelEx1.Enabled = valid;
            if (valid && !_filled)
                FillPeriodCredit();
        }

        private void FillPeriodCredit()
        {
            _filled = true;

            #region 取得領域總節數權數

            JHProgramPlanRecord record = _student.ProgramPlan;
            if (record != null)
            {
                Dictionary<string, PeriodCredit> pcs = new Dictionary<string, PeriodCredit>();

                foreach (K12.Data.ProgramSubject subject in record.Subjects)
                {
                    if ("" + subject.GradeYear == "" + _util.GetGradeYear(cboSchoolYear.Text, cboSemester.Text) && "" + subject.Semester == cboSemester.Text)
                    {
                        // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。
                        #region 各領域
                        if (!pcs.ContainsKey(subject.Domain))
                            pcs.Add(subject.Domain, new PeriodCredit(0, 0));
                        if (subject.Period.HasValue)
                            pcs[subject.Domain].Period += subject.Period.Value;
                        if (subject.Credit.HasValue)
                            pcs[subject.Domain].Credit += subject.Credit.Value;
                        #endregion

                    }
                }

                textBoxPC1.Text = pcs.ContainsKey("語文") ? pcs["語文"].ToString() : ""; // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。
                textBoxPC2.Text = pcs.ContainsKey("數學") ? pcs["數學"].ToString() : "";
                textBoxPC3.Text = pcs.ContainsKey("社會") ? pcs["社會"].ToString() : "";
                //textBoxPC4.Text = pcs.ContainsKey("藝術與人文") ? pcs["藝術與人文"].ToString() : "";
                //textBoxPC5.Text = pcs.ContainsKey("自然與生活科技") ? pcs["自然與生活科技"].ToString() : "";
                textBoxPC4.Text = pcs.ContainsKey("藝術") ? pcs["藝術"].ToString() : "";
                textBoxPC5.Text = pcs.ContainsKey("自然科學") ? pcs["自然科學"].ToString() : "";
                textBoxPC6.Text = pcs.ContainsKey("健康與體育") ? pcs["健康與體育"].ToString() : "";
                textBoxPC7.Text = pcs.ContainsKey("綜合活動") ? pcs["綜合活動"].ToString() : "";
                textBoxPC8.Text = pcs.ContainsKey("科技") ? pcs["科技"].ToString() : "";
            }
            #endregion
        }

        private void SemesterChanged(object sender)
        {
            int schoolYear = int.Parse(cboSchoolYear.Text);
            int semester = int.Parse(cboSemester.Text);

            bool found = false;

            foreach (JHSemesterScoreRecord record in _semesterScoreRecordList)
            {
                if (record.SchoolYear == schoolYear && record.Semester == semester)
                {
                    found = true;
                    break;
                }
            }

            _semesterIsValid = !found;

            if (found)
                toolTip.Show("此學年度學期已有資料", sender as Control, 10, 25);
            else
                toolTip.RemoveAll();
        }

        private bool ValidSemester()
        {
            bool valid = true;

            errorProvider.SetError(cboSchoolYear, "");
            errorProvider.SetError(cboSemester, "");

            int i;
            if (!int.TryParse(cboSchoolYear.Text, out i))
            {
                errorProvider.SetError(cboSchoolYear, "無效的學年度");
                valid = false;
            }
            if (!int.TryParse(cboSemester.Text, out i))
            {
                errorProvider.SetError(cboSemester, "無效的學期");
                valid = false;
            }

            return valid;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!inputed)
            {
                MsgBox.Show("尚未輸入成績");
                return;
            }

            if (!IsValid()) return;

            try
            {
                int schoolYear = int.Parse(cboSchoolYear.Text);
                int semester = int.Parse(cboSemester.Text);
                //int gradeYear = 0;

                //SemesterScoreRecordEditor editor = new SemesterScoreRecordEditor(_student, schoolYear, semester, gradeYear);
                JHSemesterScoreRecord newRecord = new JHSemesterScoreRecord();
                newRecord.RefStudentID = _student.ID;
                newRecord.SchoolYear = schoolYear;
                newRecord.Semester = semester;

                // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。
                if (CheckDomainValid(textBoxPC1, textBoxScore1, textBoxEffort1))
                    newRecord.Domains.Add("語文", GetDomainScore("語文", textBoxPC1, textBoxScore1, textBoxEffort1, textBoxText1));
                if (CheckDomainValid(textBoxPC2, textBoxScore2, textBoxEffort2))
                    newRecord.Domains.Add("數學", GetDomainScore("數學", textBoxPC2, textBoxScore2, textBoxEffort2, textBoxText2));
                if (CheckDomainValid(textBoxPC3, textBoxScore3, textBoxEffort3))
                    newRecord.Domains.Add("社會", GetDomainScore("社會", textBoxPC3, textBoxScore3, textBoxEffort3, textBoxText3));
                //if (CheckDomainValid(textBoxPC4, textBoxScore4, textBoxEffort4))
                //    newRecord.Domains.Add("藝術與人文", GetDomainScore("藝術與人文", textBoxPC4, textBoxScore4, textBoxEffort4, textBoxText4));
                //if (CheckDomainValid(textBoxPC5, textBoxScore5, textBoxEffort5))
                //    newRecord.Domains.Add("自然與生活科技", GetDomainScore("自然與生活科技", textBoxPC5, textBoxScore5, textBoxEffort5, textBoxText5));
                if (CheckDomainValid(textBoxPC4, textBoxScore4, textBoxEffort4))
                    newRecord.Domains.Add("藝術", GetDomainScore("藝術", textBoxPC4, textBoxScore4, textBoxEffort4, textBoxText4));
                if (CheckDomainValid(textBoxPC5, textBoxScore5, textBoxEffort5))
                    newRecord.Domains.Add("自然科學", GetDomainScore("自然科學", textBoxPC5, textBoxScore5, textBoxEffort5, textBoxText5));
                if (CheckDomainValid(textBoxPC6, textBoxScore6, textBoxEffort6))
                    newRecord.Domains.Add("健康與體育", GetDomainScore("健康與體育", textBoxPC6, textBoxScore6, textBoxEffort6, textBoxText6));
                if (CheckDomainValid(textBoxPC7, textBoxScore7, textBoxEffort7))
                    newRecord.Domains.Add("綜合活動", GetDomainScore("綜合活動", textBoxPC7, textBoxScore7, textBoxEffort7, textBoxText7));
                if (CheckDomainValid(textBoxPC8, textBoxScore8, textBoxEffort8))
                    newRecord.Domains.Add("科技", GetDomainScore("科技", textBoxPC8, textBoxScore8, textBoxEffort8, textBoxText8));

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    PeriodCredit pc = new PeriodCredit();
                    pc.Parse("" + row.Cells[chsPeriodCredit.Index].Value);
                    K12.Data.SubjectScore subject = new K12.Data.SubjectScore();
                    subject.Domain = "" + row.Cells[chsDomain.Index].Value;
                    subject.Subject = "" + row.Cells[chsSubject.Index].Value;
                    subject.Period = pc.Period;
                    subject.Credit = pc.Credit;
                    subject.Score = decimal.Parse("" + row.Cells[chsScore.Index].Value);
                    // 2018.09.22 [ischoolKingdom] Vicky 依據 [J成績][HC][03]快速新增學期成績修正 項目，使輸入成績一併儲存至原始成績。
                    subject.ScoreOrigin = decimal.Parse("" + row.Cells[chsScore.Index].Value);
                    //subject.Effort = int.Parse("" + row.Cells[chsEffort.Index].Value);
                    subject.Effort = 1;
                    subject.Text = "" + row.Cells[chsText.Index].Value;

                    newRecord.Subjects.Add(subject.Subject, subject);
                }

                if (!string.IsNullOrEmpty(textBoxX33.Text))
                    newRecord.Domains.Add("彈性課程", GetElasticDomain());
                if (!string.IsNullOrEmpty(textBoxX34.Text))
                    newRecord.LearnDomainScore = decimal.Parse(textBoxX34.Text);
                if (!string.IsNullOrEmpty(textBoxX35.Text))
                    newRecord.CourseLearnScore = decimal.Parse(textBoxX35.Text);

                JHSemesterScore.Insert(newRecord);
                SaveLog(newRecord);
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存失敗");
                this.DialogResult = DialogResult.Cancel;
                this.Close();
                return;
            }

            this.DialogResult = DialogResult.OK;
        }

        private void SaveLog(JHSemesterScoreRecord newRecord)
        {
            string semester = string.Format("{0}學年度 第{1}學期", newRecord.SchoolYear, newRecord.Semester);
            StringBuilder builder = new StringBuilder("");
            builder.Append(GetStudentInfo(_student));
            builder.Append("，新增「" + semester + "」的學期成績");

            FISCA.LogAgent.ApplicationLog.Log("成績系統.學期成績", "新增學期成績", "student", _student.ID, builder.ToString());
            //307(16號) 5550082 謝化挺，新增「97學年度 第1學期」的學期成績
        }

        private string GetStudentInfo(JHStudentRecord student)
        {
            return HsinChu.JHEvaluation.Utility.StudentInfoConvertor.GetInfoWithClass(student);
        }

        private bool CheckDomainValid(TextBoxX pc, TextBoxX score, TextBoxX effort)
        {
            bool boolPc = !string.IsNullOrEmpty(pc.Text);
            bool boolScore = !string.IsNullOrEmpty(score.Text);
            //bool boolEffort = !string.IsNullOrEmpty(effort.Text);

            if (boolPc && boolScore)// && boolEffort)
                return true;
            else
                return false;
        }

        private K12.Data.DomainScore GetElasticDomain()
        {
            K12.Data.DomainScore domain = new K12.Data.DomainScore();
            domain.Domain = "彈性課程";
            domain.Score = decimal.Parse(textBoxX33.Text);
            domain.ScoreOrigin = decimal.Parse(textBoxX33.Text); // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，輸入成績一併存入原始成績。

            PeriodCredit pc = new PeriodCredit();
            PeriodCredit temp = new PeriodCredit();
            //int pc = 0;
            //int effort = 0;
            int count = 0;
            string text = "";

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                if ("" + row.Cells[chsDomain.Index].Value == "彈性課程"|| "" + row.Cells[chsDomain.Index].Value == "")
                {
                    count++;
                    temp.Parse("" + row.Cells[chsPeriodCredit.Index].Value);
                    pc.Credit += temp.Credit;
                    pc.Period += temp.Period;
                    //effort += int.Parse("" + row.Cells[chsEffort.Index].Value);
                    if (!string.IsNullOrEmpty("" + row.Cells[chsText.Index].Value))
                    {
                        text += "" + row.Cells[chsSubject.Index].Value + ":" + row.Cells[chsText.Index].Value;
                    }
                }
            }

            domain.Period = pc.Period;
            domain.Credit = pc.Credit;
            //domain.Effort = (int)decimal.Round((decimal)effort / (decimal)count, 0, MidpointRounding.AwayFromZero);
            domain.Effort = 1;
            domain.Text = text;
            return domain;
        }

        private K12.Data.DomainScore GetDomainScore(string p, TextBoxX pc, TextBoxX score, TextBoxX effort, TextBoxX text)
        {
            PeriodCredit temp = new PeriodCredit();
            temp.Parse(pc.Text);
            K12.Data.DomainScore domain = new K12.Data.DomainScore();
            domain.Domain = p;
            domain.Period = temp.Period;
            domain.Credit = temp.Credit;
            //domain.Effort = int.Parse(effort.Text);
            domain.Effort = 1;
            domain.Score = decimal.Parse(score.Text);
            domain.ScoreOrigin = decimal.Parse(score.Text); // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。
            domain.Text = text.Text;
            return domain;
        }

        private bool IsValid()
        {
            bool valid = true;

            if (errorProvider.HasError)
                valid = false;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                    {
                        valid = false;
                        break;
                    }
                }
            }

            return valid;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 計算學習領域成績
        /// </summary>       
        private void CalculateLearnDomainScore()
        {
            PeriodCredit pc = new PeriodCredit();
            //int pc = 0;
            decimal total = 0;

            EvalDomainScore(ref pc, ref total, textBoxScore1, textBoxPC1); // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。
            EvalDomainScore(ref pc, ref total, textBoxScore2, textBoxPC2);
            EvalDomainScore(ref pc, ref total, textBoxScore3, textBoxPC3);
            EvalDomainScore(ref pc, ref total, textBoxScore4, textBoxPC4);
            EvalDomainScore(ref pc, ref total, textBoxScore5, textBoxPC5);
            EvalDomainScore(ref pc, ref total, textBoxScore6, textBoxPC6);
            EvalDomainScore(ref pc, ref total, textBoxScore7, textBoxPC7);
            EvalDomainScore(ref pc, ref total, textBoxScore8, textBoxPC8);

            if (pc.Credit > 0)
            {
                textBoxX34.Text = "" + _calculator.ParseLearnDomainScore(total / pc.Credit);
                _learnDomainPC = pc;
                _learnDomainTotal = total;
            }
        }

        private void EvalDomainScore(ref PeriodCredit pc, ref decimal total, TextBox scoreTextBox, TextBox pcTextBox)
        {
            if (ValidTextBox(scoreTextBox) && ValidTextBox(pcTextBox))
            {
                PeriodCredit temp = new PeriodCredit();
                temp.Parse(pcTextBox.Text);
                total += decimal.Parse(scoreTextBox.Text) * temp.Credit;
                pc.Credit += temp.Credit;
                pc.Period += temp.Period;
            }
        }

        private void textBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter || e.KeyData == Keys.Down)
            {
                Control control = sender as Control;
                Control next = null;
                do
                {
                    next = _manager.GetNextControl4D(control);
                    control = next;
                }
                while (next != null && next.Enabled == false);
                if (next != null)
                    next.Focus();
            }

            if (e.KeyData == Keys.Up)
            {
                Control control = sender as Control;
                Control next = null;
                do
                {
                    next = _manager.GetNextControl4U(control);
                    control = next;
                }
                while (next != null && next.Enabled == false);
                if (next != null)
                    next.Focus();
            }


            if (e.KeyData == Keys.Right)
            {
                Control control = sender as Control;
                Control next = null;
                do
                {
                    next = _manager.GetNextControl1R(control);
                    control = next;
                }
                while (next != null && next.Enabled == false);
                if (next != null)
                    next.Focus();
            }


            if (e.KeyData == Keys.Left)
            {
                Control control = sender as Control;
                Control next = null;
                do
                {
                    next = _manager.GetNextControl1L(control);
                    control = next;
                }
                while (next != null && next.Enabled == false);
                if (next != null)
                    next.Focus();
            }

        }

        private void scoreTextBox_TextChanged(object sender, EventArgs e)
        {
            inputed = true;

            TextBox txt = sender as TextBox;
            if (!ValidDecimalTextBox(txt)) return;

            #region 不及格變紅色
            txt.ForeColor = Color.Black;
            decimal d;
            if (decimal.TryParse(txt.Text, out d))
            {
                if (d < 60)
                    txt.ForeColor = Color.Red;

                if (d > 100 || d < 0)
                    txt.ForeColor = Color.Green;
            }
            #endregion

            ProcessEffortCode(txt);
            CalculateLearnDomainScore();
        }

        private void pcTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (!ValidPCTextBox(txt)) return;

            CalculateLearnDomainScore();
        }

        private void effortTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (!ValidIntTextBox(txt)) return;
        }

        private void ProcessEffortCode(TextBox txt)
        {
            if (!string.IsNullOrEmpty(txt.Text))
            {
                decimal d;
                if (decimal.TryParse(txt.Text, out d))
                {
                    TextBox effortCtrl = (TextBox)_manager.GetEffortTextBox(txt);
                    if (effortCtrl != null) effortCtrl.Text = "" + GenerateEffortCode(d);
                }
            }
        }

        private void literatureScoreTextBox_TextChanged(object sender, EventArgs e)
        {
            inputed = true;

            TextBox txt = sender as TextBox;
            if (ValidDecimalTextBox(txt))
            {
                #region 不及格變紅色
                txt.ForeColor = Color.Black;
                decimal d;
                if (decimal.TryParse(txt.Text, out d))
                {
                    if (d < 60)
                        txt.ForeColor = Color.Red;

                    if (d > 100 || d < 0)
                        txt.ForeColor = Color.Green;
                }
                #endregion

                ProcessEffortCode(txt);
                CalculateLearnDomainScore();
            }
        }

        private void literaturePCTextBox_TextChanged(object sender, EventArgs e)
        {
            if (ValidPCTextBox(sender as TextBox))
            {
                CalculateLearnDomainScore();
            }
        }

        private int GenerateEffortCode(decimal d)
        {
            foreach (decimal score in _effortScoreList)
                if (d >= score) return _effortDict[score];

            return _effortDict[_effortScoreList[_effortScoreList.Count - 1]];
        }

        private bool ValidPCTextBox(TextBox tb)
        {
            bool valid = true;

            errorProvider.SetError(tb, "");
            PeriodCredit pc = new PeriodCredit();
            if (!pc.Parse(tb.Text))
            {
                errorProvider.SetIconPadding(tb, -17);
                errorProvider.SetError(tb, pc.Error);
                valid = false;
            }

            return valid;
        }

        private bool ValidIntTextBox(TextBox txtBox)
        {
            bool valid = true;

            errorProvider.SetError(txtBox, "");
            if (!string.IsNullOrEmpty(txtBox.Text))
            {
                int i;
                if (!int.TryParse(txtBox.Text, out i))
                {
                    errorProvider.SetIconPadding(txtBox, -17);
                    errorProvider.SetError(txtBox, "必需為整數");
                    valid = false;
                }
            }

            return valid;
        }

        private bool ValidDecimalTextBox(TextBox txtBox)
        {
            bool valid = true;

            errorProvider.SetError(txtBox, "");
            if (!string.IsNullOrEmpty(txtBox.Text))
            {
                decimal d;
                if (!decimal.TryParse(txtBox.Text, out d))
                {
                    errorProvider.SetIconPadding(txtBox, -17);
                    errorProvider.SetError(txtBox, "必需為數字");
                    valid = false;
                }
            }

            return valid;
        }

        // 2018.09.22 [ischoolKingdom] Vicky依據 [J成績][HC][03] 快速新增學期成績修正，將語文領域的子領域兩個項目欄位拿掉，僅保留輸入語文領域成績的單一功能。

        private void cboSchoolYear_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                cboSemester.Focus();
            }
        }

        private void cboSemester_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                if (textBoxScore1.Enabled)
                    textBoxScore1.Focus();

                else
                    textBoxScore2.Focus();
            }
        }

        private bool ValidTextBox(TextBox txt)
        {
            if (string.IsNullOrEmpty(errorProvider.GetError(txt)) &&
                !string.IsNullOrEmpty(txt.Text) &&
                txt.Enabled)
                return true;
            return false;
        }

        private void dgv_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            foreach (DataGridViewRow row in dgv.Rows)
                ValidRow(row);
        }

        private void ValidRow(DataGridViewRow row)
        {
            if (row.IsNewRow) return;

            dgv.EndEdit();

            bool inputted = false;
            List<DataGridViewColumn> cols = new List<DataGridViewColumn>(new DataGridViewColumn[] { chsDomain, chsSubject, chsPeriodCredit, chsScore, chsEffort });
            foreach (DataGridViewColumn col in cols)
            {
                if (!string.IsNullOrEmpty("" + row.Cells[col.Index].Value))
                {
                    inputted = true;
                    break;
                }
            }

            if (inputted)
            {
                foreach (DataGridViewColumn col in cols)
                {
                    if (col.HeaderText != "領域")
                    {
                        DataGridViewCell cell = row.Cells[col.Index];
                        if (string.IsNullOrEmpty("" + cell.Value))
                            cell.ErrorText = "不可為空白";
                        else
                            cell.ErrorText = "";
                    }
                }
            }

            dgv.BeginEdit(false);
        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];

            if (cell.OwningRow.IsNewRow) return;

            #region 證驗科目
            if (cell.OwningColumn == chsSubject)
            {
                cell.ErrorText = "";
                if (string.IsNullOrEmpty("" + cell.Value))
                    cell.ErrorText = "科目不能為空白";
            }
            #endregion

            #region 驗證成績
            if (cell.OwningColumn == chsScore)
            {
                cell.ErrorText = "";
                if (!string.IsNullOrEmpty("" + cell.Value))
                {
                    decimal d;
                    if (!decimal.TryParse("" + cell.Value, out d))
                        cell.ErrorText = "成績必須為數字";
                    else
                    {
                        cell.ErrorText = "";

                        #region 不及格變紅色
                        if (d < 60)
                            cell.Style.ForeColor = Color.Red;
                        else
                            cell.Style.ForeColor = Color.Black;
                        #endregion

                        if (d > 100 || d < 0)
                            cell.Style.ForeColor = Color.Green;

                        #region 自動換算出努力程度
                        if (_effortScoreList.Count > 0)
                        {
                            cell.OwningRow.Cells[chsEffort.Index].Value = GenerateEffortCode(d);
                            cell.OwningRow.Cells[chsEffort.Index].ErrorText = "";
                        }
                        #endregion
                    }
                }
                else
                    cell.ErrorText = "成績不能為空白";
            }
            #endregion

            #region 驗證努力程度
            if (cell.OwningColumn == chsEffort)
            {
                cell.ErrorText = "";
                if (!string.IsNullOrEmpty("" + cell.Value))
                {
                    int i;
                    if (!int.TryParse("" + cell.Value, out i))
                        cell.ErrorText = "努力程度必須為整數";
                }
                else
                    cell.ErrorText = "努力程度不能為空白";
            }
            #endregion

            #region 驗證節數權數
            if (cell.OwningColumn == chsPeriodCredit)
            {
                cell.ErrorText = "";
                PeriodCredit pc = new PeriodCredit();
                if (!pc.Parse("" + cell.Value))
                {
                    cell.ErrorText = pc.Error;
                }
            }
            #endregion

            //if (!HasDataGridViewError())
            //{
                CalculateCourseLearnScore();
            //}
        }

        private void CalculateCourseLearnScore()
        {
            PeriodCredit pc = new PeriodCredit();
            //int pc = 0;
            decimal total = 0;
            decimal trydecimal;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if ("" + row.Cells[chsDomain.Index].Value == "彈性課程" || "" + row.Cells[chsDomain.Index].Value == "")
                {
                    if (decimal.TryParse("" + row.Cells[chsPeriodCredit.Index].Value, out trydecimal) && decimal.TryParse("" + row.Cells[chsScore.Index].Value, out trydecimal))
                    {
                        PeriodCredit temp = new PeriodCredit();
                        temp.Parse("" + row.Cells[chsPeriodCredit.Index].Value);
                        pc.Credit += temp.Credit;
                        pc.Period += temp.Period;
                        total += temp.Credit * decimal.Parse("" + row.Cells[chsScore.Index].Value);
                    }
                }
            }

            if (pc.Credit > 0)
            {
                textBoxX33.Text = "" + _calculator.ParseLearnDomainScore(total / pc.Credit);
                _elasticPC = pc;
                _elasticTotal = total;
                if (_learnDomainPC.Credit > 0)
                    textBoxX35.Text = "" + _calculator.ParseLearnDomainScore((_learnDomainTotal + total) / (_learnDomainPC + pc).Credit);
            }
        }

        private bool HasDataGridViewError()
        {
            bool error = false;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                    {
                        error = true;
                        break;
                    }
                }
            }

            return error;
        }

        private void textBoxX33_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            errorProvider.SetError(txt, "");
            if (!string.IsNullOrEmpty(txt.Text))
            {
                decimal d;
                if (!decimal.TryParse(txt.Text, out d))
                    errorProvider.SetError(txt, "成績必須為數字");
            }
        }
    }

    internal class TextBoxManager
    {
        private List<Control> _controls;
        private Dictionary<Control, Control> _effortCtrls;

        public TextBoxManager(params Control[] controls)
        {
            _controls = new List<Control>();
            foreach (Control c in controls)
                _controls.Add(c);
        }

        public void AddEffortTextBoxMapping(Control scoreCtrl, Control effortCtrl)
        {
            if (_effortCtrls == null) _effortCtrls = new Dictionary<Control, Control>();
            if (!_effortCtrls.ContainsKey(scoreCtrl))
                _effortCtrls.Add(scoreCtrl, effortCtrl);
            _effortCtrls[scoreCtrl] = effortCtrl;
        }

        public Control GetEffortTextBox(Control scoreCtrl)
        {
            if (_effortCtrls.ContainsKey(scoreCtrl))
                return _effortCtrls[scoreCtrl];
            return null;
        }

        //public Control GetNextControl(Control control)
        //{
        //    int index = _controls.IndexOf(control);
        //    if (index >= 0 && index + 1 < _controls.Count)
        //        return _controls[index + 1];
        //    return null;
        //}
        // 往下四個
        public Control GetNextControl4D(Control control)
        {
            int index = _controls.IndexOf(control);
            if (index >= 0 && index + 3 < _controls.Count)
                if ((index + 3) <= _controls.Count)
                    return _controls[index + 3];
                else
                    return _controls[0];

            return null;
        }

        // 往上四個
        public Control GetNextControl4U(Control control)
        {
            int index = _controls.IndexOf(control);
            if (index >= 0 && index - 3 < _controls.Count)
                if ((index - 3) > 0)
                    return _controls[index - 3];
                else
                    return _controls[0];

            return null;
        }

        // 往右1個
        public Control GetNextControl1R(Control control)
        {
            int index = _controls.IndexOf(control);
            if (index >= 0 && index + 1 < _controls.Count)
                if ((index + 1) <= _controls.Count)
                    return _controls[index + 1];
                else
                    return _controls[0];

            return null;
        }

        // 往左1個
        public Control GetNextControl1L(Control control)
        {
            int index = _controls.IndexOf(control);
            if (index >= 0 && index - 1 < _controls.Count)
                if ((index - 1) > 0)
                    return _controls[index - 1];
                else
                    return _controls[0];
            return null;
        }
    }
}