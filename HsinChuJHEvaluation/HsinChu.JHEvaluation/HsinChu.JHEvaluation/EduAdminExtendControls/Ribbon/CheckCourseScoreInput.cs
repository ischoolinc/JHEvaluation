using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using FISCA.Presentation.Controls;
using JHSchool;
using JHSchool.Evaluation;
using HsinChu.JHEvaluation.Data;
using JHSchool.Data;
using K12.Data.Configuration;
using FISCA.Data;
using System.IO;

namespace HsinChu.JHEvaluation.EduAdminExtendControls.Ribbon
{
    //CheckCourseScoreInput
    public partial class CheckCourseScoreInput : BaseForm
    {
        private List<string> _CourseIDsList;
        private List<string> _StudentIDsList;
        private List<string> _SCAttendIDsList;
        private List<string> _ExamList;
        private List<string> _SceTakeIDsList;
        /// <summary>
        /// CourseID,TeacherName
        /// </summary>
        private Dictionary<string, string> _TeacherNameDic;
        private List<K12.Data.ExamRecord> _exams = new List<K12.Data.ExamRecord>();
        private List<JHSCAttendRecord> _scAttendRecordList;
        private string _ExamName = "";
        private string _ExamID = "";
        private int _SchoolYear = 0;
        private int _Semester = 0;
        // Log
        PermRecLogProcess prlp;

        //記錄 StudentID 與 DataGridViewRow 的對應
        private Dictionary<string, DataGridViewRow> _studentRowDict;

        //文字評量代碼表
        //private Dictionary<string, string> _textMapping = new Dictionary<string, string>();

        private List<DataGridViewCell> _dirtyCellList;

        /// <summary>
        /// Constructor
        /// 傳入一個課程。
        /// </summary>
        /// <param name="course"></param>
        public CheckCourseScoreInput()
        {
            InitializeComponent();
            prlp = new PermRecLogProcess();

            #region 取得文字評量代碼表
            //K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["文字描述代碼表"];
            //if (!string.IsNullOrEmpty(cd["xml"]))
            //{
            //    K12.Data.XmlHelper helper = new K12.Data.XmlHelper(K12.Data.XmlHelper.LoadXml(cd["xml"]));
            //    foreach (XmlElement item in helper.GetElements("Item"))
            //    {
            //        string code = item.GetAttribute("Code");
            //        string content = item.GetAttribute("Content");

            //        if (!_textMapping.ContainsKey(code))
            //            _textMapping.Add(code, content);
            //    }
            //}
            #endregion

            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dgv);
        }

        /// <summary>
        /// Form_Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckCourseScoreInput_Load(object sender, EventArgs e)
        {
            int defaultSchoolYear, defaultSemester;

            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out defaultSchoolYear))
                iptSchoolYear.Value = _SchoolYear = defaultSchoolYear;
            if (int.TryParse(K12.Data.School.DefaultSemester, out defaultSemester))
                iptSemester.Value = _Semester = defaultSemester;

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

            #region 取得評量設定

            FillToComboBox();

            // 當沒有試別關閉
            if (cboExamList.Items.Count < 1)
                this.Close();
            _ExamName = cboExamList.Text;

            foreach (K12.Data.ExamRecord ex in _exams)
            {
                if (ex.Name == cboExamList.Text)
                {
                    _ExamID = ex.ID;
                    break;
                }
            }
            #endregion
        }

        /// <summary>
        /// 處理分數顏色變化
        /// </summary>
        private void LoadDvScoreColor()
        {
            // 處理初始分數變色
            foreach (DataGridViewRow dgv1 in dgv.Rows)
                foreach (DataGridViewCell cell in dgv1.Cells)
                {
                    cell.ErrorText = "";

                    if (cell.OwningColumn == chInputScore || cell.OwningColumn == chInputAssignmentScore)
                    {
                        cell.Style.ForeColor = Color.Black;
                        if (!string.IsNullOrEmpty("" + cell.Value))
                        {
                            decimal d;
                            if (!decimal.TryParse("" + cell.Value, out d))
                            {
                                if (PluginMain.ScoreTextMap.Keys.Count > 0)
                                {
                                    if (!PluginMain.ScoreTextMap.ContainsKey("" + cell.Value))
                                        cell.ErrorText = "分數必須為數字或「" + string.Join("、", PluginMain.ScoreTextMap.Keys) + "」";
                                    else
                                    {
                                        if (PluginMain.ScoreTextMap["" + cell.Value].AllowCalculation)
                                            cell.Style.ForeColor = Color.Purple;  //要計算(無論分數)，顯示紫色
                                        if (PluginMain.ScoreTextMap["" + cell.Value].Score == 0 && PluginMain.ScoreTextMap["" + cell.Value].AllowCalculation)
                                            cell.Style.ForeColor = Color.Red;  //要計算且依照0分計算，顯示紅色
                                        if (!PluginMain.ScoreTextMap["" + cell.Value].AllowCalculation)
                                            cell.Style.ForeColor = Color.Blue;//不要計算，顯示藍色
                                    }
                                }
                                else
                                {
                                    cell.ErrorText = "分數必須為數字";
                                }
                            }
                            else
                            {
                                if (d < 60)
                                    cell.Style.ForeColor = Color.Red;
                                if (d > 100 || d < 0)
                                    cell.Style.ForeColor = Color.Green;
                            }
                        }
                    }
                }
        }

        /// <summary>
        /// 將學生填入DataGridView。
        /// </summary>
        private void FillStudentsToDataGridView()
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在尋找修課學生");
            #region 取得修課學生
            _dirtyCellList = new List<DataGridViewCell>();
            _CourseIDsList = new List<string>();
            _StudentIDsList = new List<string>();
            _SCAttendIDsList = new List<string>();
            _ExamList = new List<string>();
            _SceTakeIDsList = new List<string>();
            _TeacherNameDic = new Dictionary<string, string>();
            _studentRowDict = new Dictionary<string, DataGridViewRow>();
            QueryHelper queryHelper = new QueryHelper();
            string queryCourseIDs = @"WITH student_score AS
(
    SELECT 
        student.id AS student_id
		, class.id AS classID
		, class.class_name
		, class.grade_year
		, class.display_order
        ,student.student_number
        ,student.name
        ,student.seat_no
        ,school_year
        ,semester
        ,sc_attend.id AS sc_attend_id
        , sce_take.id AS sce_take_id
		, course.id AS courseID
        ,course.course_name
        --,course.subject
        --,course.credit
        ,exam_name
        ,  teacher_name
        , (array_to_string(xpath('//Extension/Score/text()', xmlparse(content sce_take.extension)), '')::text) as Score
		, (array_to_string(xpath('//Extension/AssignmentScore/text()', xmlparse(content sce_take.extension)), '')::text) as AssignmentScore
        , te_include.ref_exam_id       
    FROM student 
	LEFT JOIN class 
	ON student.ref_class_id=class.id
    LEFT JOIN sc_attend 
        ON sc_attend.ref_student_id=student.id
    INNER JOIN ( SELECT * FROM course WHERE school_year ={0} AND semester={1} ) AS course 
        ON sc_attend.ref_course_id=course.id 
	INNER JOIN te_include 
        ON course.ref_exam_template_id = te_include.ref_exam_template_id
    LEFT JOIN   (SELECT * FROM  sce_take WHERE  ref_exam_id={2} )   AS  sce_take 
        ON  sce_take.ref_sc_attend_id=sc_attend.id
    LEFT JOIN exam
        ON exam.id=sce_take.ref_exam_id 
	LEFT JOIN 
    tc_instruct ON tc_instruct.ref_course_id = course.id AND sequence = 1
	LEFT JOIN 
    teacher ON teacher.id = tc_instruct.ref_teacher_id
    WHERE  te_include.ref_exam_id ={2}
)
SELECT * FROM student_score
WHERE
score ='' or AssignmentScore='' 
or score='-2147483648' or score ='-2147483647' 
or AssignmentScore='-2147483648' or AssignmentScore ='-2147483647'
or score ='0' or AssignmentScore='0' 
or sce_take_id IS NULL
ORDER BY courseID, grade_year, display_order, class_name, seat_no, student_number
";
            queryCourseIDs = string.Format(queryCourseIDs, _SchoolYear, _Semester, _ExamID);
            try
            {
                DataTable dt = queryHelper.Select(queryCourseIDs);
                foreach (DataRow dr in dt.Rows)
                {
                    string courseID = dr["courseID"].ToString();
                    string teacher_name = dr["teacher_name"].ToString();
                    string studentID = dr["student_id"].ToString();
                    string sc_attend_id = dr["sc_attend_id"].ToString();
                    string sce_take_id = dr["sce_take_id"].ToString();
                    string exam_ID = dr["ref_exam_id"].ToString();
                    if (!_TeacherNameDic.ContainsKey(courseID))
                        _TeacherNameDic.Add(courseID, teacher_name);
                    if (!_CourseIDsList.Contains(courseID))
                        _CourseIDsList.Add(courseID);
                    if (!_StudentIDsList.Contains(studentID))
                        _StudentIDsList.Add(studentID);
                    if (!_SCAttendIDsList.Contains(sc_attend_id))
                        _SCAttendIDsList.Add(sc_attend_id);
                    if (!_ExamList.Contains(exam_ID))
                        _ExamList.Add(exam_ID);
                    if (!_SceTakeIDsList.Contains(sce_take_id))
                        _SceTakeIDsList.Add(sce_take_id);

                }

                if (_StudentIDsList.Count > 0)
                    _scAttendRecordList = (JHSCAttend.Select(_StudentIDsList, _CourseIDsList, _SCAttendIDsList, _SchoolYear.ToString(), _Semester.ToString()));
                else
                {
                    MsgBox.Show("查無資料。");
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("查無資料");
                    dgv.Rows.Clear();
                    return;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("尋找修課學生發生錯誤："+ex.Message);
            }


            #endregion
            _studentRowDict.Clear();
            if (_scAttendRecordList == null) return;

            dgv.SuspendLayout();
            dgv.Rows.Clear();

            _scAttendRecordList.Sort(SCAttendComparer);
            foreach (var record in _scAttendRecordList)
            {
                JHStudentRecord student = record.Student;
                if (student.StatusStr != "一般") continue;

                DataGridViewRow row = new DataGridViewRow();
                string teacherName = "";
                if (_TeacherNameDic.ContainsKey(record.RefCourseID))
                    teacherName = _TeacherNameDic[record.RefCourseID];
                row.CreateCells(dgv, record.Course.Name, teacherName,
                    (student.Class != null) ? student.Class.Name : "",
                    student.SeatNo,
                    student.Name,
                    student.StudentNumber
                );
                dgv.Rows.Add(row);

                SCAttendTag tag = new SCAttendTag();
                tag.SCAttend = record;
                row.Tag = tag;

                //加入 StudentID 與 Row 的對應
                if (!_studentRowDict.ContainsKey(student.ID))
                    _studentRowDict.Add(student.ID, row);

            }
            dgv.ResumeLayout();

            GetScoresAndFill();

            LoadDvScoreColor();
            dgv.ResumeLayout();
            SetLoadDataToLog();
        }

        /// <summary>
        /// Comparision
        /// 依課程名稱、班級、座號、學號排序修課學生。
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        private int SCAttendComparer(JHSCAttendRecord a, JHSCAttendRecord b)
        {
            if (a.RefCourseID == b.RefCourseID)
            {
                //開始比
                if (a.Student.Class != null && b.Student.Class != null)
                {
                    if (a.Student.Class.ID == b.Student.Class.ID)
                    {
                        if (a.Student.SeatNo.HasValue && b.Student.SeatNo.HasValue)
                        {
                            if (a.Student.SeatNo.Value == b.Student.SeatNo.Value)
                                return a.Student.StudentNumber.CompareTo(b.Student.StudentNumber);
                            return a.Student.SeatNo.Value.CompareTo(b.Student.SeatNo.Value);
                        }
                        else if (a.Student.SeatNo.HasValue)
                            return -1;
                        else if (b.Student.SeatNo.HasValue)
                            return 1;
                        else
                            return a.Student.StudentNumber.CompareTo(b.Student.StudentNumber);
                    }
                    else
                        return a.Student.Class.Name.CompareTo(b.Student.Class.Name);
                }
                else if (a.Student.Class != null && b.Student.Class == null)
                    return -1;
                else if (a.Student.Class == null && b.Student.Class != null)
                    return 1;
                else
                    return a.Student.StudentNumber.CompareTo(b.Student.StudentNumber);
            }
            else
                return a.RefCourseID.CompareTo(b.RefCourseID);
        }


        /// <summary>
        /// 將試別填入ComboBox。
        /// </summary>
        private void FillToComboBox()
        {
            cboExamList.Items.Clear();
            _exams.Clear();
            _exams = K12.Data.Exam.SelectAll();


            foreach (K12.Data.ExamRecord exName in _exams)
            {
                cboExamList.Items.Add(exName.Name);
            }
            cboExamList.SelectedIndex = 0;
        }

        /// <summary>
        /// 載入資料 Log
        /// </summary>
        public void SetLoadDataToLog()
        {
            try
            {   // 將暫存清空
                prlp.ClearCache();
                _ExamName = cboExamList.Text;

                foreach (DataGridViewRow dgvr in dgv.Rows)
                {
                    string strClassName = string.Empty, strSeatNo = string.Empty, strName = string.Empty, strStudentNumber = string.Empty;
                    if (dgvr.IsNewRow)
                        continue;

                    if (dgvr.Cells[chClassName.Index].Value != null)
                        strClassName = dgvr.Cells[chClassName.Index].Value.ToString();

                    if (dgvr.Cells[chSeatNo.Index].Value != null)
                        strSeatNo = dgvr.Cells[chSeatNo.Index].Value.ToString();

                    if (dgvr.Cells[chName.Index].Value != null)
                        strName = dgvr.Cells[chName.Index].Value.ToString();

                    if (dgvr.Cells[chStudentNumber.Index].Value != null)
                        strStudentNumber = dgvr.Cells[chStudentNumber.Index].Value.ToString();

                    //string CoName = _course.SchoolYear + "學年度第" + _course.Semester + "學期" + _course.Name;
                    string CoName = _SchoolYear + "學年度第" + _Semester + "學期" + dgvr.Cells[chCourseName.Index].Value.ToString();

                    // key
                    string Key1 = CoName + ",試別:" + _ExamName + ",班級:" + strClassName + ",座號:" + strSeatNo + ",姓名:" + strName + ",學號:" + strStudentNumber + ",定期分數:";
                    string Key2 = CoName + ",試別:" + _ExamName + ",班級:" + strClassName + ",座號:" + strSeatNo + ",姓名:" + strName + ",學號:" + strStudentNumber + ",平時分數:";
                    string Key3 = CoName + ",試別:" + _ExamName + ",班級:" + strClassName + ",座號:" + strSeatNo + ",姓名:" + strName + ",學號:" + strStudentNumber + ",文字描述:";
                    string Value1 = string.Empty, Value2 = string.Empty, Value3 = string.Empty;

                    if (dgvr.Cells[chInputScore.Index].Value != null)
                        Value1 = dgvr.Cells[chInputScore.Index].Value.ToString();

                    if (dgvr.Cells[chInputAssignmentScore.Index].Value != null)
                        Value2 = dgvr.Cells[chInputAssignmentScore.Index].Value.ToString();

                    //if (dgvr.Cells[chInputText.Index].Value != null)
                    //    Value3 = dgvr.Cells[chInputText.Index].Value.ToString();

                    // 定期分數
                    prlp.SetBeforeSaveText(Key1, Value1);
                    // 平時分數
                    prlp.SetBeforeSaveText(Key2, Value2);
                    // 文字描述
                    //prlp.SetBeforeSaveText(Key3, Value3);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Log 發生錯誤");
            }

        }

        /// <summary>
        /// 儲存資料 Log
        /// </summary>
        public void SetSaveDataToLog()
        {
            try
            {
                foreach (DataGridViewRow dgvr in dgv.Rows)
                {
                    string strClassName = string.Empty, strSeatNo = string.Empty, strName = string.Empty, strStudentNumber = string.Empty;
                    if (dgvr.IsNewRow)
                        continue;

                    if (dgvr.Cells[chClassName.Index].Value != null)
                        strClassName = dgvr.Cells[chClassName.Index].Value.ToString();

                    if (dgvr.Cells[chSeatNo.Index].Value != null)
                        strSeatNo = dgvr.Cells[chSeatNo.Index].Value.ToString();

                    if (dgvr.Cells[chName.Index].Value != null)
                        strName = dgvr.Cells[chName.Index].Value.ToString();

                    if (dgvr.Cells[chStudentNumber.Index].Value != null)
                        strStudentNumber = dgvr.Cells[chStudentNumber.Index].Value.ToString();

                    //string CoName = _course.SchoolYear + "學年度第" + _course.Semester + "學期" + _course.Name;
                    string CoName = _SchoolYear + "學年度第" + _Semester + "學期" + dgvr.Cells[chCourseName.Index].Value.ToString();
                    // key
                    string Key1 = CoName + ",試別:" + _ExamName + ",班級:" + strClassName + ",座號:" + strSeatNo + ",姓名:" + strName + ",學號:" + strStudentNumber + ",定期分數:";
                    string Key2 = CoName + ",試別:" + _ExamName + ",班級:" + strClassName + ",座號:" + strSeatNo + ",姓名:" + strName + ",學號:" + strStudentNumber + ",平時分數:";
                    string Key3 = CoName + ",試別:" + _ExamName + ",班級:" + strClassName + ",座號:" + strSeatNo + ",姓名:" + strName + ",學號:" + strStudentNumber + ",文字描述:";
                    string Value1 = string.Empty, Value2 = string.Empty, Value3 = string.Empty;

                    if (dgvr.Cells[chInputScore.Index].Value != null)
                        Value1 = dgvr.Cells[chInputScore.Index].Value.ToString();

                    if (dgvr.Cells[chInputAssignmentScore.Index].Value != null)
                        Value2 = dgvr.Cells[chInputAssignmentScore.Index].Value.ToString();

                    //if (dgvr.Cells[chInputText.Index].Value != null)
                    //    Value3 = dgvr.Cells[chInputText.Index].Value.ToString();

                    // 定期分數
                    prlp.SetAfterSaveText(Key1, Value1);
                    // 平時分數
                    prlp.SetAfterSaveText(Key2, Value2);
                    // 文字描述
                    //prlp.SetAfterSaveText(Key3, Value3);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Log 發生錯誤");
            }

        }


        /// <summary>
        /// 選擇試別時觸發。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboExamList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboExamList.SelectedItem == null) return;

            foreach (K12.Data.ExamRecord ex in _exams)
            {
                if (ex.Name == cboExamList.Text)
                {
                    _ExamID = ex.ID;
                    break;
                }
            }

            FillStudentsToDataGridView();

        }

        /// <summary>
        /// 取得成績並填入DataGridView
        /// </summary>
        private void GetScoresAndFill()
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在取得評量成績");
            _dirtyCellList.Clear();
            lblSave.Visible = false;

            #region 清空所有評量欄位及SCETakeRecord

            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.Cells[chInputScore.Index].Value = row.Cells[chInputAssignmentScore.Index].Value = null;//row.Cells[chInputText.Index].Value = null;
                (row.Tag as SCAttendTag).SCETake = null;
            }

            #endregion

            #region 取得成績並填入
            List<HC.JHSCETakeRecord> list = JHSchool.Data.JHSCETake.Select(_CourseIDsList, _StudentIDsList, _ExamList, _SceTakeIDsList, _SCAttendIDsList).AsHCJHSCETakeRecords();
            //List<HC.JHSCETakeRecord> list = JHSchool.Data.JHSCETake.SelectByCourseAndExam(_CourseIDsList, examID).AsHCJHSCETakeRecords();
            foreach (var record in list)
            {
                if (_studentRowDict.ContainsKey(record.RefStudentID))
                {
                    DataGridViewRow row = _studentRowDict[record.RefStudentID];
                    row.Cells[chInputScore.Index].Value = record.Score;
                    row.Cells[chInputScore.Index].Tag = record.Score;
                    if (record.Score.HasValue && record.Score.Value < 0)
                    {
                        if (PluginMain.ScoreValueMap.ContainsKey(record.Score.Value))
                        {
                            row.Cells[chInputScore.Index].Value = PluginMain.ScoreValueMap[record.Score.Value].UseText;
                            row.Cells[chInputScore.Index].Tag = PluginMain.ScoreValueMap[record.Score.Value].UseText;
                        }
                    }
                    row.Cells[chInputAssignmentScore.Index].Value = record.AssignmentScore;
                    row.Cells[chInputAssignmentScore.Index].Tag = record.AssignmentScore;
                    if (record.AssignmentScore.HasValue && record.AssignmentScore.Value < 0)
                    {
                        if (PluginMain.ScoreValueMap.ContainsKey(record.AssignmentScore.Value))
                        {
                            row.Cells[chInputAssignmentScore.Index].Value = PluginMain.ScoreValueMap[record.AssignmentScore.Value].UseText;
                            row.Cells[chInputAssignmentScore.Index].Tag = PluginMain.ScoreValueMap[record.AssignmentScore.Value].UseText;
                        }
                    }
                    //row.Cells[chInputText.Index].Value = record.Text;
                    //row.Cells[chInputText.Index].Tag = record.Text;

                    if (record.Score < 60) row.Cells[chInputScore.Index].Style.ForeColor = Color.Red;
                    if (record.Score > 100 || record.Score < 0) row.Cells[chInputScore.Index].Style.ForeColor = Color.Green;
                    else row.Cells[chInputScore.Index].Style.ForeColor = Color.Black;

                    if (record.AssignmentScore < 60) row.Cells[chInputAssignmentScore.Index].Style.ForeColor = Color.Red;
                    if (record.AssignmentScore > 100 || record.AssignmentScore < 0) row.Cells[chInputAssignmentScore.Index].Style.ForeColor = Color.Green;
                    else row.Cells[chInputAssignmentScore.Index].Style.ForeColor = Color.Black;

                    SCAttendTag tag = row.Tag as SCAttendTag;
                    tag.SCETake = record;
                }
                else
                {
                    #region 除錯用，別刪掉

                    //StudentRecord student = Student.Instance.Items[record.RefStudentID];
                    //if (student == null)
                    //    MsgBox.Show("系統編號「" + record.RefStudentID + "」的學生不存在…");
                    //else
                    //{
                    //    string className = (student.Class != null) ? student.Class.Name + "  " : "";
                    //    string seatNo = string.IsNullOrEmpty(className) ? "" : (string.IsNullOrEmpty(student.SeatNo) ? "" : student.SeatNo + "  ");
                    //    string studentNumber = string.IsNullOrEmpty(student.StudentNumber) ? "" : " (" + student.StudentNumber + ")";

                    //    MsgBox.Show(className + seatNo + student.Name + studentNumber, "這個學生有問題喔…");
                    //}

                    #endregion
                }
            }

            #endregion

            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        /// <summary>
        /// 是否值有變更
        /// </summary>
        /// <returns></returns>
        private bool IsDirty()
        {
            return (_dirtyCellList.Count > 0);
        }

        /// <summary>
        /// 按下「儲存」時觸發。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cboExamList.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("沒有試別無法儲存。");
                return;
            }

            dgv.EndEdit();

            if (!IsValid())
            {
                MsgBox.Show("請先修正錯誤再儲存。");
                return;
            }

            try
            {
                bool checkLoadSave = false;
                RecordIUDLists lists = GetRecords();

                // 檢查超過 0~100
                if (lists.DeleteList.Count > 0)
                    foreach (HsinChu.JHEvaluation.Data.HC.JHSCETakeRecord sc in lists.DeleteList)
                    {
                        if (sc.AssignmentScore.HasValue)
                            if ((sc.AssignmentScore < 0 && !PluginMain.ScoreValueMap.ContainsKey(sc.AssignmentScore.Value)) || sc.AssignmentScore > 100)
                                checkLoadSave = true;
                        if (sc.Score.HasValue)
                            if ((sc.Score < 0 && !PluginMain.ScoreValueMap.ContainsKey(sc.Score.Value)) || sc.Score > 100)
                                checkLoadSave = true;

                    }

                if (lists.InsertList.Count > 0)
                    foreach (HsinChu.JHEvaluation.Data.HC.JHSCETakeRecord sc in lists.InsertList)
                    {
                        if (sc.AssignmentScore.HasValue)
                            if ((sc.AssignmentScore < 0 && !PluginMain.ScoreValueMap.ContainsKey(sc.AssignmentScore.Value)) || sc.AssignmentScore > 100)
                                checkLoadSave = true;
                        if (sc.Score.HasValue)
                            if ((sc.Score < 0 && !PluginMain.ScoreValueMap.ContainsKey(sc.Score.Value)) || sc.Score > 100)
                                checkLoadSave = true;

                    }

                if (lists.UpdateList.Count > 0)
                    foreach (HsinChu.JHEvaluation.Data.HC.JHSCETakeRecord sc in lists.UpdateList)
                    {
                        if (sc.AssignmentScore.HasValue)
                            if ((sc.AssignmentScore < 0 && !PluginMain.ScoreValueMap.ContainsKey(sc.AssignmentScore.Value)) || sc.AssignmentScore > 100)
                                checkLoadSave = true;
                        if (sc.Score.HasValue)
                            if ((sc.Score < 0 && !PluginMain.ScoreValueMap.ContainsKey(sc.Score.Value)) || sc.Score > 100)
                                checkLoadSave = true;
                    }

                if (checkLoadSave)
                {
                    CourseExtendControls.Ribbon.CheckSaveForm csf = new CourseExtendControls.Ribbon.CheckSaveForm();
                    csf.ShowDialog();
                    if (csf.DialogResult == DialogResult.Cancel)
                        return;
                }

                if (lists.InsertList.Count > 0)
                    JHSchool.Data.JHSCETake.Insert(lists.InsertList.AsJHSCETakeRecords());
                if (lists.UpdateList.Count > 0)
                    JHSchool.Data.JHSCETake.Update(lists.UpdateList.AsJHSCETakeRecords());
                if (lists.DeleteList.Count > 0)
                    JHSchool.Data.JHSCETake.Delete(lists.DeleteList.AsJHSCETakeRecords());

                ////記憶所選的試別(和成績輸入共用)
                //Campus.Configuration.ConfigData cd = Campus.Configuration.Config.User["新竹課程成績輸入考試別"];
                //cd["新竹課程成績輸入考試別"] = cboExamList.Text;
                //cd.Save();

                MsgBox.Show("儲存成功。");

                // 記修改後 log
                SetSaveDataToLog();
                // 存 Log
                prlp.SetActionBy("教務作業", "評量缺免成績查詢調整");
                prlp.SetAction("評量缺免成績查詢調整");
                prlp.SetDescTitle("");
                prlp.SaveLog("", "", "", "");

                // 重新記修改前 Log
                SetLoadDataToLog();

                //ExamComboBoxItem item = cboExamList.SelectedItem as ExamComboBoxItem;
                //K12.Data.ExamRecord exam = cboExamList.SelectedItem as K12.Data.ExamRecord;
                //string ExamID = "";
                //foreach (K12.Data.ExamRecord ex in _exams)
                //{
                //    if (ex.Name == cboExamList.Text)
                //    {
                //        ExamID = ex.ID;
                //        break;
                //    }
                //}
                //重新取得符合條件的學生
                FillStudentsToDataGridView();

            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存失敗。\n" + ex.Message);
            }

            // // 載入分數顏色
            LoadDvScoreColor();
        }

        /// <summary>
        /// 關閉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        /// <summary>
        /// 從DataGridView上取得 HC.JHSCETakeRecord
        /// </summary>
        /// <returns></returns>
        private RecordIUDLists GetRecords()
        {
            if (cboExamList.SelectedItem == null) return new RecordIUDLists();

            RecordIUDLists lists = new RecordIUDLists();

            ExamComboBoxItem item = cboExamList.SelectedItem as ExamComboBoxItem;

            //string ExamID = "";
            //foreach (K12.Data.ExamRecord ex in _exams)
            //{
            //    if (ex.Name == cboExamList.Text)
            //    {
            //        ExamID = ex.ID;
            //        break;
            //    }
            //}
            foreach (DataGridViewRow row in dgv.Rows)
            {
                SCAttendTag tag = row.Tag as SCAttendTag;
                if (tag.SCETake != null)
                {
                    HC.JHSCETakeRecord record = tag.SCETake;

                    #region 修改or刪除
                    bool is_remove = true;

                    if (chInputScore.Visible == true)
                    {
                        string value = ("" + row.Cells[chInputScore.Index].Value).Trim();
                        if (PluginMain.ScoreTextMap.ContainsKey(value))
                        {
                            value = PluginMain.ScoreTextMap[value].UseValue.ToString();
                        }
                        is_remove &= string.IsNullOrEmpty(value);
                        if (!string.IsNullOrEmpty(value))
                            record.Score = decimal.Parse(value);
                        else
                            record.Score = null;
                    }
                    if (chInputAssignmentScore.Visible == true)
                    {
                        string value = ("" + row.Cells[chInputAssignmentScore.Index].Value).Trim();
                        if (PluginMain.ScoreTextMap.ContainsKey(value))
                        {
                            value = PluginMain.ScoreTextMap[value].UseValue.ToString();
                        }
                        is_remove &= string.IsNullOrEmpty(value);
                        if (!string.IsNullOrEmpty(value))
                            record.AssignmentScore = decimal.Parse(value);
                        else
                            record.AssignmentScore = null;
                    }
                    //if (chInputText.Visible == true)
                    //{
                    //    string value = "" + row.Cells[chInputText.Index].Value;
                    //    is_remove &= string.IsNullOrEmpty(value);
                    //    record.Text = value;
                    //}

                    if (is_remove)
                        lists.DeleteList.Add(record);
                    else
                        lists.UpdateList.Add(record);
                    #endregion
                }
                else
                {
                    #region 新增
                    bool is_add = false;

                    JHSchool.Data.JHSCETakeRecord jh = new JHSchool.Data.JHSCETakeRecord();
                    HC.JHSCETakeRecord record = new HC.JHSCETakeRecord(jh);

                    record.RefCourseID = tag.SCAttend.Course.ID;
                    record.RefExamID = _ExamID;
                    record.RefSCAttendID = tag.SCAttend.ID;
                    record.RefStudentID = tag.SCAttend.Student.ID;

                    record.Score = null;
                    record.AssignmentScore = null;
                    record.Text = string.Empty;

                    if (chInputScore.Visible == true)
                    {
                        string value = ("" + row.Cells[chInputScore.Index].Value).Trim();
                        if (PluginMain.ScoreTextMap.ContainsKey(value))
                        {
                            value = PluginMain.ScoreTextMap[value].UseValue.ToString();
                        }
                        if (!string.IsNullOrEmpty(value))
                        {
                            record.Score = decimal.Parse(value);
                            is_add = true;
                        }
                    }
                    if (chInputAssignmentScore.Visible == true)
                    {
                        string value = ("" + row.Cells[chInputAssignmentScore.Index].Value).Trim();
                        if (PluginMain.ScoreTextMap.ContainsKey(value))
                        {
                            value = PluginMain.ScoreTextMap[value].UseValue.ToString();
                        }
                        if (!string.IsNullOrEmpty(value))
                        {
                            record.AssignmentScore = decimal.Parse(value);
                            is_add = true;
                        }
                    }
                    //if (chInputText.Visible == true)
                    //{
                    //    string value = "" + row.Cells[chInputText.Index].Value;
                    //    if (!string.IsNullOrEmpty(value))
                    //    {
                    //        record.Text = value;
                    //        is_add = true;
                    //    }
                    //}

                    if (is_add) lists.InsertList.Add(record);
                    #endregion
                }
            }

            return lists;
        }


        /// <summary>
        /// 驗證每個欄位是否正確。
        /// 有錯誤訊息表示不正確。
        /// </summary>
        /// <returns></returns>
        private bool IsValid()
        {
            bool valid = true;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (chInputScore.Visible == true && !string.IsNullOrEmpty(row.Cells[chInputScore.Index].ErrorText))
                    valid = false;
                if (chInputAssignmentScore.Visible == true && !string.IsNullOrEmpty(row.Cells[chInputAssignmentScore.Index].ErrorText))
                    valid = false;
                //if (chInputText.Visible == true && !string.IsNullOrEmpty(row.Cells[chInputText.Index].ErrorText))
                //    valid = false;
            }

            return valid;
        }

        /// <summary>
        /// 點欄位立即進入編輯。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != chInputScore.Index && e.ColumnIndex != chInputAssignmentScore.Index)//&& e.ColumnIndex != chInputText.Index)
                return;
            if (e.RowIndex < 0) return;

            dgv.BeginEdit(true);
        }

        /// <summary>
        /// 當欄位結束編輯，進行驗證。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != chInputScore.Index && e.ColumnIndex != chInputAssignmentScore.Index)//&& e.ColumnIndex != chInputText.Index) 
                return;
            if (e.RowIndex < 0) return;

            DataGridViewCell cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];

            if (cell.OwningColumn == chInputScore || cell.OwningColumn == chInputAssignmentScore)
            {
                #region 驗證分數 & 低於60分變紅色
                cell.Style.ForeColor = Color.Black;

                if (!string.IsNullOrEmpty("" + cell.Value))
                {
                    decimal d;
                    if (!decimal.TryParse("" + cell.Value, out d))
                    {
                        if (PluginMain.ScoreTextMap.Keys.Count > 0)
                        {
                            if (!PluginMain.ScoreTextMap.ContainsKey("" + cell.Value))
                                cell.ErrorText = "分數必須為數字或「" + string.Join("、", PluginMain.ScoreTextMap.Keys) + "」";
                            else
                            {
                                if (PluginMain.ScoreTextMap["" + cell.Value].AllowCalculation)
                                    cell.Style.ForeColor = Color.Purple;  //要計算(無論分數)，顯示紫色
                                if (PluginMain.ScoreTextMap["" + cell.Value].Score == 0 && PluginMain.ScoreTextMap["" + cell.Value].AllowCalculation)
                                    cell.Style.ForeColor = Color.Red;  //要計算且依照0分計算，顯示紅色
                                if (!PluginMain.ScoreTextMap["" + cell.Value].AllowCalculation)
                                    cell.Style.ForeColor = Color.Blue;//不要計算，顯示藍色
                            }
                        }
                        else
                        {
                            cell.ErrorText = "分數必須為數字";
                        }
                    }
                    else
                    {
                        cell.ErrorText = "";
                        if (d < 60)
                            cell.Style.ForeColor = Color.Red;
                        if (d > 100 || d < 0)
                            cell.Style.ForeColor = Color.Green;
                    }
                }
                else
                    cell.ErrorText = "";
                #endregion
            }

            #region 轉換文字描述 (文字評量代碼表)
            //else if (cell.OwningColumn == chInputText)
            //{
            //    if (!string.IsNullOrEmpty("" + cell.Value))
            //    {
            //        string orig = "" + cell.Value;
            //        foreach (string code in _textMapping.Keys)
            //        {
            //            if (orig.Contains(code))
            //            {
            //                int start = orig.IndexOf(code);
            //                orig = orig.Substring(0, start) + _textMapping[code] + orig.Substring(start + code.Length);
            //            }
            //        }
            //        cell.Value = orig;
            //    }
            //}
            #endregion

            #region 檢查是否有變更過
            if ("" + cell.Tag != "" + cell.Value)
            {
                if (!_dirtyCellList.Contains(cell)) _dirtyCellList.Add(cell);
            }
            else
            {
                if (_dirtyCellList.Contains(cell)) _dirtyCellList.Remove(cell);
            }
            lblSave.Visible = IsDirty();
            #endregion
        }

        /// <summary>
        /// Data Class
        /// 包含 AEIncludeRecord 的 ComboBoxItem
        /// </summary>
        private class ExamComboBoxItem
        {
            public string DisplayText
            {
                get
                {
                    if (_examRecord != null)
                        return _examRecord.Name;
                    return string.Empty;
                }
            }

            //public HC.JHAEIncludeRecord AEIncludeRecord
            //{
            //    get { return _aeIncludeRecord; }
            //}

            private HC.JHAEIncludeRecord _aeIncludeRecord;
            private JHExamRecord _examRecord;


            public ExamComboBoxItem(HC.JHAEIncludeRecord record)
            {
                _aeIncludeRecord = record;
                _examRecord = record.Exam;
            }
        }

        /// <summary>
        /// 排序欄位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column == chSeatNo || e.Column == chStudentNumber)
            {
                int result = 0;

                int a, b;
                if (int.TryParse("" + e.CellValue1, out a) && int.TryParse("" + e.CellValue2, out b))
                    result = a.CompareTo(b);
                else if (int.TryParse("" + e.CellValue1, out a))
                    result = -1;
                else if (int.TryParse("" + e.CellValue2, out b))
                    result = 1;
                else
                    result = 0;

                e.SortResult = result;
            }
            e.Handled = true;
        }

        /// <summary>
        /// Data Class
        /// 包含 SCAttendRecord 與 SCETakeRecord
        /// </summary>
        private class SCAttendTag
        {
            public JHSCAttendRecord SCAttend { get; set; }
            public HC.JHSCETakeRecord SCETake { get; set; }
        }

        private class RecordIUDLists
        {
            public List<HC.JHSCETakeRecord> InsertList { get; set; }
            public List<HC.JHSCETakeRecord> UpdateList { get; set; }
            public List<HC.JHSCETakeRecord> DeleteList { get; set; }

            public RecordIUDLists()
            {
                InsertList = new List<HC.JHSCETakeRecord>();
                UpdateList = new List<HC.JHSCETakeRecord>();
                DeleteList = new List<HC.JHSCETakeRecord>();
            }
        }

        /// <summary>
        /// 匯出dgv
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "匯出評量缺免成績";
            saveFileDialog1.Filter = "Excel (*.xls)|*.xls";
            //指定路徑在Reports
            string path = Path.Combine(Application.StartupPath, "Reports");
            saveFileDialog1.InitialDirectory = path;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, saveFileDialog1.FileName + ".xls");
            //

            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            DataGridViewExport export = new DataGridViewExport(dgv);
            export.Save(saveFileDialog1.FileName);
        }


    }
}
