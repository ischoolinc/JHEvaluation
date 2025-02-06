using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.IO;
using K12.Data;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.Model;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.Validation;
using JHSchool.Data;
using KaoHsiung.JHEvaluation.Data;
using KaoHsiung.ReaderScoreImport_SubjectMakeUp.Validation.RecordValidators;

namespace KaoHsiung.ReaderScoreImport_SubjectMakeUp
{
    public partial class ImportStartupForm : BaseForm
    {
        private BackgroundWorker _worker;
        private BackgroundWorker _upload;
        private BackgroundWorker _warn;

        private int SchoolYear { get; set; }
        private int Semester { get; set; }

        private DataValidator<RawData> _rawDataValidator;
        private DataValidator<DataRecord> _dataRecordValidator;

        private List<FileInfo> _files;

        private List<JHSemesterScoreRecord> _updateMakeUpScoreList = new List<JHSemesterScoreRecord>();
        private List<JHSemesterScoreRecord> _duplicateMakeUpScoreList = new List<JHSemesterScoreRecord>();

        //記錄錯誤訊息使用
        private List<string> msgList = new List<string>();

        //記錄重覆補考成績錯誤訊息使用
        private Dictionary<string, string> msg_DuplicatedDict = new Dictionary<string, string>();

        /// <summary>
        /// 儲存畫面上學號長度
        /// </summary>
        K12.Data.Configuration.ConfigData cd;

        private string _StudentNumberLenght = "國中匯入補考讀卡學號長度";
        private string _StudentNumberLenghtName = "StudentNumberLenght";

        private EffortMapper _effortMapper;
        double counter = 0; //上傳成績時，算筆數用的。

        /// <summary>
        /// 載入學號長度值
        /// </summary>
        private void LoadConfigData()
        {
            int val = 7;
            cd = School.Configuration[_StudentNumberLenght];
            Global.StudentNumberLenght = intStudentNumberLenght.Value;
            if (int.TryParse(cd[_StudentNumberLenghtName], out val))
                intStudentNumberLenght.Value = val;
        }


        /// <summary>
        /// 儲存學號長度值
        /// </summary>
        private void SaveConfigData()
        {
            Global.StudentNumberLenght = intStudentNumberLenght.Value;
            cd[_StudentNumberLenghtName] = intStudentNumberLenght.Value.ToString();
            cd.Save();
        }

        public ImportStartupForm()
        {
            // 每次開啟，就重新載入 代碼對照
            KaoHsiung.ReaderScoreImport_SubjectMakeUp.Mapper.ClassCodeMapper.Instance.Reload();
            KaoHsiung.ReaderScoreImport_SubjectMakeUp.Mapper.ExamCodeMapper.Instance.Reload();
            KaoHsiung.ReaderScoreImport_SubjectMakeUp.Mapper.SubjectCodeMapper.Instance.Reload();

            InitializeComponent();
            InitializeSemesters();

            _effortMapper = new EffortMapper();

            // 載入預設儲存值
            LoadConfigData();

            _worker = new BackgroundWorker();
            _worker.WorkerReportsProgress = true;
            _worker.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                lblMessage.Text = "" + e.UserState;
            };
            _worker.DoWork += delegate (object sender, DoWorkEventArgs e)
            {
                //將錯誤訊息刪光，以便重新統計
                msgList.Clear();


                #region Worker DoWork
                _worker.ReportProgress(0, "檢查讀卡文字格式…");

                #region 檢查文字檔
                ValidateTextFiles vtf = new ValidateTextFiles(intStudentNumberLenght.Value);
                ValidateTextResult vtResult = vtf.CheckFormat(_files);
                if (vtResult.Error)
                {
                    e.Result = vtResult;
                    return;
                }
                #endregion

                //文字檔轉 RawData
                RawDataCollection rdCollection = new RawDataCollection();
                rdCollection.ConvertFromFiles(_files);

                //RawData 轉 DataRecord
                DataRecordCollection drCollection = new DataRecordCollection();
                drCollection.ConvertFromRawData(rdCollection);

                _rawDataValidator = new DataValidator<RawData>();
                _dataRecordValidator = new DataValidator<DataRecord>();

                #region 取得驗證需要的資料
                JHCourse.RemoveAll();
                _worker.ReportProgress(0, "取得學生資料…");
                List<JHStudentRecord> studentList = GetInSchoolStudents();

                List<string> s_ids = new List<string>();
                Dictionary<string, string> studentNumberToStudentIDs = new Dictionary<string, string>();
                foreach (JHStudentRecord student in studentList)
                {
                    string sn = SCValidatorCreator.GetStudentNumberFormat(student.StudentNumber);
                    if (!studentNumberToStudentIDs.ContainsKey(sn))
                        studentNumberToStudentIDs.Add(sn, student.ID);
                }

                //紀錄現在所有txt 上 的學號，以處理重覆學號學號問題
                List<string> s_numbers = new List<string>();


                foreach (var dr in drCollection)
                {
                    if (studentNumberToStudentIDs.ContainsKey(dr.StudentNumber))
                    {
                        s_ids.Add(studentNumberToStudentIDs[dr.StudentNumber]);
                    }
                    else
                    {
                        //學號不存在系統中，下面的Validator 會進行處理

                    }

                    //2017/1/8 穎驊註解， 原本的程式碼並沒有驗證學號重覆的問題，且無法在他Validator 的處理邏輯加入，故在此驗證
                    if (!s_numbers.Contains(dr.StudentNumber))
                    {
                        s_numbers.Add(dr.StudentNumber);
                    }
                    else
                    {
                        msgList.Add("學號:" + dr.StudentNumber + "，資料重覆，請檢察資料來源txt是否有重覆填寫的學號資料。");

                    }
                }

                studentList.Clear();

                _worker.ReportProgress(0, "取得學期科目成績…");
                List<JHSemesterScoreRecord> jhssr_list = JHSemesterScore.SelectBySchoolYearAndSemester(s_ids, SchoolYear, Semester);

                #endregion

                #region 註冊驗證
                _worker.ReportProgress(0, "載入驗證規則…");
                _rawDataValidator.Register(new SubjectCodeValidator());
                _rawDataValidator.Register(new ClassCodeValidator());
                _rawDataValidator.Register(new ExamCodeValidator());

                SCValidatorCreator scCreator = new SCValidatorCreator(JHStudent.SelectByIDs(s_ids));
                _dataRecordValidator.Register(scCreator.CreateStudentValidator());
                //_dataRecordValidator.Register(new ExamValidator(examList));
                //_dataRecordValidator.Register(scCreator.CreateSCAttendValidator());
                //_dataRecordValidator.Register(new CourseExamValidator(scCreator.StudentCourseInfo, aeList, examList));
                #endregion

                #region 進行驗證
                _worker.ReportProgress(0, "進行驗證中…");


                foreach (RawData rawData in rdCollection)
                {
                    List<string> msgs = _rawDataValidator.Validate(rawData);
                    msgList.AddRange(msgs);
                }
                if (msgList.Count > 0)
                {
                    e.Result = msgList;
                    return;
                }

                foreach (DataRecord dataRecord in drCollection)
                {
                    List<string> msgs = _dataRecordValidator.Validate(dataRecord);
                    msgList.AddRange(msgs);
                }
                if (msgList.Count > 0)
                {
                    e.Result = msgList;
                    return;
                }
                #endregion

                #region 取得學生的評量成績
                _duplicateMakeUpScoreList.Clear();
                _updateMakeUpScoreList.Clear();

                //Dictionary<string, JHSCETakeRecord> sceList = new Dictionary<string, JHSCETakeRecord>();
                //FunctionSpliter<string, JHSCETakeRecord> spliterSCE = new FunctionSpliter<string, JHSCETakeRecord>(300, 3);
                //spliterSCE.Function = delegate(List<string> part)
                //{
                //    return JHSCETake.Select(null, null, null, null, part);
                //};
                //foreach (JHSCETakeRecord sce in spliterSCE.Execute(scaIDs.ToList()))
                //{
                //    string key = GetCombineKey(sce.RefStudentID, sce.RefCourseID, sce.RefExamID);
                //    if (!sceList.ContainsKey(key))
                //        sceList.Add(key, sce);
                //}

                //Dictionary<string, JHExamRecord> examTable = new Dictionary<string, JHExamRecord>();
                //Dictionary<string, JHSCAttendRecord> scaTable = new Dictionary<string, JHSCAttendRecord>();

                //foreach (JHExamRecord exam in examList)
                //    if (!examTable.ContainsKey(exam.Name))
                //        examTable.Add(exam.Name, exam);

                //foreach (JHSCAttendRecord sca in scaList)
                //{
                //    string key = GetCombineKey(sca.RefStudentID, sca.RefCourseID);
                //    if (!scaTable.ContainsKey(key))
                //        scaTable.Add(key, sca);
                //}

                //2018/1/8 穎驊新增
                //填寫補考成績
                foreach (DataRecord dr in drCollection)
                {
                    // 利用學號 將學生的isd 對應出來
                    string s_id = studentNumberToStudentIDs[dr.StudentNumber];

                    // 紀錄學生是否有學期成績，如果連學期成績都沒有，本補考成績將會無法匯入(因為不合理)
                    bool haveSemesterRecord = false;

                    foreach (JHSemesterScoreRecord jhssr in jhssr_list)
                    {
                        if (jhssr.RefStudentID == s_id)
                        {
                            haveSemesterRecord = true;

                            if (jhssr.Subjects.ContainsKey(dr.Subject))
                            {
                                // 假如原本已經有補考成績，必須將之加入waring_list，讓使用者決定是否真的要覆蓋
                                if (jhssr.Subjects[dr.Subject].ScoreMakeup != null)
                                {
                                    string warning = "學號:" + dr.StudentNumber + "學生，在學年度: " + SchoolYear + "，學期: " + Semester + "，科目: " + dr.Subject + "已有補考成績: " + jhssr.Subjects[dr.Subject].ScoreMakeup + "，本次匯入將會將其覆蓋取代。";

                                    _duplicateMakeUpScoreList.Add(jhssr);

                                    if (!msg_DuplicatedDict.ContainsKey(s_id))
                                    {
                                        msg_DuplicatedDict.Add(s_id, warning);
                                    }
                                    else
                                    {

                                    }

                                    jhssr.Subjects[dr.Subject].ScoreMakeup = dr.Score;

                                    _updateMakeUpScoreList.Add(jhssr);
                                }
                                else
                                {
                                    jhssr.Subjects[dr.Subject].ScoreMakeup = dr.Score;

                                    _updateMakeUpScoreList.Add(jhssr);
                                }
                            }
                            else
                            {
                                // 2018/1/8 穎驊註解，針對假如該學生在當學期成績卻沒有其科目成績時，跳出提醒，因為正常情境是科目成績計算出來不及格，才需要補考。
                                msgList.Add("學號:" + dr.StudentNumber + "學生，在學年度:" + SchoolYear + "，學期:" + Semester + "，並無科目:" + dr.Subject + "成績，請先至學生上確認是否已有結算資料。");
                            }
                        }
                    }

                    if (!haveSemesterRecord)
                    {
                        msgList.Add("學號:" + dr.StudentNumber + "學生，在學年度:" + SchoolYear + "，學期:" + Semester + "，並無學期成績，請先至學生上確認是否已有結算資料。");
                    }

                }

                if (msgList.Count > 0)
                {
                    e.Result = msgList;
                    return;
                }

                #endregion

                e.Result = null;
                #endregion
            };
            _worker.RunWorkerCompleted += delegate (object sender, RunWorkerCompletedEventArgs e)
            {
                #region Worker Completed
                if (e.Error == null && e.Result == null)
                {
                    if (!_upload.IsBusy)
                    {
                        //如果學生身上已有補考成績，則提醒使用者
                        if (_duplicateMakeUpScoreList.Count > 0)
                        {
                            _warn.RunWorkerAsync();
                        }
                        else
                        {
                            lblMessage.Text = "成績上傳中…";
                            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績上傳中…", 0);
                            counter = 0;
                            _upload.RunWorkerAsync();
                        }
                    }
                }
                else
                {
                    ControlEnable = true;

                    if (e.Error != null)
                    {
                        MsgBox.Show("匯入失敗。" + e.Error.Message);
                        SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);

                    }
                    else if (e.Result != null && e.Result is ValidateTextResult)
                    {
                        ValidateTextResult result = e.Result as ValidateTextResult;
                        ValidationErrorViewer viewer = new ValidationErrorViewer();
                        viewer.SetTextFileError(result.LineIndexes, result.ErrorFormatLineIndexes, result.DuplicateLineIndexes);
                        viewer.ShowDialog();
                    }
                    else if (e.Result != null && e.Result is List<string>)
                    {
                        ValidationErrorViewer viewer = new ValidationErrorViewer();
                        viewer.SetErrorLines(e.Result as List<string>);
                        viewer.ShowDialog();
                    }
                }
                #endregion
            };

            _upload = new BackgroundWorker();
            _upload.WorkerReportsProgress = true;
            _upload.ProgressChanged += new ProgressChangedEventHandler(_upload_ProgressChanged);
            //_upload.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            //{
            //    counter += double.Parse("" + e.ProgressPercentage);
            //    FISCA.Presentation.MotherForm.SetStatusBarMessage("成績上傳中…", (int)(counter * 100f / (double)_addScoreList.Count));
            //};
            _upload.DoWork += new DoWorkEventHandler(_upload_DoWork);


            _upload.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_upload_RunWorkerCompleted);

            _warn = new BackgroundWorker();
            _warn.WorkerReportsProgress = true;
            _warn.DoWork += delegate (object sender, DoWorkEventArgs e)
            {
                _warn.ReportProgress(0, "產生警告訊息...");

                Dictionary<string, string> examDict = new Dictionary<string, string>();
                foreach (JHExamRecord exam in JHExam.SelectAll())
                {
                    if (!examDict.ContainsKey(exam.ID))
                        examDict.Add(exam.ID, exam.Name);
                }

                WarningForm form = new WarningForm();
                int count = 0;
                foreach (JHSemesterScoreRecord sce in _duplicateMakeUpScoreList)
                {
                    form.Add(sce.RefStudentID, sce.Student.Name, msg_DuplicatedDict[sce.RefStudentID]);
                    _warn.ReportProgress((int)(count * 100 / _duplicateMakeUpScoreList.Count), "產生警告訊息...");
                }

                e.Result = form;
            };
            _warn.RunWorkerCompleted += delegate (object sender, RunWorkerCompletedEventArgs e)
            {
                WarningForm form = e.Result as WarningForm;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    lblMessage.Text = "成績上傳中…";
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("成績上傳中…", 0);
                    counter = 0;
                    _upload.RunWorkerAsync();
                }
                else
                {
                    this.DialogResult = DialogResult.Cancel;
                }
            };
            _warn.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
            };

            _files = new List<FileInfo>();

        }

        void _upload_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績上傳中…", (int)(counter * 100f / (double)_updateMakeUpScoreList.Count));
        }

        // 上傳成績完成
        void _upload_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                string msg = "";
                if (e.Result != null)
                    msg = e.Result.ToString();

                if (e.Cancelled)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("匯入失敗");
                    MsgBox.Show("匯入失敗" + msg);
                }
                else
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("匯入完成");
                    MsgBox.Show("匯入完成,共匯入" + msg + "筆.");
                }
                ControlEnable = true;
            }
            catch (Exception ex)
            {

            }
            //關閉視窗
            this.Close();
        }

        // 上傳
        void _upload_DoWork(object sender, DoWorkEventArgs e)
        {
            // 傳送與回傳筆數
            int SendCount = 0, RspCount = 0;

            try
            {

                //新增資料，分筆上傳
                Dictionary<int, List<JHSemesterScoreRecord>> batchDict = new Dictionary<int, List<JHSemesterScoreRecord>>();
                int bn = 150;
                int n1 = (int)(_updateMakeUpScoreList.Count / bn);

                if ((_updateMakeUpScoreList.Count % bn) != 0)
                    n1++;

                for (int i = 0; i <= n1; i++)
                    batchDict.Add(i, new List<JHSemesterScoreRecord>());


                if (_updateMakeUpScoreList.Count > 0)
                {
                    int idx = 0, count = 1;
                    // 分批
                    foreach (JHSemesterScoreRecord rec in _updateMakeUpScoreList)
                    {
                        // 100 分一批
                        if ((count % bn) == 0)
                            idx++;

                        batchDict[idx].Add(rec);
                        count++;
                    }
                }

                // 上傳資料
                foreach (KeyValuePair<int, List<JHSemesterScoreRecord>> data in batchDict)
                {
                    SendCount = 0; RspCount = 0;
                    if (data.Value.Count > 0)
                    {
                        SendCount = data.Value.Count;
                        try
                        {
                            JHSemesterScore.Update(data.Value);
                        }
                        catch (Exception ex)
                        {
                            e.Cancel = true;
                            e.Result = ex.Message;
                        }

                        counter += SendCount;

                    }
                }
                e.Result = _updateMakeUpScoreList.Count;
            }
            catch (Exception ex)
            {
                e.Result = ex.Message;
                e.Cancel = true;

            }
        }

        private List<JHStudentRecord> GetInSchoolStudents()
        {
            List<JHStudentRecord> list = new List<JHStudentRecord>();
            foreach (JHStudentRecord student in JHStudent.SelectAll())
            {
                if (student.Status == StudentRecord.StudentStatus.一般 ||
                    student.Status == StudentRecord.StudentStatus.輟學)
                    list.Add(student);
            }
            return list;
        }

        private string GetCombineKey(string s1, string s2, string s3)
        {
            return s1 + "_" + s2 + "_" + s3;
        }

        private string GetCombineKey(string s1, string s2)
        {
            return s1 + "_" + s2;
        }

        private void InitializeSemesters()
        {
            try
            {
                for (int i = -2; i <= 2; i++)
                {
                    cboSchoolYear.Items.Add(int.Parse(School.DefaultSchoolYear) + i);
                }
                cboSemester.Items.Add(1);
                cboSemester.Items.Add(2);

                cboSchoolYear.SelectedIndex = 2;
                cboSemester.SelectedIndex = int.Parse(School.DefaultSemester) - 1;
            }
            catch { }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.Filter = "純文字文件(*.txt)|*.txt";
            ofd.Multiselect = true;
            ofd.Title = "開啟檔案";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                _files.Clear();
                StringBuilder builder = new StringBuilder("");
                foreach (var file in ofd.FileNames)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    _files.Add(fileInfo);
                    builder.Append(fileInfo.Name + ", ");
                }
                string fileString = builder.ToString();
                if (fileString.EndsWith(", ")) fileString = fileString.Substring(0, fileString.Length - 2);
                txtFiles.Text = fileString;
            }
        }

        private bool ControlEnable
        {
            set
            {
                foreach (Control ctrl in this.Controls)
                    ctrl.Enabled = value;

                pic.Enabled = lblMessage.Enabled = !value;
                pic.Visible = lblMessage.Visible = !value;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (cboSchoolYear.SelectedItem == null) return;
            if (cboSemester.SelectedItem == null) return;
            if (_files.Count <= 0) return;

            ControlEnable = false;

            // 儲存設定值
            SaveConfigData();

            if (!_worker.IsBusy)
                _worker.RunWorkerAsync();
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboSchoolYear.SelectedItem != null)
                SchoolYear = (int)cboSchoolYear.SelectedItem;
        }

        private void cboSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboSemester.SelectedItem != null)
                Semester = (int)cboSemester.SelectedItem;
        }

        private void ImportStartupForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }


    }
}
