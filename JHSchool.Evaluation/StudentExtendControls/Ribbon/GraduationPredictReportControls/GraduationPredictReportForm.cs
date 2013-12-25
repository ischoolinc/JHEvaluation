using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Framework;
using JHSchool.Evaluation.Calculation;
using JHSchool.Evaluation.Calculation.GraduationConditions;
using Aspose.Cells;
using System.IO;
using JHSchool.Data;
using Campus.Report;
using FISCA.Presentation;

namespace JHSchool.Evaluation.StudentExtendControls.Ribbon.GraduationPredictReportControls
{
    public partial class GraduationPredictReportForm : FISCA.Presentation.Controls.BaseForm
    {
        private List<StudentRecord> _errorList;
        private Dictionary<string, bool> _passList;
        private EvaluationResult _result;
        private List<StudentRecord> _students;
        private BackgroundWorker _ExportWorker;
        private Aspose.Words.Document _doc;
        private Aspose.Words.Document _template;
        private ReportConfiguration _rc;
        private string ReportName = "";
        // 是否產生學生清單
        bool _UserSelExportStudentList = false;
        Workbook _wbStudentList;
        private MultiThreadBackgroundWorker<StudentRecord> _historyWorker, _inspectWorker;
        private const string _NewLine = "\r\n"; // 小郭, 2013/12/25

        public GraduationPredictReportForm(List<StudentRecord> students)
        {
            InitializeComponent();
            ReportName = "未達畢業標準通知單";
            _rc = new ReportConfiguration(ReportName);

            _students = students;
            _errorList = new List<StudentRecord>();
            _passList = new Dictionary<string, bool>();
            _result = new EvaluationResult();
            _doc = new Aspose.Words.Document();

            if (_rc.Template == null)
                _rc.Template = new ReportTemplate(Properties.Resources.未達畢業標準通知單樣板, TemplateType.Word);

            _template = _rc.Template.ToDocument();

            InitializeWorkers();
        }

        private void InitializeWorkers()
        {
            _historyWorker = new MultiThreadBackgroundWorker<StudentRecord>();
            _historyWorker.Loading = MultiThreadLoading.Light;
            _historyWorker.PackageSize = _students.Count; //暫解
            _historyWorker.AutoReportsProgress = true;
            _historyWorker.DoWork += new EventHandler<PackageDoWorkEventArgs<StudentRecord>>(HistoryWorker_DoWork);
            _historyWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(HistoryWorker_RunWorkerCompleted);
            _historyWorker.ProgressChanged += new ProgressChangedEventHandler(HistoryWorker_ProgressChanged);

            _inspectWorker = new MultiThreadBackgroundWorker<StudentRecord>();
            _inspectWorker.Loading = MultiThreadLoading.Light;
            _inspectWorker.PackageSize = _students.Count; //暫解
            _inspectWorker.AutoReportsProgress = true;
            _inspectWorker.DoWork += new EventHandler<PackageDoWorkEventArgs<StudentRecord>>(InspectWorker_DoWork);
            _inspectWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(InspectWorker_RunWorkerCompleted);
            _inspectWorker.ProgressChanged += new ProgressChangedEventHandler(InspectWorker_ProgressChanged);

            _ExportWorker = new BackgroundWorker();
            _ExportWorker.DoWork += new DoWorkEventHandler(_ExportWorker_DoWork);
            _ExportWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_ExportWorker_RunWorkerCompleted);
        }

        #region HistoryWorker Event
        private void HistoryWorker_DoWork(object sender, PackageDoWorkEventArgs<StudentRecord> e)
        {
            try
            {
                List<StudentRecord> error_list = e.Argument as List<StudentRecord>;
                error_list.AddRange(Graduation.Instance.CheckSemesterHistories(e.Items));
            }
            catch (Exception ex)
            {
                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                MsgBox.Show("檢查學期歷程時發生錯誤。" + ex.Message);
            }
        }

        private void HistoryWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //progressBar.Value = 100;

            if (e.Error != null)
            {
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                MsgBox.Show("檢查學期歷程時發生錯誤。" + e.Error.Message);
                return;
            }

            if (_errorList.Count > 0)
            {
                btnExit.Enabled = true;

                ErrorViewer viewer = new ErrorViewer();
                viewer.SetHeader("學生");
                foreach (StudentRecord student in _errorList)
                    viewer.SetMessage(student, new List<string>(new string[] { "學期歷程不完整" }));
                viewer.ShowDialog();
                return;
            }
            else
            {
                // 加入這段主要在處理當學期還沒有產生學期歷程，資料可以判斷
                UIConfig._StudentSHistoryRecDict.Clear();
                Dictionary<string, int> studGradeYearDict = new Dictionary<string, int>();

                // 取得學生ID
                List<string> studIDs = (from data in _students select data.ID).Distinct().ToList();

                foreach (JHStudentRecord stud in JHStudent.SelectByIDs(studIDs))
                {
                    if (stud.Class != null)
                    {
                        if (stud.Class.GradeYear.HasValue)
                            if (!studGradeYearDict.ContainsKey(stud.ID))
                                studGradeYearDict.Add(stud.ID, stud.Class.GradeYear.Value);
                    }
                }
                bool checkInsShi = false;
                // 取得學生學期歷程，並加入學生學習歷程Cache
                foreach (JHSemesterHistoryRecord rec in JHSemesterHistory.SelectByStudentIDs(studIDs))
                {
                    checkInsShi = true;
                    K12.Data.SemesterHistoryItem shi = new K12.Data.SemesterHistoryItem();
                    shi.SchoolYear = UIConfig._UserSetSHSchoolYear;
                    shi.Semester = UIConfig._UserSetSHSemester;
                    if (studGradeYearDict.ContainsKey(rec.RefStudentID))
                        shi.GradeYear = studGradeYearDict[rec.RefStudentID];

                    foreach (K12.Data.SemesterHistoryItem shiItem in rec.SemesterHistoryItems)
                        if (shiItem.SchoolYear == shi.SchoolYear && shiItem.Semester == shi.Semester)
                            checkInsShi = false;
                    if (checkInsShi)
                        rec.SemesterHistoryItems.Add(shi);

                    if (!UIConfig._StudentSHistoryRecDict.ContainsKey(rec.RefStudentID))
                        UIConfig._StudentSHistoryRecDict.Add(rec.RefStudentID, rec);
                }

                //lblProgress.Text = "畢業資格審查中…";
                FISCA.LogAgent.ApplicationLog.Log("成績系統.報表", "列印畢業預警報表", "產生畢業預警報表");

                _inspectWorker.RunWorkerAsync(_students, new object[] { _passList, _result });
            }
        }

        private void HistoryWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar.Value = e.ProgressPercentage;
        }
        #endregion

        #region InspectWorker Event
        private void InspectWorker_DoWork(object sender, PackageDoWorkEventArgs<StudentRecord> e)
        {
            try
            {
                object[] objs = e.Argument as object[];
                Dictionary<string, bool> passList = objs[0] as Dictionary<string, bool>;
                EvaluationResult result = objs[1] as EvaluationResult;

                Dictionary<string, bool> list = Graduation.Instance.Evaluate(e.Items);
                foreach (string id in list.Keys)
                {
                    if (!passList.ContainsKey(id))
                        passList.Add(id, list[id]);
                }

                if (Graduation.Instance.Result.Count > 0)
                {
                    MergeResult(list, Graduation.Instance.Result, result);
                }
            }
            catch (Exception ex)
            {
                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                MsgBox.Show("發生錯誤。" + ex.Message);
            }
        }

        private void InspectWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //progressBar.Value = 100;
            if (e.Error != null)
            {
                //btnExit.Enabled = true;
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                MsgBox.Show("審查時發生錯誤。" + e.Error.Message);
                return;
            }

            //lblProgress.Text = "審查完成";

            if (_passList.Count > 0)
            {
                //lblProgress.Text = "寫入類別資訊…";
                PrintOut();

                // 檢查是否需要產生通知單
                if (checkExportDoc.Checked)
                {
                    _ExportWorker.RunWorkerAsync();
                }

                btnPrint.Enabled = true;
            }


        }

        // 未達畢業標準通知單
        void _ExportWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MsgBox.Show("產生報表時發生錯誤。" + e.Error.Message);
                return;
            }

            MotherForm.SetStatusBarMessage(ReportName + "產生完成");
            ReportSaver.SaveDocument(_doc, ReportName);

            // 檢查是否產生學生清單
            if (_UserSelExportStudentList)
            {
                if (_wbStudentList != null)
                {
                    SaveFileDialog sd = new SaveFileDialog();
                    sd.FileName = "學生清單(未達畢業標準學生)";
                    sd.Filter = "Excel檔案(*.xls)|*.xls";
                    if (sd.ShowDialog() != DialogResult.OK) return;

                    try
                    {
                        _wbStudentList.Save(sd.FileName, FileFormatType.Excel2003);

                        if (MsgBox.Show("報表產生完成，是否立刻開啟？", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(sd.FileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                        MsgBox.Show("儲存失敗。" + ex.Message);
                    }
                }
            }
        }

        void _ExportWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            ExportDoc();
        }

        /// <summary>
        /// 日期轉換(2011/1/1)->民國年(100年1月1日)
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private string ConvertDate1(DateTime dt)
        {
            string retVal = string.Empty;
            retVal = (dt.Year - 1911) + "年" + dt.Month + "月" + dt.Day + "日";
            return retVal;
        }

        /// <summary>
        /// 未達畢業標準通知單
        /// </summary>
        private void ExportDoc()
        {
            if (_students.Count == 0) return;
            _doc.Sections.Clear();

            if (_rc.Template == null)
                _rc.Template = new ReportTemplate(Properties.Resources.未達畢業標準通知單樣板, TemplateType.Word);

            _template = _rc.Template.ToDocument();

            string UserSelAddresseeAddress = _rc.GetString(PrintConfigForm.setupAddresseeAddress, "聯絡地址");
            string UserSelAddresseeName = _rc.GetString(PrintConfigForm.setupAddresseeName, "監護人");
            _UserSelExportStudentList = _rc.GetBoolean(PrintConfigForm.setupExportStudentList, false);

            string UserSeldtDate = "";
            DateTime dt;
            if (DateTime.TryParse(_rc.GetString(PrintConfigForm.setupdtDocDate, ""), out dt))
                UserSeldtDate = ConvertDate1(dt);
            else
                UserSeldtDate = ConvertDate1(DateTime.Now);

            List<StudentGraduationPredictData> StudentGraduationPredictDataList = new List<StudentGraduationPredictData>();
            // 取得學生ID，製作 Dict 用
            List<string> StudIDList = (from data in _students select data.ID).ToList();

            // Student Address,Key:StudentID
            Dictionary<string, JHAddressRecord> AddressDict = new Dictionary<string, JHAddressRecord>();
            // Student Parent,Key:StudentID
            Dictionary<string, JHParentRecord> ParentDict = new Dictionary<string, JHParentRecord>();

            // 地址
            foreach (JHAddressRecord rec in JHAddress.SelectByStudentIDs(StudIDList))
                if (!AddressDict.ContainsKey(rec.RefStudentID))
                    AddressDict.Add(rec.RefStudentID, rec);

            // 父母監護人
            foreach (JHParentRecord rec in JHParent.SelectByStudentIDs(StudIDList))
                if (!ParentDict.ContainsKey(rec.RefStudentID))
                    ParentDict.Add(rec.RefStudentID, rec);




            // 資料轉換 ..
            foreach (StudentRecord StudRec in _students)
            {
                if (!_result.ContainsKey(StudRec.ID)) continue;

                StudentGraduationPredictData sgpd = new StudentGraduationPredictData();

                if (StudRec.Class != null)
                    sgpd.ClassName = StudRec.Class.Name;

                sgpd.Name = StudRec.Name;

                sgpd.SchoolAddress = K12.Data.School.Address;
                sgpd.SchoolName = K12.Data.School.ChineseName;
                sgpd.SchoolPhone = K12.Data.School.Telephone;

                sgpd.SeatNo = StudRec.SeatNo;
                sgpd.StudentNumber = StudRec.StudentNumber;

                // 文字
                if (_result.ContainsKey(StudRec.ID))
                {
                    int GrYear;
                    foreach (ResultDetail rd in _result[StudRec.ID])
                    {
                        if (int.TryParse(rd.GradeYear, out GrYear))
                        {
                            if (GrYear == 0) continue;

                            // 組訊息
                            string Detail = "";
                            if (rd.Details.Count > 0)
                                Detail = string.Join(",", rd.Details.ToArray());

                            // 一年級
                            if (GrYear == 1 || GrYear == 7)
                            {
                                if (rd.Semester.Trim() == "1")
                                    sgpd.Text11 = Detail;

                                if (rd.Semester.Trim() == "2")
                                    sgpd.Text12 = Detail;
                            }

                            // 二年級
                            if (GrYear == 2 || GrYear == 8)
                            {
                                if (rd.Semester.Trim() == "1")
                                    sgpd.Text21 = Detail;

                                if (rd.Semester.Trim() == "2")
                                    sgpd.Text22 = Detail;
                            }

                            // 三年級
                            if (GrYear == 3 || GrYear == 9)
                            {
                                if (rd.Semester.Trim() == "1")
                                    sgpd.Text31 = Detail;

                                if (rd.Semester.Trim() == "2")
                                    sgpd.Text32 = Detail;
                            }

                        }
                    }
                }

                // 地址
                if (AddressDict.ContainsKey(StudRec.ID))
                {
                    if (UserSelAddresseeAddress == "聯絡地址")
                        sgpd.AddresseeAddress = AddressDict[StudRec.ID].MailingAddress;

                    if (UserSelAddresseeAddress == "戶籍地址")
                        sgpd.AddresseeAddress = AddressDict[StudRec.ID].PermanentAddress;
                }

                // 父母監護人
                if (ParentDict.ContainsKey(StudRec.ID))
                {
                    if (UserSelAddresseeName == "父親")
                        sgpd.AddresseeName = ParentDict[StudRec.ID].FatherName;

                    if (UserSelAddresseeName == "母親")
                        sgpd.AddresseeName = ParentDict[StudRec.ID].MotherName;

                    if (UserSelAddresseeName == "監護人")
                        sgpd.AddresseeName = ParentDict[StudRec.ID].CustodianName;
                }

                sgpd.DocDate = UserSeldtDate;

                StudentGraduationPredictDataList.Add(sgpd);
            }

            // 產生Word 套印
            Dictionary<string, object> FieldData = new Dictionary<string, object>();

            // 班座排序
            StudentGraduationPredictDataList = (from data in StudentGraduationPredictDataList orderby data.ClassName, int.Parse(data.SeatNo) ascending select data).ToList();

            foreach (StudentGraduationPredictData sgpd in StudentGraduationPredictDataList)
            {
                FieldData.Clear();
                FieldData.Add("學校名稱", sgpd.SchoolName);
                FieldData.Add("學校電話", sgpd.SchoolPhone);
                FieldData.Add("學校地址", sgpd.SchoolAddress);
                FieldData.Add("收件人地址", sgpd.AddresseeAddress);
                FieldData.Add("收件人姓名", sgpd.AddresseeName);
                FieldData.Add("班級", sgpd.ClassName);
                FieldData.Add("座號", sgpd.SeatNo);
                FieldData.Add("姓名", sgpd.Name);
                FieldData.Add("學號", sgpd.StudentNumber);
                FieldData.Add("一上文字", sgpd.Text11);
                FieldData.Add("一下文字", sgpd.Text12);
                FieldData.Add("二上文字", sgpd.Text21);
                FieldData.Add("二下文字", sgpd.Text22);
                FieldData.Add("三上文字", sgpd.Text31);
                FieldData.Add("三下文字", sgpd.Text32);
                FieldData.Add("發文日期", sgpd.DocDate);

                Aspose.Words.Document each = (Aspose.Words.Document)_template.Clone(true);
                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(each);
                // 合併
                if (FieldData.Count > 0)
                    builder.Document.MailMerge.Execute(FieldData.Keys.ToArray(), FieldData.Values.ToArray());

                foreach (Aspose.Words.Section sec in each.Sections)
                    _doc.Sections.Add(_doc.ImportNode(sec, true));

            }

            // 產生學生清單
            if (_UserSelExportStudentList)
            {
                _wbStudentList = new Workbook();
                _wbStudentList.Worksheets[0].Cells[0, 0].PutValue("班級");
                _wbStudentList.Worksheets[0].Cells[0, 1].PutValue("座號");
                _wbStudentList.Worksheets[0].Cells[0, 2].PutValue("學號");
                _wbStudentList.Worksheets[0].Cells[0, 3].PutValue("學生姓名");
                _wbStudentList.Worksheets[0].Cells[0, 4].PutValue("收件人姓名");
                _wbStudentList.Worksheets[0].Cells[0, 5].PutValue("地址");
                //班級	座號	學號	學生姓名	收件人姓名	地址

                int rowIdx = 1;
                foreach (StudentGraduationPredictData sgpd in StudentGraduationPredictDataList)
                {
                    _wbStudentList.Worksheets[0].Cells[rowIdx, 0].PutValue(sgpd.ClassName);
                    _wbStudentList.Worksheets[0].Cells[rowIdx, 1].PutValue(sgpd.SeatNo);
                    _wbStudentList.Worksheets[0].Cells[rowIdx, 2].PutValue(sgpd.StudentNumber);
                    _wbStudentList.Worksheets[0].Cells[rowIdx, 3].PutValue(sgpd.Name);
                    _wbStudentList.Worksheets[0].Cells[rowIdx, 4].PutValue(sgpd.AddresseeName);
                    _wbStudentList.Worksheets[0].Cells[rowIdx, 5].PutValue(sgpd.AddresseeAddress);
                    rowIdx++;
                }

                _wbStudentList.Worksheets[0].AutoFitColumns();
            }
        }


        /// <summary>
        /// 未達畢業標準名冊
        /// </summary>
        private void PrintOut()
        {
            if (_students.Count <= 0) return;

            SaveFileDialog sd = new SaveFileDialog();
            sd.FileName = "未達畢業標準學生名冊";
            sd.Filter = "Excel檔案(*.xls)|*.xls";
            if (sd.ShowDialog() != DialogResult.OK) return;

            Workbook template = new Workbook();
            template.Open(new MemoryStream(Resources.未達畢業標準學生名冊template));
            Worksheet tempsheet = template.Worksheets[0];
            Worksheet tempsheet1 = template.Worksheets[1];

            Workbook book = new Workbook();
            book.Open(new MemoryStream(Resources.未達畢業標準學生名冊template));
            Worksheet sheet = book.Worksheets[0];
            Worksheet sheet1 = book.Worksheets[1];
            sheet.Name = "未達畢業標準學生";
            sheet1.Name = "未達畢業標準學生-依畢業總平均";

            Range temprow = tempsheet.Cells.CreateRange(3, 1, false);

            sheet.Cells[0, 0].PutValue(School.DefaultSchoolYear + "學年度 " + School.ChineseName);

            Range temprow1 = tempsheet1.Cells.CreateRange(3, 1, false);

            sheet1.Cells[0, 0].PutValue(School.DefaultSchoolYear + "學年度 " + School.ChineseName);

            int rowIndex = 3, rowIndex1 = 3; ;
            List<StudentRecord> sorted = new List<StudentRecord>();
            foreach (StudentRecord stu in _students)
                sorted.Add(stu);
            sorted.Sort();

            #region 處理畢業成績類類
            foreach (StudentRecord stu in sorted)
            {
                if (!_result.ContainsKey(stu.ID)) continue;

                sheet1.Cells.CreateRange(rowIndex1, 1, false).Copy(temprow1);
                sheet1.Cells.CreateRange(rowIndex1, 1, false).CopyStyle(temprow1);

                if (stu.Class != null) sheet1.Cells[rowIndex1, 0].PutValue(stu.Class.Name);
                sheet1.Cells[rowIndex1, 1].PutValue(stu.SeatNo);
                sheet1.Cells[rowIndex1, 2].PutValue(stu.StudentNumber);
                sheet1.Cells[rowIndex1, 3].PutValue(stu.Name);


                foreach (ResultDetail rd in _result[stu.ID])
                {
                    int gradeYear = int.Parse(rd.GradeYear);

                    if (gradeYear == 0)
                    {
                        string details = string.Empty;
                        foreach (string detail in rd.Details)
                            details += detail + "," + _NewLine;   // 小郭, 2013/12/25
                        if (details.EndsWith("," + _NewLine)) details = details.Substring(0, details.Length - ("," + _NewLine).Length); // 小郭, 2013/12/25
                        
                        sheet1.Cells[rowIndex1, 4].PutValue(details);
                    }
                }

                rowIndex1++;
            }
            #endregion

            #region 處理學期類
            foreach (StudentRecord stu in sorted)
            {
                if (!_result.ContainsKey(stu.ID)) continue;

                sheet.Cells.CreateRange(rowIndex, 1, false).Copy(temprow);
                sheet.Cells.CreateRange(rowIndex, 1, false).CopyStyle(temprow);

                if (stu.Class != null) sheet.Cells[rowIndex, 0].PutValue(stu.Class.Name);
                sheet.Cells[rowIndex, 1].PutValue(stu.SeatNo);
                sheet.Cells[rowIndex, 2].PutValue(stu.StudentNumber);
                sheet.Cells[rowIndex, 3].PutValue(stu.Name);


                foreach (ResultDetail rd in _result[stu.ID])
                {
                    int gradeYear = int.Parse(rd.GradeYear);

                    if (gradeYear == 0)
                        continue;

                    if (gradeYear > 6) gradeYear -= 6;

                    int index = (gradeYear - 1) * 2 + int.Parse(rd.Semester);

                    string details = string.Empty;
                    foreach (string detail in rd.Details)
                        details += detail + "," + _NewLine;   // 小郭, 2013/12/25
                    if (details.EndsWith("," + _NewLine)) details = details.Substring(0, details.Length - ("," + _NewLine).Length); // 小郭, 2013/12/25
                    sheet.Cells[rowIndex, index + 3].PutValue(details);
                }

                if (stu.Class != null)
                {
                    int i;
                    if (int.TryParse(stu.Class.GradeYear, out i))
                    {
                        int last = (i - 1) * 2 + 2;
                        for (int j = 6; j > last; j--)
                            sheet.Cells[rowIndex, j + 3].PutValue("- - - - -");
                    }
                }

                rowIndex++;
            }
            #endregion
            try
            {

                // 判斷工作表是否顯示
                if (Config.rpt_isCheckGraduateDomain == false)
                    book.Worksheets.RemoveAt("未達畢業標準學生-依畢業總平均");

                if (Config.rpt_isCheckSemesterDomain == false)
                    book.Worksheets.RemoveAt("未達畢業標準學生");


                book.Save(sd.FileName, FileFormatType.Excel2003);

                if (MsgBox.Show("報表產生完成，是否立刻開啟？", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(sd.FileName);
                }
            }
            catch (Exception ex)
            {
                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                MsgBox.Show("儲存失敗。" + ex.Message);
            }
        }

        private void InspectWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar.Value = e.ProgressPercentage;
        }
        #endregion

        private void btnPrint_Click(object sender, EventArgs e)
        {
            _result.Clear();

            // 檢查學業勾選
            Config.rpt_isCheckGraduateDomain = Config.rpt_isCheckSemesterDomain = false;
            /* 小郭, 2013/12/25
            if (chCondition1.Checked == true || chCondition2.Checked == true)
                Config.rpt_isCheckSemesterDomain = true;

            if (chConditionGr1.Checked)
                Config.rpt_isCheckGraduateDomain = true;
            */
            // 取得使用者選擇學年度學期
            int.TryParse(cboSchoolYear.Text, out UIConfig._UserSetSHSchoolYear);
            int.TryParse(cboSemester.Text, out UIConfig._UserSetSHSemester);

            if (UIConfig._UserSetSHSchoolYear == 0 || UIConfig._UserSetSHSemester == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("日前學年度或學期輸入錯誤.");
                return;
            }


            List<string> conditions = new List<string>();

            // 檢查畫面上畢業成績使用者勾選
            foreach (Control ctrl in gpScore.Controls)
            {
                if (ctrl is System.Windows.Forms.CheckBox)
                {
                    System.Windows.Forms.CheckBox chk = ctrl as System.Windows.Forms.CheckBox;
                    if (chk.Checked && !string.IsNullOrEmpty("" + chk.Tag))
                    {
                        if (!conditions.Contains("" + chk.Tag))
                            conditions.Add("" + chk.Tag);
                    }
                }
            }

            // 檢查畫面上日常生活表現使用者勾選
            foreach (Control ctrl in gpDaily.Controls)
            {
                if (ctrl is System.Windows.Forms.CheckBox)
                {
                    System.Windows.Forms.CheckBox chk = ctrl as System.Windows.Forms.CheckBox;
                    if (chk.Checked && !string.IsNullOrEmpty("" + chk.Tag))
                    {
                        if (!conditions.Contains("" + chk.Tag))
                            conditions.Add("" + chk.Tag);
                    }
                }
            }

            // 檢查使用者是否有勾選，至少有勾一項才執行。
            if (conditions.Count > 0)
            {

                Config.rpt_isCheckSemesterDomain = IsShowSemesterDomainSheet(); // 小郭, 2013/12/25
                Config.rpt_isCheckGraduateDomain = IsShowGraduateDomainSheet(); // 小郭, 2013/12/25

                IEvaluateFactory factory = new ConditionalEvaluateFactory(conditions);
                Graduation.Instance.SetFactory(factory);
                Graduation.Instance.Refresh();

                if (!_historyWorker.IsBusy)
                {
                    btnPrint.Enabled = false;
                    _historyWorker.RunWorkerAsync(_students, _errorList);
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("請勾選檢查條件。");
                return;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MergeResult(Dictionary<string, bool> passList, EvaluationResult sourceResult, EvaluationResult targetResult)
        {
            foreach (string student_id in sourceResult.Keys)
            {
                if (passList.ContainsKey(student_id) && passList[student_id]) continue;
                targetResult.MergeResults(student_id, sourceResult[student_id]);
            }
        }

        private void GraduationPredictReportForm_Load(object sender, EventArgs e)
        {
            // 學期
            cboSemester.MaxLength = 2;
            cboSemester.MaxLength = 1;
            cboSemester.Items.AddRange(new string[] { "1", "2" });
            cboSemester.Text = School.DefaultSemester;

            // 學年度
            cboSchoolYear.Text = School.DefaultSchoolYear;
            int sc;
            int.TryParse(School.DefaultSchoolYear, out sc);
            for (int i = sc - 2; i <= sc + 2; i++) cboSchoolYear.Items.Add(i.ToString());
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Object each in gpDaily.Controls)
            {
                if (each is System.Windows.Forms.CheckBox)
                {
                    System.Windows.Forms.CheckBox cb = each as System.Windows.Forms.CheckBox;
                    if (cb.Text.Contains("各學期"))
                    {
                        cb.Checked = checkBox2.Checked;
                    }
                }
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Object each in gpDaily.Controls)
            {
                if (each is System.Windows.Forms.CheckBox)
                {
                    System.Windows.Forms.CheckBox cb = each as System.Windows.Forms.CheckBox;
                    if (cb.Text.Contains("第六學期"))
                    {
                        cb.Checked = checkBox3.Checked;
                    }
                }
            }
        }

        private void ExportDoctSetup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 樣板與列印設定
            PrintConfigForm pcf = new PrintConfigForm();
            pcf.ShowDialog();
        }

        private void chConditionGr1_CheckedChanged(object sender, EventArgs e)
        {

        }

        #region 判斷工作表是否出現
        /// <summary>
        /// 是否顯示"未達畢業標準學生-依畢業總平均"的工作表, 小郭, 2013/12/25
        /// </summary>
        private bool IsShowGraduateDomainSheet()
        {
            // 假如有勾選"學習領域畢業總平均成績符合規範。"
            if (chConditionGr1.Checked && (!string.IsNullOrEmpty("" + chConditionGr1.Tag)))
                return true;

            return false;
        }

        /// <summary>
        /// 是否顯示"未達畢業標準學生"的工作表, 小郭, 2013/12/25
        /// </summary>
        private bool IsShowSemesterDomainSheet()
        {
            // 假如有勾選"各學期領域成績均符合規範。"
            if (chCondition1.Checked && (!string.IsNullOrEmpty("" + chCondition1.Tag)))
                return true;
            // 假如有勾選"第六學期各領域成績符合規範。"
            if (chCondition2.Checked && (!string.IsNullOrEmpty("" + chCondition2.Tag)))
                return true;

            // 只要"日常生活表現"其中一項有勾選就顯示
            foreach (Control ctrl in gpDaily.Controls)
            {
                if (ctrl is System.Windows.Forms.CheckBox)
                {
                    System.Windows.Forms.CheckBox chk = ctrl as System.Windows.Forms.CheckBox;
                    if (chk.Checked && !string.IsNullOrEmpty("" + chk.Tag))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        #endregion
    }
}
