using Aspose.Words;
using Campus.Report;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using JHScoreReportDAL;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HsinChu.StudentRecordReportVer2
{
    public partial class PrintForm_StudentRecordReport : BaseForm
    {
        internal const string _ConfigName = "StudentReport";

        private List<string> _StudentIDs { get; set; }

        private ReportPreference _Preference { get; set; }

        private BackgroundWorker _MasterWorker = new BackgroundWorker();

        private BackgroundWorker _ConvertToPDF_Worker = new BackgroundWorker();

        //領域科目資料管理設定的資料
        JHScoreReportDAL.Config _DomainSubjectConfig = new Config();
        List<string> _SubjectTemplateList = new List<string>();

        //上課節次設定(列數、不列入)
        List<K12.Data.PeriodMappingInfo> _PeriodMappingInfos = K12.Data.PeriodMapping.SelectAll();

        //缺曠的節次(一般)名稱
        List<string> _AbsencePeriod = new List<string>();

        //學生清單
        private List<K12.Data.StudentRecord> _PrintStudents = new List<K12.Data.StudentRecord>();

        //學生基本資料 [studentID, Data]
        Dictionary<string, K12.Data.StudentRecord> _StudentRecordDic = new Dictionary<string, K12.Data.StudentRecord>();

        //學生家長基本資料 [studentID, Data]
        Dictionary<string, K12.Data.ParentRecord> _StudentParentRecordDic = new Dictionary<string, K12.Data.ParentRecord>();

        //學生聯繫資料(住址) [studentID, Data]
        Dictionary<string, K12.Data.AddressRecord> _StudentAddressRecordDic = new Dictionary<string, K12.Data.AddressRecord>();

        //學生聯繫資料(電話) [studentID, Data]
        Dictionary<string, K12.Data.PhoneRecord> _StudentPhoneRecordDic = new Dictionary<string, K12.Data.PhoneRecord>();

        //學期歷程 [studentID, Data]
        Dictionary<string, K12.Data.SemesterHistoryRecord> _SemesterHistoryRecordDic = new Dictionary<string, K12.Data.SemesterHistoryRecord>();


        //缺曠 [studentID, List<Data>]
        Dictionary<string, List<K12.Data.AttendanceRecord>> _AttendanceRecordDic = new Dictionary<string, List<K12.Data.AttendanceRecord>>();

        //學期成績(領域、科目) [studentID, List<Data>]
        Dictionary<string, List<JHSchool.Data.JHSemesterScoreRecord>> _SemesterScoreRecordDic = new Dictionary<string, List<JHSchool.Data.JHSemesterScoreRecord>>();

        //畢業分數 [studentID, Data]
        Dictionary<string, K12.Data.GradScoreRecord> _GraduateScoreRecordDic = new Dictionary<string, K12.Data.GradScoreRecord>();

        //異動 [studentID, List<Data>]
        Dictionary<string, List<K12.Data.UpdateRecordRecord>> _UpdateRecordRecordDic = new Dictionary<string, List<K12.Data.UpdateRecordRecord>>();

        //日常生活表現、校內外特殊表現 [studentID, List<Data>]
        Dictionary<string, List<K12.Data.MoralScoreRecord>> _MoralScoreRecordDic = new Dictionary<string, List<K12.Data.MoralScoreRecord>>();

        //單檔列印時的資料夾路徑
        private string _FileBrowserDialogPath = "";

        private DoWorkEventArgs _E_For_ConvertToPDF_Worker;

        public PrintForm_StudentRecordReport(List<string> studentIDs)
        {
            InitializeComponent();

            _StudentIDs = studentIDs;
            _Preference = new ReportPreference(_ConfigName, Resources.新竹國中學籍表_樣板_);

            _MasterWorker.DoWork += new DoWorkEventHandler(MasterWorker_DoWork);
            _MasterWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(MasterWorker_RunWorkerCompleted);
            _MasterWorker.WorkerReportsProgress = true;
            _MasterWorker.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
            };

            _ConvertToPDF_Worker.DoWork += new DoWorkEventHandler(ConvertToPDF_Worker_DoWork);
            _ConvertToPDF_Worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ConvertToPDF_Worker_RunWorkerCompleted);
            _ConvertToPDF_Worker.WorkerReportsProgress = true;
            _ConvertToPDF_Worker.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                MotherForm.SetStatusBarMessage("" + e.UserState.ToString(), e.ProgressPercentage);
            };

            //是否列印PDF
            rtnPDF.Checked = _Preference.ConvertToPDF;

            //是否要單檔列印
            OneFileSave.Checked = _Preference.OneFileSave;
        }

        private void MasterWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 清除字典資料
            _SubjectTemplateList.Clear();
            _StudentRecordDic.Clear();
            _StudentParentRecordDic.Clear();
            _StudentAddressRecordDic.Clear();
            _StudentPhoneRecordDic.Clear();
            _SemesterHistoryRecordDic.Clear();
            _AttendanceRecordDic.Clear();
            _SemesterScoreRecordDic.Clear();
            _GraduateScoreRecordDic.Clear();
            _UpdateRecordRecordDic.Clear();
            _MoralScoreRecordDic.Clear();
            #endregion

            #region 抓取資料
            //取得領域科目資料管理設定的資料
            _DomainSubjectConfig.GetConfigData();
            foreach (ConfigItem item in _DomainSubjectConfig.GetSubjectItemList())
            {
                _SubjectTemplateList.Add(item.Name);
            }

            //抓取學生資料
            //學生基本資料
            List<K12.Data.StudentRecord> studentRecordList = K12.Data.Student.SelectByIDs(_StudentIDs);

            //學生家長基本資料
            List<K12.Data.ParentRecord> studentParentRecordList = K12.Data.Parent.SelectByStudentIDs(_StudentIDs);

            //學生聯繫資料(住址)
            List<K12.Data.AddressRecord> studentAddressRecordList = K12.Data.Address.SelectByStudentIDs(_StudentIDs);

            //學生聯繫資料(電話)
            List<K12.Data.PhoneRecord> studentPhoneRecoedList = K12.Data.Phone.SelectByStudentIDs(_StudentIDs);

            //學期歷程
            List<K12.Data.SemesterHistoryRecord> semesterHistoryRecordList = K12.Data.SemesterHistory.SelectByStudentIDs(_StudentIDs);

            //缺曠
            List<K12.Data.AttendanceRecord> attendRecordsList = K12.Data.Attendance.SelectByStudentIDs(_StudentIDs);

            //學期成績(包含領域、科目)
            List<JHSemesterScoreRecord> semesterScoreRecordList = JHSemesterScore.SelectByStudentIDs(_StudentIDs);

            //畢業分數
            List<K12.Data.GradScoreRecord> graduateScoreRecordList = K12.Data.GradScore.SelectByIDs<K12.Data.GradScoreRecord>(_StudentIDs);

            //異動
            List<K12.Data.UpdateRecordRecord> updateRecordRecordList = K12.Data.UpdateRecord.SelectByStudentIDs(_StudentIDs);

            //日常生活表現、校內外特殊表現
            List<K12.Data.MoralScoreRecord> moralScoreRecordList = K12.Data.MoralScore.SelectByStudentIDs(_StudentIDs);
            #endregion

            #region 整理資料
            //整理學生基本資料
            foreach (K12.Data.StudentRecord studentRecord in studentRecordList)
            {
                if (!_StudentRecordDic.ContainsKey(studentRecord.ID))
                {
                    _StudentRecordDic.Add(studentRecord.ID, studentRecord);
                }
            }

            //整理學生家長基本資料
            foreach (K12.Data.ParentRecord studentParentRecord in studentParentRecordList)
            {
                if (!_StudentParentRecordDic.ContainsKey(studentParentRecord.RefStudentID))
                {
                    _StudentParentRecordDic.Add(studentParentRecord.RefStudentID, studentParentRecord);
                }
            }

            //整理學生聯繫資料(住址)
            foreach (K12.Data.AddressRecord studentAddressRecord in studentAddressRecordList)
            {
                if (!_StudentAddressRecordDic.ContainsKey(studentAddressRecord.RefStudentID))
                {
                    _StudentAddressRecordDic.Add(studentAddressRecord.RefStudentID, studentAddressRecord);
                }
            }

            //整理學生聯繫資料(電話)
            foreach (K12.Data.PhoneRecord studentPhoneRecord in studentPhoneRecoedList)
            {
                if (!_StudentPhoneRecordDic.ContainsKey(studentPhoneRecord.RefStudentID))
                {
                    _StudentPhoneRecordDic.Add(studentPhoneRecord.RefStudentID, studentPhoneRecord);
                }
            }

            //整理學期歷程
            foreach (K12.Data.SemesterHistoryRecord semesterHistoryRecord in semesterHistoryRecordList)
            {
                if (!_SemesterHistoryRecordDic.ContainsKey(semesterHistoryRecord.RefStudentID))
                {
                    _SemesterHistoryRecordDic.Add(semesterHistoryRecord.RefStudentID, semesterHistoryRecord);
                }
            }

            //整理缺曠紀錄
            foreach (K12.Data.AttendanceRecord attendanceRecord in attendRecordsList)
            {
                if (!_AttendanceRecordDic.ContainsKey(attendanceRecord.RefStudentID))
                {
                    _AttendanceRecordDic.Add(attendanceRecord.RefStudentID, new List<K12.Data.AttendanceRecord>());
                    _AttendanceRecordDic[attendanceRecord.RefStudentID].Add(attendanceRecord);
                }
                else
                {
                    _AttendanceRecordDic[attendanceRecord.RefStudentID].Add(attendanceRecord);
                }
            }

            //整理學期成績(包含領域、科目)紀錄
            //將科目依照科目資料管理排序
            foreach (JHSemesterScoreRecord semesterScoreRecord in semesterScoreRecordList)
            {
                Dictionary<string, K12.Data.SubjectScore> subjectScoreDic = new Dictionary<string, K12.Data.SubjectScore>();
                List<string> subjectNameList = new List<string>();
                subjectNameList = semesterScoreRecord.Subjects.Keys.ToList();

                subjectNameList.Sort(new HsinChu.StudentRecordReportVer2.StringComparer(_SubjectTemplateList.ToArray()));

                foreach (string subjectName in subjectNameList)
                {
                    K12.Data.SubjectScore score = semesterScoreRecord.Subjects[subjectName];
                    subjectScoreDic.Add(subjectName, score);
                }

                semesterScoreRecord.Subjects = subjectScoreDic;
            }

            //整理學期成績
            foreach (JHSemesterScoreRecord semesterScoreRecord in semesterScoreRecordList)
            {
                if (!_SemesterScoreRecordDic.ContainsKey(semesterScoreRecord.RefStudentID))
                {
                    _SemesterScoreRecordDic.Add(semesterScoreRecord.RefStudentID, new List<JHSemesterScoreRecord>());
                    _SemesterScoreRecordDic[semesterScoreRecord.RefStudentID].Add(semesterScoreRecord);
                }
                else
                {
                    _SemesterScoreRecordDic[semesterScoreRecord.RefStudentID].Add(semesterScoreRecord);
                }
            }

            //整理畢業分數
            foreach (K12.Data.GradScoreRecord gradScoreRecord in graduateScoreRecordList)
            {
                if (!_GraduateScoreRecordDic.ContainsKey(gradScoreRecord.RefStudentID))
                {
                    _GraduateScoreRecordDic.Add(gradScoreRecord.RefStudentID, gradScoreRecord);
                }
            }

            //整理異動紀錄
            foreach (K12.Data.UpdateRecordRecord updateRecordRecord in updateRecordRecordList)
            {
                if (!_UpdateRecordRecordDic.ContainsKey(updateRecordRecord.StudentID))
                {
                    _UpdateRecordRecordDic.Add(updateRecordRecord.StudentID, new List<K12.Data.UpdateRecordRecord>());
                    _UpdateRecordRecordDic[updateRecordRecord.StudentID].Add(updateRecordRecord);
                }
                else
                {
                    _UpdateRecordRecordDic[updateRecordRecord.StudentID].Add(updateRecordRecord);
                }
            }

            //整理日常生活表現、校內外特殊表現
            foreach (K12.Data.MoralScoreRecord moralScoreRecord in moralScoreRecordList)
            {
                if (!_MoralScoreRecordDic.ContainsKey(moralScoreRecord.RefStudentID))
                {
                    _MoralScoreRecordDic.Add(moralScoreRecord.RefStudentID, new List<K12.Data.MoralScoreRecord>());
                    _MoralScoreRecordDic[moralScoreRecord.RefStudentID].Add(moralScoreRecord);
                }
                else
                {
                    _MoralScoreRecordDic[moralScoreRecord.RefStudentID].Add(moralScoreRecord);
                }
            }
            #endregion

            #region 建立合併欄位總表
            //建立合併欄位總表
            DataTable table = new DataTable();

            #region 基本資料
            //基本資料
            table.Columns.Add("學生姓名");
            table.Columns.Add("學生班級");
            table.Columns.Add("學生性別");
            table.Columns.Add("出生日期");
            table.Columns.Add("入學年月");
            table.Columns.Add("學生身分證字號");
            table.Columns.Add("學號");
            table.Columns.Add("戶籍地址");
            table.Columns.Add("戶籍電話");
            table.Columns.Add("聯絡地址");
            table.Columns.Add("聯絡電話");
            table.Columns.Add("學校名稱");
            #endregion

            #region 學年度
            //學年度
            table.Columns.Add("學年度1");
            table.Columns.Add("學年度2");
            table.Columns.Add("學年度3");
            #endregion

            #region 家長資料
            table.Columns.Add("監護人姓名");
            table.Columns.Add("監護人關係");
            table.Columns.Add("監護人行動電話");
            table.Columns.Add("父親姓名");
            table.Columns.Add("父親行動電話");
            table.Columns.Add("母親姓名");
            table.Columns.Add("母親行動電話");
            #endregion

            #region 學期歷程 班級、座號、班導師資料
            //班級座號資料
            table.Columns.Add("班級1");
            table.Columns.Add("座號1");
            table.Columns.Add("班導師1");

            table.Columns.Add("班級2");
            table.Columns.Add("座號2");
            table.Columns.Add("班導師2");

            table.Columns.Add("班級3");
            table.Columns.Add("座號3");
            table.Columns.Add("班導師3");

            table.Columns.Add("班級4");
            table.Columns.Add("座號4");
            table.Columns.Add("班導師4");

            table.Columns.Add("班級5");
            table.Columns.Add("座號5");
            table.Columns.Add("班導師5");

            table.Columns.Add("班級6");
            table.Columns.Add("座號6");
            table.Columns.Add("班導師6");
            #endregion

            #region 異動紀錄
            //異動紀錄
            table.Columns.Add("異動紀錄1_日期");
            table.Columns.Add("異動紀錄1_校名");
            table.Columns.Add("異動紀錄1_類別");
            table.Columns.Add("異動紀錄1_核准日期");
            table.Columns.Add("異動紀錄1_核准文號");
            table.Columns.Add("異動紀錄2_日期");
            table.Columns.Add("異動紀錄2_校名");
            table.Columns.Add("異動紀錄2_類別");
            table.Columns.Add("異動紀錄2_核准日期");
            table.Columns.Add("異動紀錄2_核准文號");
            table.Columns.Add("異動紀錄3_日期");
            table.Columns.Add("異動紀錄3_校名");
            table.Columns.Add("異動紀錄3_類別");
            table.Columns.Add("異動紀錄3_核准日期");
            table.Columns.Add("異動紀錄3_核准文號");
            table.Columns.Add("異動紀錄4_日期");
            table.Columns.Add("異動紀錄4_校名");
            table.Columns.Add("異動紀錄4_類別");
            table.Columns.Add("異動紀錄4_核准日期");
            table.Columns.Add("異動紀錄4_核准文號");
            table.Columns.Add("異動紀錄5_日期");
            table.Columns.Add("異動紀錄5_校名");
            table.Columns.Add("異動紀錄5_類別");
            table.Columns.Add("異動紀錄5_核准日期");
            table.Columns.Add("異動紀錄5_核准文號");
            table.Columns.Add("異動紀錄6_日期");
            table.Columns.Add("異動紀錄6_校名");
            table.Columns.Add("異動紀錄6_類別");
            table.Columns.Add("異動紀錄6_核准日期");
            table.Columns.Add("異動紀錄6_核准文號");
            table.Columns.Add("異動紀錄7_日期");
            table.Columns.Add("異動紀錄7_校名");
            table.Columns.Add("異動紀錄7_類別");
            table.Columns.Add("異動紀錄7_核准日期");
            table.Columns.Add("異動紀錄7_核准文號");
            table.Columns.Add("異動紀錄8_日期");
            table.Columns.Add("異動紀錄8_校名");
            table.Columns.Add("異動紀錄8_類別");
            table.Columns.Add("異動紀錄8_核准日期");
            table.Columns.Add("異動紀錄8_核准文號");
            table.Columns.Add("異動紀錄9_日期");
            table.Columns.Add("異動紀錄9_校名");
            table.Columns.Add("異動紀錄9_類別");
            table.Columns.Add("異動紀錄9_核准日期");
            table.Columns.Add("異動紀錄9_核准文號");
            table.Columns.Add("異動紀錄10_日期");
            table.Columns.Add("異動紀錄10_校名");
            table.Columns.Add("異動紀錄10_類別");
            table.Columns.Add("異動紀錄10_核准日期");
            table.Columns.Add("異動紀錄10_核准文號");
            #endregion

            #region 缺曠紀錄
            //缺曠紀錄
            table.Columns.Add("應出席日數_1");
            table.Columns.Add("應出席日數_2");
            table.Columns.Add("應出席日數_3");
            table.Columns.Add("應出席日數_4");
            table.Columns.Add("應出席日數_5");
            table.Columns.Add("應出席日數_6");


            table.Columns.Add("事假日數_1");
            table.Columns.Add("事假日數_2");
            table.Columns.Add("事假日數_3");
            table.Columns.Add("事假日數_4");
            table.Columns.Add("事假日數_5");
            table.Columns.Add("事假日數_6");


            table.Columns.Add("病假日數_1");
            table.Columns.Add("病假日數_2");
            table.Columns.Add("病假日數_3");
            table.Columns.Add("病假日數_4");
            table.Columns.Add("病假日數_5");
            table.Columns.Add("病假日數_6");

            table.Columns.Add("公假日數_1");
            table.Columns.Add("公假日數_2");
            table.Columns.Add("公假日數_3");
            table.Columns.Add("公假日數_4");
            table.Columns.Add("公假日數_5");
            table.Columns.Add("公假日數_6");

            table.Columns.Add("喪假日數_1");
            table.Columns.Add("喪假日數_2");
            table.Columns.Add("喪假日數_3");
            table.Columns.Add("喪假日數_4");
            table.Columns.Add("喪假日數_5");
            table.Columns.Add("喪假日數_6");

            table.Columns.Add("曠課日數_1");
            table.Columns.Add("曠課日數_2");
            table.Columns.Add("曠課日數_3");
            table.Columns.Add("曠課日數_4");
            table.Columns.Add("曠課日數_5");
            table.Columns.Add("曠課日數_6");

            table.Columns.Add("缺席總日數_1");
            table.Columns.Add("缺席總日數_2");
            table.Columns.Add("缺席總日數_3");
            table.Columns.Add("缺席總日數_4");
            table.Columns.Add("缺席總日數_5");
            table.Columns.Add("缺席總日數_6");
            #endregion

            #region 領域成績
            //領域成績
            table.Columns.Add("領域_語文_成績_1");
            table.Columns.Add("領域_語文_成績_2");
            table.Columns.Add("領域_語文_成績_3");
            table.Columns.Add("領域_語文_成績_4");
            table.Columns.Add("領域_語文_成績_5");
            table.Columns.Add("領域_語文_成績_6");


            table.Columns.Add("領域_語文_等第_1");
            table.Columns.Add("領域_語文_等第_2");
            table.Columns.Add("領域_語文_等第_3");
            table.Columns.Add("領域_語文_等第_4");
            table.Columns.Add("領域_語文_等第_5");
            table.Columns.Add("領域_語文_等第_6");


            table.Columns.Add("領域_數學_成績_1");
            table.Columns.Add("領域_數學_成績_2");
            table.Columns.Add("領域_數學_成績_3");
            table.Columns.Add("領域_數學_成績_4");
            table.Columns.Add("領域_數學_成績_5");
            table.Columns.Add("領域_數學_成績_6");


            table.Columns.Add("領域_數學_等第_1");
            table.Columns.Add("領域_數學_等第_2");
            table.Columns.Add("領域_數學_等第_3");
            table.Columns.Add("領域_數學_等第_4");
            table.Columns.Add("領域_數學_等第_5");
            table.Columns.Add("領域_數學_等第_6");


            table.Columns.Add("領域_生活課程_成績_1");
            table.Columns.Add("領域_生活課程_成績_2");
            table.Columns.Add("領域_生活課程_成績_3");
            table.Columns.Add("領域_生活課程_成績_4");
            table.Columns.Add("領域_生活課程_成績_5");
            table.Columns.Add("領域_生活課程_成績_6");


            table.Columns.Add("領域_生活課程_等第_1");
            table.Columns.Add("領域_生活課程_等第_2");
            table.Columns.Add("領域_生活課程_等第_3");
            table.Columns.Add("領域_生活課程_等第_4");
            table.Columns.Add("領域_生活課程_等第_5");
            table.Columns.Add("領域_生活課程_等第_6");


            table.Columns.Add("領域_自然科學_成績_1");
            table.Columns.Add("領域_自然科學_成績_2");
            table.Columns.Add("領域_自然科學_成績_3");
            table.Columns.Add("領域_自然科學_成績_4");
            table.Columns.Add("領域_自然科學_成績_5");
            table.Columns.Add("領域_自然科學_成績_6");


            table.Columns.Add("領域_自然科學_等第_1");
            table.Columns.Add("領域_自然科學_等第_2");
            table.Columns.Add("領域_自然科學_等第_3");
            table.Columns.Add("領域_自然科學_等第_4");
            table.Columns.Add("領域_自然科學_等第_5");
            table.Columns.Add("領域_自然科學_等第_6");


            table.Columns.Add("領域_科技_成績_1");
            table.Columns.Add("領域_科技_成績_2");
            table.Columns.Add("領域_科技_成績_3");
            table.Columns.Add("領域_科技_成績_4");
            table.Columns.Add("領域_科技_成績_5");
            table.Columns.Add("領域_科技_成績_6");


            table.Columns.Add("領域_科技_等第_1");
            table.Columns.Add("領域_科技_等第_2");
            table.Columns.Add("領域_科技_等第_3");
            table.Columns.Add("領域_科技_等第_4");
            table.Columns.Add("領域_科技_等第_5");
            table.Columns.Add("領域_科技_等第_6");


            table.Columns.Add("領域_藝術_成績_1");
            table.Columns.Add("領域_藝術_成績_2");
            table.Columns.Add("領域_藝術_成績_3");
            table.Columns.Add("領域_藝術_成績_4");
            table.Columns.Add("領域_藝術_成績_5");
            table.Columns.Add("領域_藝術_成績_6");


            table.Columns.Add("領域_藝術_等第_1");
            table.Columns.Add("領域_藝術_等第_2");
            table.Columns.Add("領域_藝術_等第_3");
            table.Columns.Add("領域_藝術_等第_4");
            table.Columns.Add("領域_藝術_等第_5");
            table.Columns.Add("領域_藝術_等第_6");


            table.Columns.Add("領域_自然與生活科技_成績_1");
            table.Columns.Add("領域_自然與生活科技_成績_2");
            table.Columns.Add("領域_自然與生活科技_成績_3");
            table.Columns.Add("領域_自然與生活科技_成績_4");
            table.Columns.Add("領域_自然與生活科技_成績_5");
            table.Columns.Add("領域_自然與生活科技_成績_6");


            table.Columns.Add("領域_自然與生活科技_等第_1");
            table.Columns.Add("領域_自然與生活科技_等第_2");
            table.Columns.Add("領域_自然與生活科技_等第_3");
            table.Columns.Add("領域_自然與生活科技_等第_4");
            table.Columns.Add("領域_自然與生活科技_等第_5");
            table.Columns.Add("領域_自然與生活科技_等第_6");


            table.Columns.Add("領域_藝術與人文_成績_1");
            table.Columns.Add("領域_藝術與人文_成績_2");
            table.Columns.Add("領域_藝術與人文_成績_3");
            table.Columns.Add("領域_藝術與人文_成績_4");
            table.Columns.Add("領域_藝術與人文_成績_5");
            table.Columns.Add("領域_藝術與人文_成績_6");


            table.Columns.Add("領域_藝術與人文_等第_1");
            table.Columns.Add("領域_藝術與人文_等第_2");
            table.Columns.Add("領域_藝術與人文_等第_3");
            table.Columns.Add("領域_藝術與人文_等第_4");
            table.Columns.Add("領域_藝術與人文_等第_5");
            table.Columns.Add("領域_藝術與人文_等第_6");



            table.Columns.Add("領域_社會_成績_1");
            table.Columns.Add("領域_社會_成績_2");
            table.Columns.Add("領域_社會_成績_3");
            table.Columns.Add("領域_社會_成績_4");
            table.Columns.Add("領域_社會_成績_5");
            table.Columns.Add("領域_社會_成績_6");


            table.Columns.Add("領域_社會_等第_1");
            table.Columns.Add("領域_社會_等第_2");
            table.Columns.Add("領域_社會_等第_3");
            table.Columns.Add("領域_社會_等第_4");
            table.Columns.Add("領域_社會_等第_5");
            table.Columns.Add("領域_社會_等第_6");


            table.Columns.Add("領域_健康與體育_成績_1");
            table.Columns.Add("領域_健康與體育_成績_2");
            table.Columns.Add("領域_健康與體育_成績_3");
            table.Columns.Add("領域_健康與體育_成績_4");
            table.Columns.Add("領域_健康與體育_成績_5");
            table.Columns.Add("領域_健康與體育_成績_6");


            table.Columns.Add("領域_健康與體育_等第_1");
            table.Columns.Add("領域_健康與體育_等第_2");
            table.Columns.Add("領域_健康與體育_等第_3");
            table.Columns.Add("領域_健康與體育_等第_4");
            table.Columns.Add("領域_健康與體育_等第_5");
            table.Columns.Add("領域_健康與體育_等第_6");


            table.Columns.Add("領域_綜合活動_成績_1");
            table.Columns.Add("領域_綜合活動_成績_2");
            table.Columns.Add("領域_綜合活動_成績_3");
            table.Columns.Add("領域_綜合活動_成績_4");
            table.Columns.Add("領域_綜合活動_成績_5");
            table.Columns.Add("領域_綜合活動_成績_6");


            table.Columns.Add("領域_綜合活動_等第_1");
            table.Columns.Add("領域_綜合活動_等第_2");
            table.Columns.Add("領域_綜合活動_等第_3");
            table.Columns.Add("領域_綜合活動_等第_4");
            table.Columns.Add("領域_綜合活動_等第_5");
            table.Columns.Add("領域_綜合活動_等第_6");

            table.Columns.Add("領域_彈性課程_成績_1");
            table.Columns.Add("領域_彈性課程_成績_2");
            table.Columns.Add("領域_彈性課程_成績_3");
            table.Columns.Add("領域_彈性課程_成績_4");
            table.Columns.Add("領域_彈性課程_成績_5");
            table.Columns.Add("領域_彈性課程_成績_6");


            table.Columns.Add("領域_彈性課程_等第_1");
            table.Columns.Add("領域_彈性課程_等第_2");
            table.Columns.Add("領域_彈性課程_等第_3");
            table.Columns.Add("領域_彈性課程_等第_4");
            table.Columns.Add("領域_彈性課程_等第_5");
            table.Columns.Add("領域_彈性課程_等第_6");


            table.Columns.Add("領域_學習領域總成績_成績_1");
            table.Columns.Add("領域_學習領域總成績_成績_2");
            table.Columns.Add("領域_學習領域總成績_成績_3");
            table.Columns.Add("領域_學習領域總成績_成績_4");
            table.Columns.Add("領域_學習領域總成績_成績_5");
            table.Columns.Add("領域_學習領域總成績_成績_6");


            table.Columns.Add("領域_學習領域總成績_等第_1");
            table.Columns.Add("領域_學習領域總成績_等第_2");
            table.Columns.Add("領域_學習領域總成績_等第_3");
            table.Columns.Add("領域_學習領域總成績_等第_4");
            table.Columns.Add("領域_學習領域總成績_等第_5");
            table.Columns.Add("領域_學習領域總成績_等第_6");

            table.Columns.Add("領域_課程學習成績_成績_1");
            table.Columns.Add("領域_課程學習成績_成績_2");
            table.Columns.Add("領域_課程學習成績_成績_3");
            table.Columns.Add("領域_課程學習成績_成績_4");
            table.Columns.Add("領域_課程學習成績_成績_5");
            table.Columns.Add("領域_課程學習成績_成績_6");


            table.Columns.Add("領域_課程學習成績_等第_1");
            table.Columns.Add("領域_課程學習成績_等第_2");
            table.Columns.Add("領域_課程學習成績_等第_3");
            table.Columns.Add("領域_課程學習成績_等第_4");
            table.Columns.Add("領域_課程學習成績_等第_5");
            table.Columns.Add("領域_課程學習成績_等第_6");
            #endregion

            #region 科目成績
            //科目成績
            table.Columns.Add("語文1_科目名稱");
            table.Columns.Add("語文2_科目名稱");
            table.Columns.Add("語文3_科目名稱");
            table.Columns.Add("語文4_科目名稱");
            table.Columns.Add("語文5_科目名稱");
            table.Columns.Add("語文6_科目名稱");
            table.Columns.Add("數學1_科目名稱");
            table.Columns.Add("數學2_科目名稱");
            table.Columns.Add("數學3_科目名稱");
            table.Columns.Add("數學4_科目名稱");
            table.Columns.Add("數學5_科目名稱");
            table.Columns.Add("數學6_科目名稱");
            table.Columns.Add("生活課程1_科目名稱");
            table.Columns.Add("生活課程2_科目名稱");
            table.Columns.Add("生活課程3_科目名稱");
            table.Columns.Add("生活課程4_科目名稱");
            table.Columns.Add("生活課程5_科目名稱");
            table.Columns.Add("生活課程6_科目名稱");
            table.Columns.Add("自然科學1_科目名稱");
            table.Columns.Add("自然科學2_科目名稱");
            table.Columns.Add("自然科學3_科目名稱");
            table.Columns.Add("自然科學4_科目名稱");
            table.Columns.Add("自然科學5_科目名稱");
            table.Columns.Add("自然科學6_科目名稱");
            table.Columns.Add("科技1_科目名稱");
            table.Columns.Add("科技2_科目名稱");
            table.Columns.Add("科技3_科目名稱");
            table.Columns.Add("科技4_科目名稱");
            table.Columns.Add("科技5_科目名稱");
            table.Columns.Add("科技6_科目名稱");
            table.Columns.Add("社會1_科目名稱");
            table.Columns.Add("社會2_科目名稱");
            table.Columns.Add("社會3_科目名稱");
            table.Columns.Add("社會4_科目名稱");
            table.Columns.Add("社會5_科目名稱");
            table.Columns.Add("社會6_科目名稱");
            table.Columns.Add("藝術1_科目名稱");
            table.Columns.Add("藝術2_科目名稱");
            table.Columns.Add("藝術3_科目名稱");
            table.Columns.Add("藝術4_科目名稱");
            table.Columns.Add("藝術5_科目名稱");
            table.Columns.Add("藝術6_科目名稱");
            table.Columns.Add("健康與體育1_科目名稱");
            table.Columns.Add("健康與體育2_科目名稱");
            table.Columns.Add("健康與體育3_科目名稱");
            table.Columns.Add("健康與體育4_科目名稱");
            table.Columns.Add("健康與體育5_科目名稱");
            table.Columns.Add("健康與體育6_科目名稱");
            table.Columns.Add("綜合活動1_科目名稱");
            table.Columns.Add("綜合活動2_科目名稱");
            table.Columns.Add("綜合活動3_科目名稱");
            table.Columns.Add("綜合活動4_科目名稱");
            table.Columns.Add("綜合活動5_科目名稱");
            table.Columns.Add("綜合活動6_科目名稱");
            table.Columns.Add("藝術與人文1_科目名稱");
            table.Columns.Add("藝術與人文2_科目名稱");
            table.Columns.Add("藝術與人文3_科目名稱");
            table.Columns.Add("藝術與人文4_科目名稱");
            table.Columns.Add("藝術與人文5_科目名稱");
            table.Columns.Add("藝術與人文6_科目名稱");
            table.Columns.Add("自然與生活科技1_科目名稱");
            table.Columns.Add("自然與生活科技2_科目名稱");
            table.Columns.Add("自然與生活科技3_科目名稱");
            table.Columns.Add("自然與生活科技4_科目名稱");
            table.Columns.Add("自然與生活科技5_科目名稱");
            table.Columns.Add("自然與生活科技6_科目名稱");

            table.Columns.Add("彈性課程1_科目名稱");
            table.Columns.Add("彈性課程2_科目名稱");
            table.Columns.Add("彈性課程3_科目名稱");
            table.Columns.Add("彈性課程4_科目名稱");
            table.Columns.Add("彈性課程5_科目名稱");
            table.Columns.Add("彈性課程6_科目名稱");
            table.Columns.Add("彈性課程7_科目名稱");
            table.Columns.Add("彈性課程8_科目名稱");
            table.Columns.Add("彈性課程9_科目名稱");
            table.Columns.Add("彈性課程10_科目名稱");
            #endregion

            #region 畢業總成績
            //畢業總成績
            table.Columns.Add("畢業總成績_平均");
            table.Columns.Add("畢業總成績_等第");
            table.Columns.Add("准予畢業");
            table.Columns.Add("發給修業證書");
            #endregion

            #region 日常生活表現及具體建議
            //日常生活表現及具體建議
            table.Columns.Add("日常生活表現及具體建議_1");
            table.Columns.Add("日常生活表現及具體建議_2");
            table.Columns.Add("日常生活表現及具體建議_3");
            table.Columns.Add("日常生活表現及具體建議_4");
            table.Columns.Add("日常生活表現及具體建議_5");
            table.Columns.Add("日常生活表現及具體建議_6");
            #endregion

            #region 校內外特殊表現
            //校內外特殊表現
            table.Columns.Add("校內外特殊表現_1");
            table.Columns.Add("校內外特殊表現_2");
            table.Columns.Add("校內外特殊表現_3");
            table.Columns.Add("校內外特殊表現_4");
            table.Columns.Add("校內外特殊表現_5");
            table.Columns.Add("校內外特殊表現_6");
            #endregion 
            #endregion

            Aspose.Words.Document document = new Document();

            _E_For_ConvertToPDF_Worker = e;

            #region 整理所有的假別
            //整理所有的假別
            List<string> absenceType_list = new List<string>();
            absenceType_list.Add("事假");
            absenceType_list.Add("病假");
            absenceType_list.Add("公假");
            absenceType_list.Add("喪假");
            absenceType_list.Add("曠課");
            absenceType_list.Add("缺席總");
            #endregion

            #region 整理所有的領域_OO_成績
            //整理所有的領域_OO_成績
            List<string> domainScoreType_list = new List<string>();
            domainScoreType_list.Add("領域_語文_成績_");
            domainScoreType_list.Add("領域_數學_成績_");
            domainScoreType_list.Add("領域_生活課程_成績_");
            domainScoreType_list.Add("領域_自然科學_成績_");
            domainScoreType_list.Add("領域_科技_成績_");
            domainScoreType_list.Add("領域_藝術_成績_");
            domainScoreType_list.Add("領域_自然與生活科技_成績_");
            domainScoreType_list.Add("領域_藝術與人文_成績_");
            domainScoreType_list.Add("領域_社會_成績_");
            domainScoreType_list.Add("領域_健康與體育_成績_");
            domainScoreType_list.Add("領域_綜合活動_成績_");
            domainScoreType_list.Add("領域_學習領域總成績_成績_");
            domainScoreType_list.Add("領域_課程學習成績_成績_");
            #endregion

            #region 整理所有的領域_OO_等第
            //整理所有的領域_OO_等第
            List<string> domainLevelType_list = new List<string>();
            domainLevelType_list.Add("領域_語文_等第_");
            domainLevelType_list.Add("領域_數學_等第_");
            domainLevelType_list.Add("領域_生活課程_等第_");
            domainScoreType_list.Add("領域_自然科學_等第_");
            domainScoreType_list.Add("領域_科技_等第_");
            domainScoreType_list.Add("領域_藝術_等第_");
            domainLevelType_list.Add("領域_自然與生活科技_等第_");
            domainLevelType_list.Add("領域_藝術與人文_等第_");
            domainLevelType_list.Add("領域_社會_等第_");
            domainLevelType_list.Add("領域_健康與體育_等第_");
            domainLevelType_list.Add("領域_綜合活動_等第_");
            domainLevelType_list.Add("領域_學習領域總成績_等第_");
            domainLevelType_list.Add("領域_課程學習成績_等第_");
            #endregion

            #region 整理科目_OO_成績
            //整理科目_OO_成績
            List<string> subjectScoreType_list = new List<string>();
            subjectScoreType_list.Add("語文1_科目成績_");
            subjectScoreType_list.Add("語文2_科目成績_");
            subjectScoreType_list.Add("語文3_科目成績_");
            subjectScoreType_list.Add("語文4_科目成績_");
            subjectScoreType_list.Add("語文5_科目成績_");
            subjectScoreType_list.Add("語文6_科目成績_");
            subjectScoreType_list.Add("數學1_科目成績_");
            subjectScoreType_list.Add("數學2_科目成績_");
            subjectScoreType_list.Add("數學3_科目成績_");
            subjectScoreType_list.Add("數學4_科目成績_");
            subjectScoreType_list.Add("數學5_科目成績_");
            subjectScoreType_list.Add("數學6_科目成績_");
            subjectScoreType_list.Add("生活課程1_科目成績_");
            subjectScoreType_list.Add("生活課程2_科目成績_");
            subjectScoreType_list.Add("生活課程3_科目成績_");
            subjectScoreType_list.Add("生活課程4_科目成績_");
            subjectScoreType_list.Add("生活課程5_科目成績_");
            subjectScoreType_list.Add("生活課程6_科目成績_");
            subjectScoreType_list.Add("自然科學1_科目成績_");
            subjectScoreType_list.Add("自然科學2_科目成績_");
            subjectScoreType_list.Add("自然科學3_科目成績_");
            subjectScoreType_list.Add("自然科學4_科目成績_");
            subjectScoreType_list.Add("自然科學5_科目成績_");
            subjectScoreType_list.Add("自然科學6_科目成績_");
            subjectScoreType_list.Add("科技1_科目成績_");
            subjectScoreType_list.Add("科技2_科目成績_");
            subjectScoreType_list.Add("科技3_科目成績_");
            subjectScoreType_list.Add("科技4_科目成績_");
            subjectScoreType_list.Add("科技5_科目成績_");
            subjectScoreType_list.Add("科技6_科目成績_");
            subjectScoreType_list.Add("社會1_科目成績_");
            subjectScoreType_list.Add("社會2_科目成績_");
            subjectScoreType_list.Add("社會3_科目成績_");
            subjectScoreType_list.Add("社會4_科目成績_");
            subjectScoreType_list.Add("社會5_科目成績_");
            subjectScoreType_list.Add("社會6_科目成績_");
            subjectScoreType_list.Add("藝術1_科目成績_");
            subjectScoreType_list.Add("藝術2_科目成績_");
            subjectScoreType_list.Add("藝術3_科目成績_");
            subjectScoreType_list.Add("藝術4_科目成績_");
            subjectScoreType_list.Add("藝術5_科目成績_");
            subjectScoreType_list.Add("藝術6_科目成績_");
            subjectScoreType_list.Add("健康與體育1_科目成績_");
            subjectScoreType_list.Add("健康與體育2_科目成績_");
            subjectScoreType_list.Add("健康與體育3_科目成績_");
            subjectScoreType_list.Add("健康與體育4_科目成績_");
            subjectScoreType_list.Add("健康與體育5_科目成績_");
            subjectScoreType_list.Add("健康與體育6_科目成績_");
            subjectScoreType_list.Add("綜合活動1_科目成績_");
            subjectScoreType_list.Add("綜合活動2_科目成績_");
            subjectScoreType_list.Add("綜合活動3_科目成績_");
            subjectScoreType_list.Add("綜合活動4_科目成績_");
            subjectScoreType_list.Add("綜合活動5_科目成績_");
            subjectScoreType_list.Add("綜合活動6_科目成績_");
            subjectScoreType_list.Add("藝術與人文1_科目成績_");
            subjectScoreType_list.Add("藝術與人文2_科目成績_");
            subjectScoreType_list.Add("藝術與人文3_科目成績_");
            subjectScoreType_list.Add("藝術與人文4_科目成績_");
            subjectScoreType_list.Add("藝術與人文5_科目成績_");
            subjectScoreType_list.Add("藝術與人文6_科目成績_");
            subjectScoreType_list.Add("自然與生活科技1_科目成績_");
            subjectScoreType_list.Add("自然與生活科技2_科目成績_");
            subjectScoreType_list.Add("自然與生活科技3_科目成績_");
            subjectScoreType_list.Add("自然與生活科技4_科目成績_");
            subjectScoreType_list.Add("自然與生活科技5_科目成績_");
            subjectScoreType_list.Add("自然與生活科技6_科目成績_");
            subjectScoreType_list.Add("彈性課程1_科目成績_");
            subjectScoreType_list.Add("彈性課程2_科目成績_");
            subjectScoreType_list.Add("彈性課程3_科目成績_");
            subjectScoreType_list.Add("彈性課程4_科目成績_");
            subjectScoreType_list.Add("彈性課程5_科目成績_");
            subjectScoreType_list.Add("彈性課程6_科目成績_");
            subjectScoreType_list.Add("彈性課程7_科目成績_");
            subjectScoreType_list.Add("彈性課程8_科目成績_");
            subjectScoreType_list.Add("彈性課程9_科目成績_");
            subjectScoreType_list.Add("彈性課程10_科目成績_");
            #endregion

            #region 整理科目_OO_等第
            //整理科目_OO_等第
            List<string> subjectLevelType_list = new List<string>();
            subjectLevelType_list.Add("語文1_科目等第_");
            subjectLevelType_list.Add("語文2_科目等第_");
            subjectLevelType_list.Add("語文3_科目等第_");
            subjectLevelType_list.Add("語文4_科目等第_");
            subjectLevelType_list.Add("語文5_科目等第_");
            subjectLevelType_list.Add("語文6_科目等第_");
            subjectLevelType_list.Add("數學1_科目等第_");
            subjectLevelType_list.Add("數學2_科目等第_");
            subjectLevelType_list.Add("數學3_科目等第_");
            subjectLevelType_list.Add("數學4_科目等第_");
            subjectLevelType_list.Add("數學5_科目等第_");
            subjectLevelType_list.Add("數學6_科目等第_");
            subjectLevelType_list.Add("生活課程1_科目等第_");
            subjectLevelType_list.Add("生活課程2_科目等第_");
            subjectLevelType_list.Add("生活課程3_科目等第_");
            subjectLevelType_list.Add("生活課程4_科目等第_");
            subjectLevelType_list.Add("生活課程5_科目等第_");
            subjectLevelType_list.Add("生活課程6_科目等第_");
            subjectLevelType_list.Add("自然科學1_科目等第_");
            subjectLevelType_list.Add("自然科學2_科目等第_");
            subjectLevelType_list.Add("自然科學3_科目等第_");
            subjectLevelType_list.Add("自然科學4_科目等第_");
            subjectLevelType_list.Add("自然科學5_科目等第_");
            subjectLevelType_list.Add("自然科學6_科目等第_");
            subjectLevelType_list.Add("科技1_科目等第_");
            subjectLevelType_list.Add("科技2_科目等第_");
            subjectLevelType_list.Add("科技3_科目等第_");
            subjectLevelType_list.Add("科技4_科目等第_");
            subjectLevelType_list.Add("科技5_科目等第_");
            subjectLevelType_list.Add("科技6_科目等第_");
            subjectLevelType_list.Add("社會1_科目等第_");
            subjectLevelType_list.Add("社會2_科目等第_");
            subjectLevelType_list.Add("社會3_科目等第_");
            subjectLevelType_list.Add("社會4_科目等第_");
            subjectLevelType_list.Add("社會5_科目等第_");
            subjectLevelType_list.Add("社會6_科目等第_");
            subjectLevelType_list.Add("藝術1_科目等第_");
            subjectLevelType_list.Add("藝術2_科目等第_");
            subjectLevelType_list.Add("藝術3_科目等第_");
            subjectLevelType_list.Add("藝術4_科目等第_");
            subjectLevelType_list.Add("藝術5_科目等第_");
            subjectLevelType_list.Add("藝術6_科目等第_");
            subjectLevelType_list.Add("健康與體育1_科目等第_");
            subjectLevelType_list.Add("健康與體育2_科目等第_");
            subjectLevelType_list.Add("健康與體育3_科目等第_");
            subjectLevelType_list.Add("健康與體育4_科目等第_");
            subjectLevelType_list.Add("健康與體育5_科目等第_");
            subjectLevelType_list.Add("健康與體育6_科目等第_");
            subjectLevelType_list.Add("綜合活動1_科目等第_");
            subjectLevelType_list.Add("綜合活動2_科目等第_");
            subjectLevelType_list.Add("綜合活動3_科目等第_");
            subjectLevelType_list.Add("綜合活動4_科目等第_");
            subjectLevelType_list.Add("綜合活動5_科目等第_");
            subjectLevelType_list.Add("綜合活動6_科目等第_");
            subjectLevelType_list.Add("藝術與人文1_科目等第_");
            subjectLevelType_list.Add("藝術與人文2_科目等第_");
            subjectLevelType_list.Add("藝術與人文3_科目等第_");
            subjectLevelType_list.Add("藝術與人文4_科目等第_");
            subjectLevelType_list.Add("藝術與人文5_科目等第_");
            subjectLevelType_list.Add("藝術與人文6_科目等第_");
            subjectLevelType_list.Add("自然與生活科技1_科目等第_");
            subjectLevelType_list.Add("自然與生活科技2_科目等第_");
            subjectLevelType_list.Add("自然與生活科技3_科目等第_");
            subjectLevelType_list.Add("自然與生活科技4_科目等第_");
            subjectLevelType_list.Add("自然與生活科技5_科目等第_");
            subjectLevelType_list.Add("自然與生活科技6_科目等第_");
            subjectLevelType_list.Add("彈性課程1_科目等第_");
            subjectLevelType_list.Add("彈性課程2_科目等第_");
            subjectLevelType_list.Add("彈性課程3_科目等第_");
            subjectLevelType_list.Add("彈性課程4_科目等第_");
            subjectLevelType_list.Add("彈性課程5_科目等第_");
            subjectLevelType_list.Add("彈性課程6_科目等第_");
            subjectLevelType_list.Add("彈性課程7_科目等第_");
            subjectLevelType_list.Add("彈性課程8_科目等第_");
            subjectLevelType_list.Add("彈性課程9_科目等第_");
            subjectLevelType_list.Add("彈性課程10_科目等第_");
            #endregion

            // 領域分數、等第 的對照
            Dictionary<string, decimal?> domainScore_dict = new Dictionary<string, decimal?>();
            Dictionary<string, string> domainLevel_dict = new Dictionary<string, string>();

            // 科目分數、等第 的對照
            Dictionary<string, decimal?> subjectScore_dict = new Dictionary<string, decimal?>();
            Dictionary<string, string> subjectLevel_dict = new Dictionary<string, string>();

            // 缺曠節次 、日數 的對照
            Dictionary<string, decimal> arStatistic_dict = new Dictionary<string, decimal>();
            Dictionary<string, decimal> arStatistic_dict_days = new Dictionary<string, decimal>();

            //文字評量(日常生活表現及具體建議、校內外特殊表現)的對照
            Dictionary<string, string> textScore_dict = new Dictionary<string, string>();

            //上課節次設定(列數、不列入)
            foreach (K12.Data.PeriodMappingInfo var in _PeriodMappingInfos)
            {
                if (var.Type == "一般" & !_AbsencePeriod.Contains(var.Name))
                    _AbsencePeriod.Add(var.Name);
            }

            int student_counter = 1;

            foreach (string stuID in _StudentIDs)
            {
                //把每一筆資料的字典都清乾淨，避免資料汙染
                arStatistic_dict.Clear();
                arStatistic_dict_days.Clear();
                domainScore_dict.Clear();
                domainLevel_dict.Clear();
                subjectScore_dict.Clear();
                subjectLevel_dict.Clear();
                textScore_dict.Clear();

                // 建立缺曠 對照字典
                foreach (string ab in absenceType_list)
                {
                    for (int i = 1; i <= 6; i++)
                    {
                        arStatistic_dict.Add(ab + "日數_" + i, 0);
                    }
                }

                // 建立領域成績 對照字典
                foreach (string dst in domainScoreType_list)
                {
                    for (int i = 1; i <= 6; i++)
                    {
                        domainScore_dict.Add(dst + i, null);
                    }
                }

                // 建立領域等第 對照字典
                foreach (string dlt in domainLevelType_list)
                {
                    for (int i = 1; i <= 6; i++)
                    {
                        domainLevel_dict.Add(dlt + i, null);
                    }
                }

                // 建立科目成績 對照字典
                foreach (string sst in subjectScoreType_list)
                {
                    for (int i = 1; i <= 6; i++)
                    {
                        subjectScore_dict.Add(sst + i, null);

                        if (!table.Columns.Contains(sst + i))
                        {
                            table.Columns.Add(sst + i);
                        }

                    }
                }

                // 建立科目等第 對照字典
                foreach (string slt in subjectLevelType_list)
                {
                    for (int i = 1; i <= 6; i++)
                    {
                        subjectLevel_dict.Add(slt + i, null);

                        if (!table.Columns.Contains(slt + i))
                        {
                            table.Columns.Add(slt + i);
                        }
                    }
                }

                // 建立文字評量 對照字典
                for (int i = 1; i <= 6; i++)
                {
                    textScore_dict.Add("日常生活表現及具體建議_" + i, null);
                    textScore_dict.Add("校內外特殊表現_" + i, null);
                }

                // 存放 各年級與 學年的對照變數
                int schoolyear_grade1 = 0;
                int schoolyear_grade2 = 0;
                int schoolyear_grade3 = 0;

                DataRow row = table.NewRow();

                //學生基本資料
                if (_StudentRecordDic.ContainsKey(stuID))
                {
                    DateTime birthday = new DateTime();

                    row["學生姓名"] = _StudentRecordDic[stuID].Name;
                    row["學生班級"] = _StudentRecordDic[stuID].Class != null ? _StudentRecordDic[stuID].Class.Name : "";
                    row["學生性別"] = _StudentRecordDic[stuID].Gender;

                    if (_StudentRecordDic[stuID].Birthday != null)
                    {
                        birthday = (DateTime)_StudentRecordDic[stuID].Birthday;
                        // 轉換出生時間 成 2005/09/06 的格式
                        row["出生日期"] = birthday.ToString("yyyy/MM/dd");
                    }
                    else
                    {
                        row["出生日期"] = "";
                    }

                    row["入學年月"] = "";
                    row["學生身分證字號"] = _StudentRecordDic[stuID].IDNumber;
                    row["學號"] = _StudentRecordDic[stuID].StudentNumber;
                    row["學校名稱"] = K12.Data.School.ChineseName;

                    _PrintStudents.Add(_StudentRecordDic[stuID]);
                }

                //學生家長基本資料
                if (_StudentParentRecordDic.ContainsKey(stuID))
                {
                    row["監護人姓名"] = _StudentParentRecordDic[stuID].CustodianName;
                    row["監護人關係"] = _StudentParentRecordDic[stuID].CustodianRelationship;
                    row["監護人行動電話"] = _StudentParentRecordDic[stuID].CustodianPhone;
                    row["父親姓名"] = _StudentParentRecordDic[stuID].FatherName;
                    row["父親行動電話"] = _StudentParentRecordDic[stuID].FatherPhone;
                    row["母親姓名"] = _StudentParentRecordDic[stuID].MotherName; ;
                    row["母親行動電話"] = _StudentParentRecordDic[stuID].MotherPhone;
                }

                //學生聯繫資料(住址)
                if (_StudentAddressRecordDic.ContainsKey(stuID))
                {
                    row["戶籍地址"] = _StudentAddressRecordDic[stuID].PermanentAddress;
                    row["聯絡地址"] = _StudentAddressRecordDic[stuID].MailingAddress;
                }

                //學生聯繫資料(電話)
                if (_StudentPhoneRecordDic.ContainsKey(stuID))
                {
                    row["戶籍電話"] = _StudentPhoneRecordDic[stuID].Permanent;
                    row["聯絡電話"] = _StudentPhoneRecordDic[stuID].Contact;
                }

                //學期歷程
                if (_SemesterHistoryRecordDic.ContainsKey(stuID))
                {
                    foreach (var item in _SemesterHistoryRecordDic[stuID].SemesterHistoryItems)
                    {
                        if (item.GradeYear == 1)
                        {
                            row["學年度1"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade1 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_1"] = item.SchoolDayCount;
                                row["班級1"] = item.ClassName;
                                row["座號1"] = item.SeatNo;
                                row["班導師1"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_2"] = item.SchoolDayCount;
                                row["班級2"] = item.ClassName;
                                row["座號2"] = item.SeatNo;
                                row["班導師2"] = item.Teacher;
                            }
                        }
                        if (item.GradeYear == 2)
                        {
                            row["學年度2"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade2 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_3"] = item.SchoolDayCount;
                                row["班級3"] = item.ClassName;
                                row["座號3"] = item.SeatNo;
                                row["班導師3"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_4"] = item.SchoolDayCount;
                                row["班級4"] = item.ClassName;
                                row["座號4"] = item.SeatNo;
                                row["班導師4"] = item.Teacher;
                            }
                        }
                        if (item.GradeYear == 3)
                        {
                            row["學年度3"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade3 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_5"] = item.SchoolDayCount;
                                row["班級5"] = item.ClassName;
                                row["座號5"] = item.SeatNo;
                                row["班導師5"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_6"] = item.SchoolDayCount;
                                row["班級6"] = item.ClassName;
                                row["座號6"] = item.SeatNo;
                                row["班導師6"] = item.Teacher;
                            }
                        }
                    }
                }

                //學年度與年級的對照字典
                Dictionary<int, int> schoolyear_grade_dict = new Dictionary<int, int>();

                schoolyear_grade_dict.Add(1, schoolyear_grade1);
                schoolyear_grade_dict.Add(2, schoolyear_grade2);
                schoolyear_grade_dict.Add(3, schoolyear_grade3);

                //出缺勤
                if (_AttendanceRecordDic.ContainsKey(stuID))
                {
                    for (int grade = 1; grade <= 3; grade++)
                    {
                        foreach (var ar in _AttendanceRecordDic[stuID])
                        {
                            if (ar.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                if (ar.Semester == 1)
                                {
                                    foreach (var detail in ar.PeriodDetail)
                                    {
                                        // 假如該缺曠結束 沒有在 節次管理 設定為一般則跳過
                                        if (!_AbsencePeriod.Contains(detail.Period))
                                        {
                                            continue;
                                        }
                                        if (arStatistic_dict.ContainsKey(detail.AbsenceType + "日數_" + (grade * 2 - 1)))
                                        {
                                            //加一節，整學期節次與日數的關係，再最後再結算
                                            arStatistic_dict[detail.AbsenceType + "日數_" + (grade * 2 - 1)] += 1;

                                            // 不管是啥缺席，缺席總日數都加一節
                                            arStatistic_dict["缺席總日數_" + (grade * 2 - 1)] += 1;
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (var detail in ar.PeriodDetail)
                                    {
                                        // 假如該缺曠結束 沒有在 節次管理 設定為一般則跳過
                                        if (!_AbsencePeriod.Contains(detail.Period))
                                        {
                                            continue;
                                        }

                                        if (arStatistic_dict.ContainsKey(detail.AbsenceType + "日數_" + grade * 2))
                                        {

                                            //加一節，整學期節次與日數的關係，再最後再結算
                                            arStatistic_dict[detail.AbsenceType + "日數_" + grade * 2] += 1;

                                            // 不管是啥缺席，缺席總日數都加一節
                                            arStatistic_dict["缺席總日數_" + (grade * 2)] += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (string key in arStatistic_dict.Keys)
                    {
                        arStatistic_dict_days.Add(key, arStatistic_dict[key]);
                    }

                    //真正的填值，填日數，所以要做節次轉換
                    foreach (string key in arStatistic_dict_days.Keys)
                    {
                        // 一天幾節課
                        int periodsADay = _AbsencePeriod.Count;

                        row[key] = Math.Round(arStatistic_dict_days[key] / periodsADay, 2);
                    }
                }

                // 學期成績(包含領域、科目)
                //一般科目 科目名稱與科目編號對照表
                Dictionary<string, Dictionary<string, int>> SubjectCourseDict = new Dictionary<string, Dictionary<string, int>>()
                {
                    { "語文", new Dictionary<string, int>() }
                    , { "數學", new Dictionary<string, int>() }
                    , { "生活課程", new Dictionary<string, int>() }
                    , { "自然科學", new Dictionary<string, int>() }
                    , { "科技", new Dictionary<string, int>() }
                    , { "社會", new Dictionary<string, int>() }
                    , { "藝術", new Dictionary<string, int>() }
                    , { "健康與體育", new Dictionary<string, int>() }
                    , { "綜合活動", new Dictionary<string, int>() }
                    , { "藝術與人文", new Dictionary<string, int>() }
                    , { "自然與生活科技", new Dictionary<string, int>() }
                };

                if (_SemesterScoreRecordDic.ContainsKey(stuID))
                {
                    // 任一領域的科目數量是否超過
                    bool isExceed = false;

                    for (int grade = 1; grade <= 3; grade++)
                    {
                        foreach (JHSemesterScoreRecord semesterScoreRecord in _SemesterScoreRecordDic[stuID])
                        {
                            if (semesterScoreRecord.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                foreach (var subjectscore in semesterScoreRecord.Subjects)
                                {
                                    // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                    if (subjectscore.Value.Domain != "彈性課程" && subjectscore.Value.Domain != "彈性學習" && subjectscore.Value.Domain != "")
                                    {
                                        if (SubjectCourseDict.ContainsKey(subjectscore.Value.Domain))
                                        {
                                            int subjectCourseCount = SubjectCourseDict[subjectscore.Value.Domain].Count;

                                            if (SubjectCourseDict[subjectscore.Value.Domain].ContainsKey(subjectscore.Value.Subject))
                                            {
                                                continue;
                                            }

                                            subjectCourseCount++;

                                            // 目前僅支援 一個學生六學年之中同一領域僅能有 6個科目
                                            if (subjectCourseCount > 6)
                                            {
                                                isExceed = true;
                                                continue;
                                            }

                                            row[subjectscore.Value.Domain + subjectCourseCount + "_科目名稱"] = subjectscore.Value.Subject;

                                            SubjectCourseDict[subjectscore.Value.Domain].Add(subjectscore.Value.Subject, subjectCourseCount);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (isExceed)
                    {
                        MessageBox.Show("科目數超過報表變數可支援數量，超過的將不會顯示在學籍表中");
                    }
                }

                // 彈性課程 科目名稱 與彈性課程編號的對照
                Dictionary<string, int> AlternativeCourseDict = new Dictionary<string, int>();

                // 先統計 該學生 在全學年間 有的 彈性課程科目
                if (_SemesterScoreRecordDic.ContainsKey(stuID))
                {
                    // 彈性課程記數
                    int AlternativeCourse = 0;

                    for (int grade = 1; grade <= 3; grade++)
                    {
                        foreach (JHSemesterScoreRecord semesterScoreRecord in _SemesterScoreRecordDic[stuID])
                        {
                            if (semesterScoreRecord.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                foreach (var subjectscore in semesterScoreRecord.Subjects)
                                {
                                    // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                    if (subjectscore.Value.Domain == "彈性課程" || subjectscore.Value.Domain == "彈性學習" || subjectscore.Value.Domain == "")
                                    {
                                        // 對照科目名稱如果已經有，跳過
                                        if (AlternativeCourseDict.ContainsKey(subjectscore.Value.Subject))
                                        {
                                            continue;
                                        }

                                        AlternativeCourse++;

                                        // 目前僅先支援 一個學生在六年之中有 10個 彈性課程
                                        if (AlternativeCourse > 10)
                                        {
                                            MessageBox.Show("彈性科目數超過可支援數量，超過的將不會顯示在學籍表中");
                                            break;
                                        }

                                        row["彈性課程" + AlternativeCourse + "_科目名稱"] = subjectscore.Value.Subject;

                                        AlternativeCourseDict.Add(subjectscore.Value.Subject, AlternativeCourse);
                                    }
                                }
                            }
                        }
                    }
                }

                if (_SemesterScoreRecordDic.ContainsKey(stuID))
                {
                    for (int grade = 1; grade <= 3; grade++)
                    {
                        foreach (JHSemesterScoreRecord semesterScoreRecord in _SemesterScoreRecordDic[stuID])
                        {
                            if (semesterScoreRecord.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                if (semesterScoreRecord.Semester == 1)
                                {
                                    //領域
                                    foreach (var domainscore in semesterScoreRecord.Domains)
                                    {
                                        //紀錄成績
                                        if (domainScore_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2 - 1)))
                                        {
                                            domainScore_dict["領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2 - 1)] = domainscore.Value.Score;
                                        }

                                        //換算等第
                                        if (domainLevel_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2 - 1)))
                                        {
                                            domainLevel_dict["領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2 - 1)] = ScoreToLevel(domainscore.Value.Score);
                                        }
                                    }

                                    //科目
                                    foreach (var subjectscore in semesterScoreRecord.Subjects)
                                    {
                                        // 彈性課程記數
                                        int AlternativeCourse = 0;
                                        int SubjectCourseNum = 0;

                                        // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                        if (subjectscore.Value.Domain == "彈性課程" || subjectscore.Value.Domain == "彈性學習" || subjectscore.Value.Domain == "")
                                        {
                                            if (AlternativeCourseDict.ContainsKey(subjectscore.Value.Subject))
                                            {
                                                AlternativeCourse = AlternativeCourseDict[subjectscore.Value.Subject];

                                                //紀錄成績
                                                if (subjectScore_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2 - 1)))
                                                {

                                                    subjectScore_dict["彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                }

                                                //紀錄等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2 - 1)))
                                                {
                                                    subjectLevel_dict["彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2 - 1)] = ScoreToLevel(subjectscore.Value.Score);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (SubjectCourseDict.ContainsKey(subjectscore.Value.Domain))
                                            {
                                                if (SubjectCourseDict[subjectscore.Value.Domain].ContainsKey(subjectscore.Value.Subject))
                                                {
                                                    SubjectCourseNum = SubjectCourseDict[subjectscore.Value.Domain][subjectscore.Value.Subject];

                                                    //紀錄成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2 - 1)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                    }

                                                    //換算等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2 - 1)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2 - 1)] = ScoreToLevel(subjectscore.Value.Score);
                                                    }
                                                }
                                            }

                                        }

                                    }

                                    //學期學習領域(七大)成績(不包括彈性課程成績)
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_學習領域總成績_成績_" + (grade * 2 - 1)))
                                    {
                                        domainScore_dict["領域_學習領域總成績_成績_" + (grade * 2 - 1)] = semesterScoreRecord.LearnDomainScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_學習領域總成績_等第_" + (grade * 2 - 1)))
                                    {
                                        domainLevel_dict["領域_學習領域總成績_等第_" + (grade * 2 - 1)] = ScoreToLevel(semesterScoreRecord.LearnDomainScore);
                                    }

                                    //課程學習成績(包括彈性課程成績)
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_課程學習成績_成績_" + (grade * 2 - 1)))
                                    {
                                        domainScore_dict["領域_課程學習成績_成績_" + (grade * 2 - 1)] = semesterScoreRecord.CourseLearnScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_課程學習成績_等第_" + (grade * 2 - 1)))
                                    {
                                        domainLevel_dict["領域_課程學習成績_等第_" + (grade * 2 - 1)] = ScoreToLevel(semesterScoreRecord.CourseLearnScore);
                                    }


                                }
                                else
                                {
                                    //領域
                                    foreach (var domainscore in semesterScoreRecord.Domains)
                                    {
                                        if (domainScore_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2)))
                                        {
                                            domainScore_dict["領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2)] = domainscore.Value.Score;
                                        }

                                        //換算等第
                                        if (domainLevel_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2)))
                                        {
                                            domainLevel_dict["領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2)] = ScoreToLevel(domainscore.Value.Score);
                                        }
                                    }

                                    //科目
                                    foreach (var subjectscore in semesterScoreRecord.Subjects)
                                    {
                                        // 彈性課程記數
                                        int AlternativeCourse = 0;
                                        int SubjectCourseNum = 0;

                                        // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                        if (subjectscore.Value.Domain == "彈性課程" || subjectscore.Value.Domain == "彈性學習" || subjectscore.Value.Domain == "")
                                        {
                                            if (AlternativeCourseDict.ContainsKey(subjectscore.Value.Subject))
                                            {
                                                AlternativeCourse = AlternativeCourseDict[subjectscore.Value.Subject];
                                                //紀錄成績
                                                if (subjectScore_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2)))
                                                {
                                                    subjectScore_dict["彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2)] = subjectscore.Value.Score;
                                                }

                                                //紀錄等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2)))
                                                {
                                                    subjectLevel_dict["彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2)] = ScoreToLevel(subjectscore.Value.Score);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (SubjectCourseDict.ContainsKey(subjectscore.Value.Domain))
                                            {
                                                if (SubjectCourseDict[subjectscore.Value.Domain].ContainsKey(subjectscore.Value.Subject))
                                                {
                                                    SubjectCourseNum = SubjectCourseDict[subjectscore.Value.Domain][subjectscore.Value.Subject];
                                                    //紀錄成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2)] = subjectscore.Value.Score;
                                                    }

                                                    //換算等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2)] = ScoreToLevel(subjectscore.Value.Score);
                                                    }
                                                }
                                            }
                                        }

                                    }

                                    //學期學習領域(七大)成績
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_學習領域總成績_成績_" + (grade * 2)))
                                    {
                                        domainScore_dict["領域_學習領域總成績_成績_" + (grade * 2)] = semesterScoreRecord.LearnDomainScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_學習領域總成績_等第_" + (grade * 2)))
                                    {
                                        domainLevel_dict["領域_學習領域總成績_等第_" + (grade * 2)] = ScoreToLevel(semesterScoreRecord.LearnDomainScore);
                                    }

                                    //課程學習成績(包括彈性課程成績)
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_課程學習成績_成績_" + (grade * 2)))
                                    {
                                        domainScore_dict["領域_課程學習成績_成績_" + (grade * 2)] = semesterScoreRecord.CourseLearnScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_課程學習成績_等第_" + (grade * 2)))
                                    {
                                        domainLevel_dict["領域_課程學習成績_等第_" + (grade * 2)] = ScoreToLevel(semesterScoreRecord.CourseLearnScore);
                                    }
                                }
                            }
                        }

                    }

                    // 填領域分數
                    foreach (string key in domainScore_dict.Keys)
                    {
                        row[key] = domainScore_dict[key];
                    }

                    // 填領域等第
                    foreach (string key in domainLevel_dict.Keys)
                    {
                        row[key] = domainLevel_dict[key];
                    }

                    // 填科目分數
                    foreach (string key in subjectScore_dict.Keys)
                    {
                        row[key] = subjectScore_dict[key];
                    }

                    // 填科目等第
                    foreach (string key in subjectLevel_dict.Keys)
                    {
                        row[key] = subjectLevel_dict[key];
                    }
                }

                //畢業分數
                if (_GraduateScoreRecordDic.ContainsKey(stuID))
                {
                    row["畢業總成績_平均"] = _GraduateScoreRecordDic[stuID].LearnDomainScore;
                    row["畢業總成績_等第"] = ScoreToLevel(_GraduateScoreRecordDic[stuID].LearnDomainScore);

                    // 60 分 就可以 准予畢業
                    row["准予畢業"] = _GraduateScoreRecordDic[stuID].LearnDomainScore >= 60 ? "■" : "□";
                    row["發給修業證書"] = _GraduateScoreRecordDic[stuID].LearnDomainScore >= 60 ? "□" : "■";
                }

                // 異動資料
                if (_UpdateRecordRecordDic.ContainsKey(stuID))
                {
                    int updateRecordCount = 1;

                    _UpdateRecordRecordDic[stuID].Sort((x, y) => { return x.UpdateDate.CompareTo(y.UpdateDate); });

                    foreach (K12.Data.UpdateRecordRecord urr in _UpdateRecordRecordDic[stuID])
                    {
                        // 新生異動為1 ，且理論上 一個人 會有1筆新生異動
                        if (urr.UpdateCode == "1")
                        {
                            DateTime enterday = new DateTime();

                            enterday = DateTime.Parse(urr.UpdateDate);
                            // 轉換入學時間 成 2005/09/06 的格式
                            row["入學年月"] = enterday.ToString("yyyy/MM");

                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "" + urr.Attributes["GraduateSchool"];
                            row["異動紀錄" + updateRecordCount + "_類別"] = "新生";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 畢業（2）
                        if (urr.UpdateCode == "2")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "畢業";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        //  當異動為 轉入 (3) 、轉出(4)時 要依序顯示在 報表上
                        if (urr.UpdateCode == "3" || urr.UpdateCode == "4")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = urr.Attributes["ImportExportSchool"]; // 取得異動校名的方法
                            row["異動紀錄" + updateRecordCount + "_類別"] = urr.UpdateCode == "3" ? "轉入" : "轉出";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;
                            updateRecordCount++;
                        }

                        //因為無法新增休學（5）及續讀（8）這兩種異動，所以這邊先不處理這兩種異動

                        // 復學（6）
                        if (urr.UpdateCode == "6")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "復學";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 中輟（7）
                        if (urr.UpdateCode == "7")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "中輟";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 更正學籍（9）
                        if (urr.UpdateCode == "9")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "更正學籍";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 延長修業年限（10）
                        if (urr.UpdateCode == "10")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "延長修業年限";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 死亡（11）
                        if (urr.UpdateCode == "11")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "死亡";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }
                    }
                }

                // 日常生活表現、校內外特殊表現
                if (_MoralScoreRecordDic.ContainsKey(stuID))
                {
                    for (int grade = 1; grade <= 3; grade++)
                    {
                        foreach (var msr in _MoralScoreRecordDic[stuID])
                        {
                            if (msr.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                if (msr.Semester == 1)
                                {
                                    if (textScore_dict.ContainsKey("日常生活表現及具體建議_" + (grade * 2 - 1)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("DailyLifeRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["日常生活表現及具體建議_" + (grade * 2 - 1)] = msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"].Value;
                                            }
                                        }
                                    }
                                    if (textScore_dict.ContainsKey("校內外特殊表現_" + (grade * 2 - 1)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("OtherRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["校內外特殊表現_" + (grade * 2 - 1)] = msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"].Value;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (textScore_dict.ContainsKey("日常生活表現及具體建議_" + (grade * 2)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("DailyLifeRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["日常生活表現及具體建議_" + (grade * 2)] = msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"].Value;
                                            }
                                        }
                                    }
                                    if (textScore_dict.ContainsKey("校內外特殊表現_" + (grade * 2)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("OtherRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["校內外特殊表現_" + (grade * 2)] = msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"].Value;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //填值
                    foreach (string key in textScore_dict.Keys)
                    {
                        row[key] = textScore_dict[key];
                    }
                }

                table.Rows.Add(row);

                //回報進度
                int percent = ((student_counter * 100 / _StudentIDs.Count));

                _MasterWorker.ReportProgress(percent, "學生學籍表產生中...進行到第" + (student_counter) + "/" + _StudentIDs.Count + "學生");

                student_counter++;
            }

            //選擇 目前的樣板
            document = new Document(_Preference.Template.GetStream());
            //document = Preference.Template.ToDocument();

            //執行 合併列印
            document.MailMerge.Execute(table);

            // 最終產物 .doc
            e.Result = document;

            _MasterWorker.ReportProgress(100);
        }

        private void MasterWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Util.EnableControls(this);

            if (e.Error == null)
            {
                Document doc = e.Result as Document;

                //單檔列印
                if (OneFileSave.Checked)
                {
                    FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                    folderBrowserDialog.Description = "請選擇儲存資料夾";
                    folderBrowserDialog.ShowNewFolderButton = true;

                    if (folderBrowserDialog.ShowDialog() == DialogResult.Cancel)
                    {
                        return;
                    }

                    _FileBrowserDialogPath = folderBrowserDialog.SelectedPath;
                    Util.DisableControls(this);
                    _ConvertToPDF_Worker.RunWorkerAsync();
                }
                else
                {
                    if (_Preference.ConvertToPDF)
                    {
                        MotherForm.SetStatusBarMessage("正在轉換PDF格式... 請耐心等候");
                    }
                    Util.DisableControls(this);
                    _ConvertToPDF_Worker.RunWorkerAsync();
                }
            }
            else
            {
                MessageBox.Show(e.Error.Message);
            }

            if (_Preference.ConvertToPDF)
            {
                MotherForm.SetStatusBarMessage("正在轉換PDF格式", 0);
            }
            else
            {
                MotherForm.SetStatusBarMessage("產生完成", 100);
            }
        }

        private void ConvertToPDF_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Document doc = _E_For_ConvertToPDF_Worker.Result as Document;

            if (!OneFileSave.Checked)
            {
                Util.Save(doc, "學籍表", _Preference.ConvertToPDF);
            }
            else
            {
                int i = 0;

                foreach (Section section in doc.Sections)
                {
                    // 依照 學號_身分字號_班級_座號_姓名 .doc 來存檔
                    string fileName = "";

                    Document document = new Document();
                    document.Sections.Clear();
                    document.Sections.Add(document.ImportNode(section, true));

                    fileName = _PrintStudents[i].StudentNumber;

                    fileName = _PrintStudents[i].StudentNumber;

                    fileName += "_" + _PrintStudents[i].IDNumber;

                    if (!string.IsNullOrEmpty(_PrintStudents[i].RefClassID))
                    {
                        fileName += "_" + _PrintStudents[i].Class.Name;
                    }
                    else
                    {
                        fileName += "_";
                    }

                    fileName += "_" + _PrintStudents[i].SeatNo;

                    fileName += "_" + _PrintStudents[i].Name;

                    //document.Save(fbd.SelectedPath + "\\" +fileName+ ".doc");

                    if (_Preference.ConvertToPDF)
                    {
                        //string fPath = fbd.SelectedPath + "\\" + fileName + ".pdf";

                        string fPath = _FileBrowserDialogPath + "\\" + fileName + ".pdf";

                        FileInfo fi = new FileInfo(fPath);

                        DirectoryInfo folder = new DirectoryInfo(Path.Combine(fi.DirectoryName, Path.GetRandomFileName()));
                        if (!folder.Exists) folder.Create();

                        FileInfo fileinfo = new FileInfo(Path.Combine(folder.FullName, fi.Name));

                        string XmlFileName = fileinfo.FullName.Substring(0, fileinfo.FullName.Length - fileinfo.Extension.Length) + ".xml";
                        string PDFFileName = fileinfo.FullName.Substring(0, fileinfo.FullName.Length - fileinfo.Extension.Length) + ".pdf";

                        document.Save(XmlFileName, Aspose.Words.SaveFormat.Pdf);

                        Aspose.Pdf.Generator.Pdf pdf1 = new Aspose.Pdf.Generator.Pdf();

                        pdf1.BindXML(XmlFileName, null);
                        pdf1.Save(PDFFileName);

                        if (File.Exists(fPath))
                            File.Delete(Path.Combine(fi.DirectoryName, fi.Name));

                        File.Move(PDFFileName, fPath);
                        folder.Delete(true);

                        int percent = (((i + 1) * 100 / doc.Sections.Count));

                        _ConvertToPDF_Worker.ReportProgress(percent, "PDF轉換中...進行到" + (i + 1) + "/" + doc.Sections.Count + "個檔案");
                    }
                    else
                    {
                        document.Save(_FileBrowserDialogPath + "\\" + fileName + ".doc");

                        int percent = (((i + 1) * 100 / doc.Sections.Count));

                        _ConvertToPDF_Worker.ReportProgress(percent, "Doc存檔...進行到" + (i + 1) + "/" + doc.Sections.Count + "個檔案");
                    }

                    i++;
                }
            }
        }

        private void ConvertToPDF_Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Util.EnableControls(this);

            if (_Preference.ConvertToPDF)
            {
                MotherForm.SetStatusBarMessage("PDF轉換完成", 100);
            }
            else
            {
                MotherForm.SetStatusBarMessage("存檔完成", 100);
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (_StudentIDs.Count <= 0)
            {
                MessageBox.Show("沒有任何學生可以列印");
                return;
            }

            _Preference.ConvertToPDF = rtnPDF.Checked;
            _Preference.OneFileSave = OneFileSave.Checked;
            _Preference.Save(); //儲存設定值

            //關閉畫面控制項
            Util.DisableControls(this);

            _MasterWorker.RunWorkerAsync();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void lnkTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportTemplate defaultTemplate = new ReportTemplate(Resources.新竹國中學籍表_樣板_, TemplateType.Word);
            TemplateSettingForm templateForm = new TemplateSettingForm(_Preference.Template, defaultTemplate);
            templateForm.DefaultFileName = "學籍表(樣板).doc";

            if (templateForm.ShowDialog() == DialogResult.OK)
            {
                _Preference.Template = (templateForm.Template == defaultTemplate) ? null : templateForm.Template;
                _Preference.Save();
            }
        }

        //供使用者下載學籍表合併欄位總表
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Aspose.Words.Document document = new Aspose.Words.Document();
            document = new Aspose.Words.Document(new MemoryStream(Resources.新竹國中學籍表功能變數));

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "另存新檔";
            saveFileDialog.FileName = "學籍表合併欄位總表" + ".doc";
            saveFileDialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    document.Save(saveFileDialog.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(saveFileDialog.FileName);
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        ///<summary>
        ///換算分數與等第
        ///</summary>
        private string ScoreToLevel(decimal? score)
        {
            string level = "";
            if (score >= 90)
            {
                level = "優";
            }
            else if (score >= 80 && score < 90)
            {
                level = "甲";
            }
            else if (score >= 70 && score < 80)
            {
                level = "乙";
            }
            else if (score >= 60 && score < 70)
            {
                level = "丙";
            }
            else if (score < 60)
            {
                level = "丁";
            }
            else
            {
                level = "";
            }

            return level;
        }
    }
}