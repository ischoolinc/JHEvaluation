using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using K12.Data;
using Aspose.Cells;
using System.IO;
using System.Windows.Forms;

namespace ImportMakeUpScore
{
    public partial class MainForm : BaseForm
    {

        List<string> ClassIDList;
        int selSchoolYear = 0;
        int selSemester = 0;

        BackgroundWorker bgWorker;
        bool userSelectDoamin = false;

        public MainForm(List<string> classIDs)
        {
            InitializeComponent();
            ClassIDList = classIDs;
            bgWorker = new BackgroundWorker();
            bgWorker.DoWork += BgWorker_DoWork;
            bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
            bgWorker.WorkerReportsProgress = true;
            bgWorker.ProgressChanged += BgWorker_ProgressChanged;
        }

        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級補考成績匯入表產生中...", e.ProgressPercentage);
        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            btnPrint.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            if (e.Error != null)
            {
                MsgBox.Show(e.Error.Message);
                return;
            }

            try
            {
                Workbook wb = e.Result as Workbook;
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "另存新檔";

                string st = "領域";
                if (userSelectDoamin == false)
                    st = "科目";

                save.FileName = Application.StartupPath + "\\Reports\\" + selSchoolYear + "學年度第" + selSemester + "學期" + st + "補考成績匯入表.xls";
                save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

                wb.Save(save.FileName, FileFormatType.Excel97To2003);
                System.Diagnostics.Process.Start(save.FileName);
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存錯誤," + ex.Message);
                return;
            }

        }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {

            List<ClassRecord> classRecords = K12.Data.Class.SelectByIDs(ClassIDList);


            List<StudentRecord> studentRecords = K12.Data.Student.SelectByClassIDs(ClassIDList);
            List<JHSemesterScoreRecord> semesterScoreList = JHSemesterScore.SelectBySchoolYearAndSemester(studentRecords.Select(x => x.ID).ToList(), selSchoolYear, selSemester);

            Dictionary<string, ClassRecord> classIdToRecord = new Dictionary<string, ClassRecord>();

            //班級名稱對照
            foreach (ClassRecord cr in classRecords)
            {
                if (!classIdToRecord.ContainsKey(cr.ID))
                {
                    classIdToRecord.Add(cr.ID, cr);
                }
            }

            //學生物件整理
            List<StudentObj> stuObjs = new List<StudentObj>();
            foreach (StudentRecord s in studentRecords)
            {
                StudentObj obj = new StudentObj(s);
                obj.ClassRecord = classIdToRecord.ContainsKey(s.RefClassID) ? classIdToRecord[s.RefClassID] : new ClassRecord();
                stuObjs.Add(obj);
            }

            stuObjs.Sort(delegate (StudentObj x, StudentObj y)
            {
                string x1 = x.ClassRecord.DisplayOrder.PadLeft(3, '0');
                string xx = (x.ClassRecord.GradeYear + "").PadLeft(3, '0');
                xx += x1 == "000" ? "999" : x1;
                xx += x.ClassRecord.Name.PadLeft(20, '0');
                xx += (x.StudentRecord.SeatNo + "").PadLeft(3, '0');

                string y1 = y.ClassRecord.DisplayOrder.PadLeft(3, '0');
                string yy = (y.ClassRecord.GradeYear + "").PadLeft(3, '0');
                yy += y1 == "000" ? "999" : y1;
                yy += y.ClassRecord.Name.PadLeft(20, '0');
                yy += (y.StudentRecord.SeatNo + "").PadLeft(3, '0');

                return xx.CompareTo(yy);
            });

            // 領域補對照
            Dictionary<string, Dictionary<string, DomainScore>> MakeUpDomainDic = new Dictionary<string, Dictionary<string, DomainScore>>();

            // 科目補考對照
            Dictionary<string, Dictionary<string, SubjectScore>> MakeUpSubjDic = new Dictionary<string, Dictionary<string, SubjectScore>>();

            // 處理需要補考學期成績
            foreach (JHSemesterScoreRecord JHssr in semesterScoreList)
            {
                //領域
                foreach (KeyValuePair<string, DomainScore> item in JHssr.Domains)
                {

                    if (item.Value.Score.HasValue && item.Value.Score.Value < 60)
                    {
                        if (!MakeUpDomainDic.ContainsKey(JHssr.RefStudentID))
                            MakeUpDomainDic.Add(JHssr.RefStudentID, new Dictionary<string, DomainScore>());

                        MakeUpDomainDic[JHssr.RefStudentID].Add(item.Key, item.Value);

                    }
                }

                //科目
                foreach (string subj in JHssr.Subjects.Keys)
                {
                    SubjectScore ss = JHssr.Subjects[subj];
                    if (JHssr.Subjects[subj].Score.HasValue && JHssr.Subjects[subj].Score.Value < 60)
                    {
                        if (!MakeUpSubjDic.ContainsKey(JHssr.RefStudentID))
                            MakeUpSubjDic.Add(JHssr.RefStudentID, new Dictionary<string, SubjectScore>());

                        //科目名稱不應該有重覆
                        MakeUpSubjDic[JHssr.RefStudentID].Add(subj, ss);
                    }
                }
            }

            // templae
            Workbook wb = null;
            Worksheet wst = null;
            int rowIdx = 1;
            // 使用領域
            if (userSelectDoamin)
            {
                wb = new Workbook(new MemoryStream(Properties.Resources.匯出學期領域成績_領域));
                wst = wb.Worksheets[0];
                rowIdx = 1;

                if (MakeUpDomainDic.Count > 0)
                {
                    foreach (StudentObj so in stuObjs)
                    {
                        //只列印一般生
                        if (so.StudentRecord.Status != StudentRecord.StudentStatus.一般)
                            continue;

                        if (MakeUpDomainDic.ContainsKey(so.StudentRecord.ID))
                        {
                            foreach (string dName in MakeUpDomainDic[so.StudentRecord.ID].Keys)
                            {
                                // 學號 0
                                wst.Cells[rowIdx, 0].PutValue(so.StudentRecord.StudentNumber);

                                // 班級 1
                                wst.Cells[rowIdx, 1].PutValue(so.ClassRecord.Name);

                                // 座號 2
                                if (so.StudentRecord.SeatNo.HasValue)
                                    wst.Cells[rowIdx, 2].PutValue(so.StudentRecord.SeatNo.Value);

                                // 姓名 3
                                wst.Cells[rowIdx, 3].PutValue(so.StudentRecord.Name);

                                // 領域 4
                                wst.Cells[rowIdx, 4].PutValue(dName);
                                // 學年度 5
                                wst.Cells[rowIdx, 5].PutValue(selSchoolYear);

                                // 學期 6
                                wst.Cells[rowIdx, 6].PutValue(selSemester);

                                // 補考成績 7
                                if (MakeUpDomainDic[so.StudentRecord.ID][dName].ScoreMakeup.HasValue)
                                    wst.Cells[rowIdx, 7].PutValue(MakeUpDomainDic[so.StudentRecord.ID][dName].ScoreMakeup.Value);

                                rowIdx++;
                            }

                        }

                    }
                }
                e.Result = wb;
            }
            else
            {
                wb = new Workbook(new MemoryStream(Properties.Resources.匯出學期科目成績_科目));
                wst = wb.Worksheets[0];
                rowIdx = 1;

                if (MakeUpSubjDic.Count > 0)
                {
                    foreach (StudentObj so in stuObjs)
                    {
                        //只列印一般生
                        if (so.StudentRecord.Status != StudentRecord.StudentStatus.一般)
                            continue;

                        if (MakeUpSubjDic.ContainsKey(so.StudentRecord.ID))
                        {
                            foreach (string sName in MakeUpSubjDic[so.StudentRecord.ID].Keys)
                            {
                                // 學號 0
                                wst.Cells[rowIdx, 0].PutValue(so.StudentRecord.StudentNumber);

                                // 班級 1
                                wst.Cells[rowIdx, 1].PutValue(so.ClassRecord.Name);

                                // 座號 2
                                if (so.StudentRecord.SeatNo.HasValue)
                                    wst.Cells[rowIdx, 2].PutValue(so.StudentRecord.SeatNo.Value);

                                // 姓名 3
                                wst.Cells[rowIdx, 3].PutValue(so.StudentRecord.Name);

                                // 領域 4
                                wst.Cells[rowIdx, 4].PutValue(MakeUpSubjDic[so.StudentRecord.ID][sName].Domain);

                                // 科目 5
                                wst.Cells[rowIdx, 5].PutValue(MakeUpSubjDic[so.StudentRecord.ID][sName].Subject);

                                // 學年度 6
                                wst.Cells[rowIdx, 6].PutValue(selSchoolYear);

                                // 學期 7
                                wst.Cells[rowIdx, 7].PutValue(selSemester);

                                // 補考成績 8
                                if (MakeUpSubjDic[so.StudentRecord.ID][sName].ScoreMakeup.HasValue)
                                    wst.Cells[rowIdx, 8].PutValue(MakeUpSubjDic[so.StudentRecord.ID][sName].ScoreMakeup.Value);

                                rowIdx++;
                            }

                        }

                    }
                }

                e.Result = wb;
            }

        }

        public class StudentObj
        {
            public StudentRecord StudentRecord;
            public ClassRecord ClassRecord;
            public StudentObj(StudentRecord s)
            {
                this.StudentRecord = s;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");
            ribDomain.Checked = true;

            try
            {
                int sc = int.Parse(K12.Data.School.DefaultSchoolYear);
                for (int i = sc - 3; i <= sc + 5; i++)
                    cboSchoolYear.Items.Add(i);

                // 目前學年度學期
                cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
                cboSemester.Text = K12.Data.School.DefaultSemester;
            }
            catch (Exception ex)
            {
                MsgBox.Show("錯誤訊息" + ex.Message);
            }

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            userSelectDoamin = false;

            if (ribDomain.Checked)
                userSelectDoamin = true;

            int sy, ss;
            if (!int.TryParse(cboSchoolYear.Text, out sy))
            {
                MsgBox.Show("學年度必須選擇為數字");
                return;
            }

            if (!int.TryParse(cboSemester.Text, out ss))
            {
                MsgBox.Show("學期必須選擇為數字");
                return;
            }

            selSchoolYear = sy;
            selSemester = ss;

            btnPrint.Enabled = false;

            bgWorker.RunWorkerAsync();


        }

    }
}
