using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using KaoHsingReExamScoreReport.DAO;

namespace KaoHsingReExamScoreReport.Forms
{
    public partial class ReDomainForTeacherForm : BaseForm
    {
        BackgroundWorker _bgWorker;
        int _SchoolYear = 0;
        int _Semester = 0;
        List<ClassData> _ClassDataList;
        List<StudentData> _StudentDataList;
        List<string> _SelectClassIDList;

        public ReDomainForTeacherForm(List<string> ClassIDList)
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _ClassDataList = new List<ClassData>();
            _StudentDataList = new List<StudentData>();
            _SelectClassIDList = ClassIDList;
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("補考名單產生中 ...", e.ProgressPercentage);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnPrint.Enabled = true;
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ReDomainForTeacherForm_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;

            // 預設學年度、學期
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            
            // 使用者選學年度、學期
            _SchoolYear = iptSchoolYear.Value;
            _Semester = iptSemester.Value;

            _bgWorker.RunWorkerAsync();
        }
    }
}
