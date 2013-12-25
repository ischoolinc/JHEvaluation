using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace JHEvaluation.ClassSemesterScoreAvgComparison
{
    public partial class ClassSemsScoreAvgCmpForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _BGWork;        
        private List<string> _ClassIDList;
        private int _SchoolYear;
        private int _Semester;
        private List<string> _SelectSubjName;
        private List<string> _SelectDomainName;
        private bool _isBGWorkBusy = false;

        public ClassSemsScoreAvgCmpForm()
        {
            InitializeComponent();

            _SelectSubjName = new List<string>();
            _SelectDomainName = new List<string>();
            _BGWork = new BackgroundWorker();
            _BGWork.DoWork += new DoWorkEventHandler(_BGWork_DoWork);
            _BGWork.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWork_RunWorkerCompleted);
        }

        void _BGWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBGWorkBusy)
            {
                _isBGWorkBusy = false;
                _BGWork.RunWorkerAsync();
                return;
            }
            LoadSubjectDomainNameToForm();
        }

        void _BGWork_DoWork(object sender, DoWorkEventArgs e)
        {
            DAL.DALTransfer.LoadSemesterScoreRecord(_SchoolYear, _Semester, _ClassIDList);
        }

        // 畫面預設
        private void LoadDefaultSchoolYearAndSemester()
        {
            cbxSchoolYear.Items.Clear();
            cbxSemester.Items.Clear();
            cbxSchoolYear.Text = JHSchool.School.DefaultSchoolYear;
            cbxSemester.Text = JHSchool.School.DefaultSemester;

            int intDefaultSchoolYear;
            int.TryParse(JHSchool.School.DefaultSchoolYear, out intDefaultSchoolYear);

            for (int i = intDefaultSchoolYear - 2; i <= intDefaultSchoolYear + 2; i++)
                cbxSchoolYear.Items.Add(i + "");

            cbxSemester.Items.Add("1"); 
            cbxSemester.Items.Add("2");
        }

        private void LoadSubjectDomainNameToForm()
        {
            lstSubj.Items.Clear();
            lstDomain.Items.Clear();           

            List<string> SubjNameList = DAL.DALTransfer.GetSubjectName();
            List<string> DoaminNameList = DAL.DALTransfer.GetDoaminName();

            ScoreHeaderIndexer shtSubj = new ScoreHeaderIndexer();
            ScoreHeaderIndexer shtDomain = new ScoreHeaderIndexer();
            foreach (string str in SubjNameList)
                shtSubj.Add(str, false, 0);
            shtSubj.Sort(DAL.DALTransfer.GetDomainMapping());

            foreach (Header h in shtSubj)
                lstSubj.Items.Add(h.Name);

            foreach (string str in DoaminNameList)
                shtDomain.Add(str, true, 0);
            shtDomain.Sort(DAL.DALTransfer.GetDomainMapping());

            foreach (Header h in shtDomain)
                lstDomain.Items.Add(h.Name);

            pictureBox1.Visible = false;
            pictureBox2.Visible = false;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ClassSemsScoreAvgCmpForm_Load(object sender, EventArgs e)
        {
            _ClassIDList = K12.Presentation.NLDPanels.Class.SelectedSource;
            int.TryParse(JHSchool.School.DefaultSchoolYear, out _SchoolYear);
            int.TryParse(JHSchool.School.DefaultSemester, out _Semester);
            LoadDefaultSchoolYearAndSemester();
            _BGWork.RunWorkerAsync();

            
            

        }

        private void cbxSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReloadFormData();               
        }

        private void cbxSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReloadFormData();
        }

        private void ReloadFormData()
        {
            pictureBox1.Visible = true;
            pictureBox2.Visible = true;
            int.TryParse(cbxSchoolYear.Text, out _SchoolYear);
            int.TryParse(cbxSemester.Text, out _Semester);
            if (_BGWork.IsBusy)
                _isBGWorkBusy = true;
            else
                _BGWork.RunWorkerAsync();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            _SelectSubjName.Clear ();
            _SelectDomainName.Clear ();
            ClassSemsScoreAvgCmpReporter.UserSelectCount = 0;
            foreach (ListViewItem liv in lstSubj.CheckedItems)
                _SelectSubjName.Add(liv.Text);

            foreach (ListViewItem liv in lstDomain.CheckedItems)
                _SelectDomainName.Add(liv.Text);

            // 檢查是否超過最大列印數
            ClassSemsScoreAvgCmpReporter.UserSelectCount = _SelectSubjName.Count + _SelectDomainName.Count;
            if ((_SelectSubjName.Count + _SelectDomainName.Count) > 20)
            {
                MsgBox.Show("最多只能勾選20個產生,請重新勾選.");
                return;
            }

            btnPrint.Enabled = false;
            // 取得成績
            Dictionary<string, DAL.ClassEntity> ClassEntityDic = DAL.DALTransfer.GetClassEntityDic(_SelectSubjName, _SelectDomainName);

            // 產生報表
            ClassSemsScoreAvgCmpReporter cssacr = new ClassSemsScoreAvgCmpReporter(ClassEntityDic, _SelectSubjName, _SelectDomainName,JHSchool.School.ChineseName,cbxSchoolYear.Text,cbxSemester.Text);

            btnPrint.Enabled = true;
        }

    }
}
