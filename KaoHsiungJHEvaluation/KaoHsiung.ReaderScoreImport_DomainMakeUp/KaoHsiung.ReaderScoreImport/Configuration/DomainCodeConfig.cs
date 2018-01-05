using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using FISCA.UDT;
using KaoHsiung.ReaderScoreImport_DomainMakeUp.UDT;
using KaoHsiung.ReaderScoreImport_DomainMakeUp.Mapper;

namespace KaoHsiung.ReaderScoreImport_DomainMakeUp
{
    public partial class DomainCodeConfig : BaseForm
    {
        //private List<JHClassRecord> _classes;
        protected BackgroundWorker _worker;
        protected AccessHelper _accessHelper;
        private List<DomainCode_DomainMakeUp> _list;

        public DomainCodeConfig()
        {
            InitializeComponent();
            InitAccessHelper();
            InitWorker();
            //Init();
            Startup();
        }

        private void InitAccessHelper()
        {
            _accessHelper = new AccessHelper();
        }

        private void InitWorker()
        {
            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(_worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);
            picWaiting.Visible = true;
            //_worker.RunWorkerAsync();
        }



        private void _worker_DoWork(object sender, DoWorkEventArgs e)
        {
            DoWork(sender, e);
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            picWaiting.Visible = false;
            RunWorkerCompleted(sender, e);
        }


        public void Startup()
        {
            _worker.RunWorkerAsync();
        }

        //private void Init()
        //{
        //    _classes = new List<JHClassRecord>();
        //}

        protected void DoWork(object sender, DoWorkEventArgs e)
        {
            _list = _accessHelper.Select<DomainCode_DomainMakeUp>();
            //if (_list.Count <= 0)
            //    _classes = JHClass.SelectAll();
        }

        protected void RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            dataGridViewX1.SuspendLayout();

            //foreach (var cla in _classes)
            //{
            //    DataGridViewRow row = new DataGridViewRow();
            //    row.CreateCells(dataGridViewX1, cla.Name, "");

            //    dataGridViewX1.Rows.Add(row);
            //    ValidateCodeFormat(row.Cells[chCode.Index],3);
            //}

            foreach (var record in _list)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1, record.Domain, record.Code);

                dataGridViewX1.Rows.Add(row);
                ValidateCodeFormat(row.Cells[chCode.Index], 5);
            }

            dataGridViewX1.ResumeLayout();
        }

        protected void Save()
        {
            dataGridViewX1.EndEdit();

            if (!IsValid(dataGridViewX1))
            {
                MsgBox.Show("請先修正錯誤");
                return;
            }

            List<DomainCode_DomainMakeUp> newList = new List<DomainCode_DomainMakeUp>();
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;

                DomainCode_DomainMakeUp cc = new DomainCode_DomainMakeUp();
                cc.Domain = "" + row.Cells[chDomainName.Index].Value;
                cc.Code = "" + row.Cells[chCode.Index].Value;
                newList.Add(cc);
            }
            _accessHelper.DeletedValues(_list.ToArray());
            _accessHelper.InsertValues(newList.ToArray());
            ClassCodeMapper.Instance.Reload();

            this.DialogResult = DialogResult.OK;
        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];

            if (cell.OwningColumn == chCode)
            {
                ValidateCodeFormat(cell, 5);
                ValidateDuplication(dataGridViewX1, chCode.Index, "代碼不能重覆");
            }
            else
            {
                ValidateDuplication(dataGridViewX1, chDomainName.Index, "領域名稱不能重覆");
            }
        }

        protected void Export()
        {
            ImportExport.Export(dataGridViewX1, "匯出領域代碼設定");
        }

        protected void Import()
        {
            ImportExport.Import(dataGridViewX1, new List<string>() { "領域", "代碼" });

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;
                DataGridViewCell cell = row.Cells[chCode.Index];
                ValidateCodeFormat(cell, 5);
            }
            ValidateDuplication(dataGridViewX1, chCode.Index, "代碼不能重覆");
            ValidateDuplication(dataGridViewX1, chDomainName.Index, "領域名稱不能重覆");
        }



        protected bool IsValid(DataGridView dataGridViewX1)
        {
            bool valid = true;

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                    {
                        valid &= false;
                        break;
                    }
                }
            }

            return valid;
        }


        protected void ValidateDuplication(DataGridView dataGridViewX1, int index, string message)
        {
            List<string> codeList = new List<string>();
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;

                DataGridViewCell runningCell = row.Cells[index];
                if (string.IsNullOrEmpty("" + runningCell.Value)) continue;

                if (!codeList.Contains("" + runningCell.Value))
                {
                    codeList.Add("" + runningCell.Value);
                    runningCell.ErrorText = string.Empty;
                }
                else if (string.IsNullOrEmpty(runningCell.ErrorText))
                    runningCell.ErrorText = message;
            }
        }

        protected void ValidateCodeFormat(DataGridViewCell cell, int num)
        {
            int i;
            string code = "" + cell.Value;
            if (code.Length != num)
                cell.ErrorText = "代碼必須為 " + num + " 位數";
            else if (!int.TryParse(code, out i))
                cell.ErrorText = "代碼必須為數字";
            else
                cell.ErrorText = "";
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            Export();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            Import();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
