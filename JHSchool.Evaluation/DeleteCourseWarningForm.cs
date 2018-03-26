using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Aspose.Cells;

namespace JHSchool.Evaluation
{
    public partial class DeleteCourseWarningForm : BaseForm
    {

        Dictionary<string, List<JHSchool.Data.JHSCETakeRecord>> _studentScoreDict;

        public DeleteCourseWarningForm(string s, Dictionary<string, List<JHSchool.Data.JHSCETakeRecord>> studentScoreDict)
        {
            InitializeComponent();
            _studentScoreDict = studentScoreDict;
            labelX3.Text = s;
        }

        //匯出
        private void buttonX1_Click(object sender, EventArgs e)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells[0, 0].PutValue("班級");
            ws.Cells[0, 1].PutValue("座號");
            ws.Cells[0, 2].PutValue("學號");
            ws.Cells[0, 3].PutValue("姓名");
            ws.Cells[0, 4].PutValue("課程");
            ws.Cells[0, 5].PutValue("試別");
            ws.Cells[0, 6].PutValue("分數");

            int row_index = 1;

            foreach (KeyValuePair<string, List<JHSchool.Data.JHSCETakeRecord>> p in _studentScoreDict)
            {

                foreach (JHSchool.Data.JHSCETakeRecord r in p.Value)
                {
                    ws.Cells[row_index, 0].PutValue(p.Value[0].Student.Class != null ? p.Value[0].Student.Class.Name : "");
                    ws.Cells[row_index, 1].PutValue(p.Value[0].Student.SeatNo);
                    ws.Cells[row_index, 2].PutValue(p.Value[0].Student.StudentNumber);
                    ws.Cells[row_index, 3].PutValue(p.Value[0].Student.Name);
                    ws.Cells[row_index, 4].PutValue(r.Course.Name);
                    ws.Cells[row_index, 5].PutValue(r.Exam.Name);
                    ws.Cells[row_index, 6].PutValue(r.Score);

                    row_index++;
                }                
            }

            System.Windows.Forms.SaveFileDialog sd1 = new System.Windows.Forms.SaveFileDialog();
            sd1.Title = "另存新檔";
            sd1.FileName = "學生課程試別成績清單.xls";
            sd1.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(sd1.FileName, FileFormatType.Excel2003);
                    System.Diagnostics.Process.Start(sd1.FileName);
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }
            }
        }

        // 確認
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
            this.Close();
        }

        //離開
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
            
        }
    }
}

