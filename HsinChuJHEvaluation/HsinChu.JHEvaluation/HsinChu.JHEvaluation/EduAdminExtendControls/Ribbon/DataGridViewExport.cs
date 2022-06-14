using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;
using System.Windows.Forms;

namespace HsinChu.JHEvaluation.EduAdminExtendControls.Ribbon
{
    internal class DataGridViewExport
    {
        Workbook _workbook;
        Worksheet _worksheet;
        List<int> _colIndexes;

        public DataGridViewExport(DataGridView dgv)
        {
            _workbook = new Workbook();
            _workbook.Worksheets.Clear();
            _worksheet = _workbook.Worksheets[_workbook.Worksheets.Add()];
            _worksheet.Name = "Sheet1";
            _colIndexes = new List<int>();

            int sheetRowIndex = 0;
            int sheetColIndex = 0;
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Visible == false) continue;
                _colIndexes.Add(col.Index);
                _worksheet.Cells[sheetRowIndex, sheetColIndex++].PutValue(col.HeaderText);
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                sheetRowIndex++;
                sheetColIndex = 0;
                foreach (int colIndex in _colIndexes)
                    _worksheet.Cells[sheetRowIndex, sheetColIndex++].PutValue("" + row.Cells[colIndex].Value);
            }

             _worksheet.AutoFitColumns();
        }

        public void Save(string fileName)
        {
            try
            {
                _workbook.Save(fileName, FileFormatType.Excel2003);

                DialogResult result = new DialogResult();
                result = FISCA.Presentation.Controls.MsgBox.Show("檔案儲存完成，是否開啟檔案？", "是否開啟", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(fileName);
                    }
                    catch (Exception ex)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("開啟檔案發生錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗：" + ex.Message);
            }
        }
    }
}
