using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;
using System.Xml;
using System.IO;
using Campus.Report2014;

namespace KaoHsiung.StudentRecordReport.Forms
{
    public partial class PrintConfigForm : BaseForm
    {
        private ReportConfiguration _config;

        private ReportConfiguration _Dylanconfig;

        public PrintConfigForm()
        {
            InitializeComponent();
            this.MinimumSize = this.Size;
            this.MaximumSize = this.Size;
            _config = new ReportConfiguration(Global.ReportName);

            _Dylanconfig = new ReportConfiguration(Global.OneFileSave);

            //SetupDefaultTemplate();
            LoadConfig();
        }

        //private void SetupDefaultTemplate()
        //{
        //    //如果設定中沒有範本，使用預設範本。
        //    if (_config.Template == null)
        //    {
        //        ReportTemplate template = new ReportTemplate(Properties.Resources.高雄國中學籍表, TemplateType.Word);
        //        _config.Template = template;
        //    }
        //}

        private void LoadConfig()
        {
            string print = _config.GetString("領域科目設定", "Domain");
            if (print == "Domain")
                rbDomain.Checked = true;
            else if (print == "Subject")
                rbSubject.Checked = true;

            chkPeriod.Checked = _config.GetBoolean("列印節數", true);
            chkCredit.Checked = _config.GetBoolean("列印權數", false);

            chkText.Checked = _config.GetBoolean("列印文字評語", true);
            rtnPDF.Checked = _config.GetBoolean("輸出成PDF格式", false);

            checkBoxX1.Checked = _Dylanconfig.GetBoolean("單檔儲存", false);
            checkBoxIDNumberFormat.Checked = _Dylanconfig.GetBoolean("使用身分證號格式", false);

            // 處理機敏資料，使用原本 config
            cbxName.Checked = _config.GetBoolean("遮罩姓名", false);
            cbxIDNumber.Checked = _config.GetBoolean("遮罩身分證號", false);
            cbxBirthday.Checked = _config.GetBoolean("遮罩生日", false);
            cbxPhone.Checked = _config.GetBoolean("遮罩電話", false);
            cbxAddress.Checked = _config.GetBoolean("遮罩地址", false);


        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _config.SetString("領域科目設定", (rbDomain.Checked) ? "Domain" : "Subject");
            _config.SetBoolean("列印節數", chkPeriod.Checked);
            _config.SetBoolean("列印權數", chkCredit.Checked);
            _config.SetBoolean("列印文字評語", chkText.Checked);
            _config.SetBoolean("輸出成PDF格式", rtnPDF.Checked);
            _config.SetBoolean("遮罩姓名", cbxName.Checked);
            _config.SetBoolean("遮罩身分證號", cbxIDNumber.Checked);
            _config.SetBoolean("遮罩生日", cbxBirthday.Checked);
            _config.SetBoolean("遮罩電話", cbxPhone.Checked);
            _config.SetBoolean("遮罩地址", cbxAddress.Checked);
            _config.Save();

            // 儲存單檔列印設定
            _Dylanconfig.SetBoolean("單檔儲存", checkBoxX1.Checked || checkBoxIDNumberFormat.Checked);
            _Dylanconfig.SetBoolean("使用身分證號格式", checkBoxIDNumberFormat.Checked);
            _Dylanconfig.Save();

            this.DialogResult = DialogResult.OK;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            // 如果選擇了學號_班級_座號格式，取消身分證號_姓名格式
            if (checkBoxX1.Checked)
            {
                checkBoxIDNumberFormat.Checked = false;
            }
        }

        private void checkBoxIDNumberFormat_CheckedChanged(object sender, EventArgs e)
        {
            // 如果選擇了身分證號_姓名格式，取消學號_班級_座號格式
            if (checkBoxIDNumberFormat.Checked)
            {
                checkBoxX1.Checked = false;
            }
        }
    }
}
