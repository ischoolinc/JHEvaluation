﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using K12.Data.Configuration;
using K12.Data;
using FISCA.Presentation.Controls;

namespace JHSchool.Evaluation.EduAdminExtendControls.Ribbon
{
    public partial class DomainListTable : BaseForm
    {
        private const string ConfigName = "JHEvaluation_Subject_Ordinal";
        private const string ColumnKey = "DomainOrdinal";

        public DomainListTable()
        {
            InitializeComponent();

            // 取得108課綱領域白名單對照
            Global.LoadDomainMap108();

            LoadDomain();
            List<int> cols = new List<int>() { 1 };
            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dgv, cols);
        }

        private void LoadDomain()
        {
            bool containsIEPDomin = false;
            bool containsTechnologydomin = false;

            ConfigData cd = K12.Data.School.Configuration[ConfigName];
            if (cd.Contains(ColumnKey))
            {
                XmlElement element = cd.GetXml(ColumnKey, XmlHelper.LoadXml("<Domains/>"));
                foreach (XmlElement domainElement in element.SelectNodes("Domain"))
                {
                    string group = domainElement.GetAttribute("Group");
                    string name = domainElement.GetAttribute("Name");
                    string englishName = domainElement.GetAttribute("EnglishName");

                    if (name == "特殊需求")
                    {
                        containsIEPDomin = true;
                    }
                    if (name == "科技")
                    {
                        containsTechnologydomin = true;
                    }

                    DataGridViewRow row = new DataGridViewRow();

                    bool chkC = false;
                    bool chkG = false;

                    // 判斷領域是白名單內 check true
                    if (Global.DomainMap108List.Contains(name))
                    {
                        chkC = true;
                        chkG = true;
                    }


                    row.CreateCells(dgv, group, name, englishName, chkC, chkG);

                    dgv.Rows.Add(row);
                }
            }

            if (!containsIEPDomin)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgv, "", "特殊需求", "", false, false);
                dgv.Rows.Add(row);
            }
            if (!containsTechnologydomin)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgv, "", "科技", "", false, false);
                dgv.Rows.Add(row);
            }

        }

        /// <summary>
        /// 驗證領域，名稱不能空白、不能重覆。
        /// </summary>
        private void ValidDomain()
        {
            List<string> uniqueList = new List<string>();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;

                DataGridViewCell cell = row.Cells[chName.Index];
                string name = "" + cell.Value;

                if (string.IsNullOrEmpty(name))
                    cell.ErrorText = "領域名稱不能為空白";
                else if (uniqueList.Contains(name))
                    cell.ErrorText = "領域名稱不能重覆";
                else
                {
                    cell.ErrorText = string.Empty;
                    uniqueList.Add(name);
                }
            }
        }

        /// <summary>
        /// 檢查畫面上是否有錯誤訊息
        /// </summary>
        /// <returns></returns>
        private bool IsValid()
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                        return false;
                }
            }
            return true;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            ValidDomain();

            if (!IsValid())
            {
                MsgBox.Show("請先修正錯誤再儲存");
                return;
            }

            XmlDocument doc = new XmlDocument();
            doc.LoadXml("<Domains/>");

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement domainElement = doc.CreateElement("Domain");
                domainElement.SetAttribute("Group", "" + row.Cells[chGroup.Index].Value);
                domainElement.SetAttribute("Name", "" + row.Cells[chName.Index].Value);
                domainElement.SetAttribute("EnglishName", "" + row.Cells[chEnglishName.Index].Value);
                doc.DocumentElement.AppendChild(domainElement);
            }

            ConfigData cd = K12.Data.School.Configuration[ConfigName];
            cd.SetXml(ColumnKey, doc.DocumentElement);
            cd.Save();

            this.DialogResult = DialogResult.OK;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ValidDomain();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ImportExport.Export(dgv, "匯出領域清單");
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            ImportExport.Import(dgv, new List<string>() { "群組", "領域名稱", "領域英文名稱" });
            ValidDomain();
        }

        private void DomainListTable_Load(object sender, EventArgs e)
        {

        }
    }
}
