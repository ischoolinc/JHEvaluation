﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using System.IO;

namespace HsinChuSemesterClassFixedRank
{
    [FISCA.UDT.TableName(Global._UDTTableName)]
    public class Configure : ActiveRecord
    {
        public Configure()
        {
            PrintSubjectList = new List<string>();
        }

        /// <summary>
        /// 設定檔名稱
        /// </summary>
        [FISCA.UDT.Field]
        public string Name { get; set; }


        /// <summary>
        /// 學年度
        /// </summary>
        [FISCA.UDT.Field]
        public string SchoolYear { get; set; }
        /// <summary>
        /// 學期
        /// </summary>
        [FISCA.UDT.Field]
        public string Semester { get; set; }

        /// <summary>
        /// 列印樣板
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream { get; set; }
        public Aspose.Words.Document Template { get; set; }

        /// <summary>
        /// 樣板中支援列印科目的最大數
        /// </summary>
        [FISCA.UDT.Field]
        public int SubjectLimit { get; set; }
        
        /// <summary>
        /// 列印科別
        /// </summary>
        [FISCA.UDT.Field]
        private string PrintSubjectListString { get; set; }
        public List<string> PrintSubjectList { get; set; }


        /// <summary>
        /// 列印領域
        /// </summary>
        [FISCA.UDT.Field]
        private string PrintDomainListString { get; set; }
        public List<string> PrintDomainList { get; set; }

        
        /// <summary>
        /// 列印時選樣板設定檔
        /// </summary>
        [FISCA.UDT.Field]
        public string SelSetConfigName { get; set; }

        /// <summary>
        /// 需補考
        /// </summary>
        [FISCA.UDT.Field]
        public string NeeedReScoreMark { get; set; }

        /// <summary>
        /// 補考成績
        /// </summary>
        [FISCA.UDT.Field]
        public string ReScoreMark { get; set; }

        /// <summary>
        /// 不及格
        /// </summary>
        [FISCA.UDT.Field]
        public string FailScoreMark { get; set; }





        /// <summary>
        /// 在儲存前，把資料填入儲存欄位中
        /// </summary>
        public void Encode()
        {
          
            // 科目
            this.PrintSubjectListString = "";
            if (this.PrintSubjectList == null)
                this.PrintSubjectList = new List<string>();
            this.PrintDomainListString = "";
            if (this.PrintDomainList == null)
                this.PrintDomainList = new List<string>();

            foreach (var item in this.PrintSubjectList)
            {
                this.PrintSubjectListString += (this.PrintSubjectListString == "" ? "" : "^^^") + item;
            }

            foreach (var item in this.PrintDomainList)
            {
                this.PrintDomainListString += (this.PrintDomainListString == "" ? "" : "^^^") + item;
            }

            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            this.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
            this.TemplateStream = Convert.ToBase64String(stream.ToArray());
        }
        /// <summary>
        /// 在資料取出後，把資料從儲存欄位轉換至資料欄位
        /// </summary>
        public void Decode()
        {
            

            // 科目
            if (!string.IsNullOrEmpty(this.PrintSubjectListString))
                this.PrintSubjectList = new List<string>(this.PrintSubjectListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            else
                this.PrintSubjectList = new List<string>();

            // 領域
            if (!string.IsNullOrEmpty(this.PrintDomainListString))
                this.PrintDomainList = new List<string>(this.PrintDomainListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            else
                this.PrintDomainList = new List<string>();

            this.Template = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream)));
        }
    }
}
