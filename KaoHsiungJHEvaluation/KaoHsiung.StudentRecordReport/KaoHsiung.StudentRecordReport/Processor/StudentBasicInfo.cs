using System;
using System.Collections.Generic;
using System.Text;
using JHSchool.Data;
using Aspose.Words;
using Aspose.Words.Drawing;
using K12.Data;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;

namespace KaoHsiung.StudentRecordReport.Processor
{
    class StudentBasicInfo
    {
        private List<string> _names;
        private Dictionary<string, string> _data;
        private DocumentBuilder _builder;
        private List<SemesterHistoryItem> _semesterHistoryList;

        private DAO.chkMaskInfo _maskInfo;

        public StudentBasicInfo(DocumentBuilder builder, DAO.chkMaskInfo maskInfo)
        {
            InitializeField();

            _builder = builder;
            _maskInfo = maskInfo;
            _builder.Document.MailMerge.FieldMergingCallback = new MailMerge_MergeField();
        }

        class MailMerge_MergeField : Aspose.Words.Reporting.IFieldMergingCallback
        {

            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                #region 處理照片
                if (e.FieldName == "照片")
                {
                    DocumentBuilder builder1 = new DocumentBuilder(e.Document);
                    builder1.MoveToField(e.Field, true);

                    byte[] photoBytes = null;
                    try
                    {
                        photoBytes = Convert.FromBase64String("" + e.FieldValue);
                    }
                    catch (Exception ex)
                    {
                        //builder1.Write("照片粘貼處");
                        e.Field.Remove();
                        return;
                    }

                    if (photoBytes == null || photoBytes.Length == 0)
                    {
                        //builder1.Write("照片粘貼處");
                        e.Field.Remove();
                        return;
                    }

                    e.Field.Remove();

                    Shape photoShape = new Shape(e.Document, ShapeType.Image);
                    photoShape.ImageData.SetImage(photoBytes);
                    photoShape.WrapType = WrapType.Inline;

                    #region AutoResize

                    double origHWRate = photoShape.ImageData.ImageSize.HeightPoints / photoShape.ImageData.ImageSize.WidthPoints;
                    double shapeHeight = (builder1.CurrentParagraph.ParentNode.ParentNode as Row).RowFormat.Height;
                    double shapeWidth = (builder1.CurrentParagraph.ParentNode as Cell).CellFormat.Width;
                    if ((shapeHeight / shapeWidth) < origHWRate)
                        shapeWidth = shapeHeight / origHWRate;
                    else
                        shapeHeight = shapeWidth * origHWRate;

                    #endregion

                    photoShape.Height = shapeHeight * 0.9;
                    photoShape.Width = shapeWidth * 0.9;

                    builder1.InsertNode(photoShape);
                }
                #endregion
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {

            }
        }

        public void SetStudent(JHStudentRecord student, List<SemesterHistoryItem> semesterHistoryList)
        {
            ClearField();
            // 學期歷程
            _semesterHistoryList = semesterHistoryList;
            JHParentRecord parent = JHParent.SelectByStudent(student);
            JHPhoneRecord phone = JHPhone.SelectByStudent(student);
            JHAddressRecord address = JHAddress.SelectByStudent(student);
            string tmpStr = "      ";

            JHLeaveInfoRecord leave = JHLeaveIfno.SelectByStudent(student);
            string number1 = string.Empty;
            string number2 = string.Empty;
            if (leave != null)
            {
                if (leave.Reason == "畢業")
                    number1 = leave.DiplomaNumber;
                else if (leave.Reason == "修業")
                    number2 = leave.DiplomaNumber;
            }

            #region 學生照片，若沒有畢業照，則印出入學照
            string base64 = K12.Data.Photo.SelectGraduatePhoto(student.ID);
            string freshmanPhoto = K12.Data.Photo.SelectFreshmanPhoto(student.ID);

            if (string.IsNullOrEmpty(base64))
                base64 = freshmanPhoto;

            #endregion



            _data["姓名"] = student.Name;
            _data["家長或監護人"] = (parent != null) ? parent.Custodian.Name : "";

            if (_data["姓名"] != null && _maskInfo.isMaskName)
            {
                int len = _data["姓名"].Length;
                if (len > 1)
                {
                    string mask = new String('○', len - 1);
                    _data["姓名"] = _data["姓名"].Substring(0, 1) + mask;
                }
            }

            if (_data["家長或監護人"] != null && _maskInfo.isMaskName)
            {
                int len = _data["家長或監護人"].Length;
                if (len > 1)
                {
                    string mask = new String('○', len - 1);
                    _data["家長或監護人"] = _data["家長或監護人"].Substring(0, 1) + mask;
                }
            }
            
            _data["身分證字號"] = student.IDNumber;
            if (_data["身分證字號"] != null && _maskInfo.isMaskIDNumber)
            {
                int len = _data["身分證字號"].Length;
                if (len > 1)
                {
                    string mask = new String('○', len - 1);
                    _data["身分證字號"] = _data["身分證字號"].Substring(0, 1) + mask;
                }
            }


            _data["性別"] = student.Gender;
            
            _data["學號"] = student.StudentNumber;
            _data["班級"] = (student.Class != null ? student.Class.Name : "");
            _data["座號"] = "" + student.SeatNo;

            if (_maskInfo.isMaskBirthday)
            {
                _data["出生"] = DateConvert.ChineseUnitDate(DateConvert.CDateM(student.Birthday.HasValue ? student.Birthday.Value.ToShortDateString() : ""));
            }
            else
            {
                _data["出生"] = DateConvert.ChineseUnitDate(DateConvert.CDate(student.Birthday.HasValue ? student.Birthday.Value.ToShortDateString() : ""));
            }
            
            _data["出生地"] = student.BirthPlace;
            
            _data["關係"] = (parent != null) ? parent.Custodian.Relationship : "";
            
            _data["聯絡電話"] = (phone != null) ? "" + phone.Contact : "";
            if (_data["聯絡電話"] != null && _maskInfo.isMaskPhone)
            {
                int len = _data["聯絡電話"].Length;
                if (len > 0)
                {
                    string mask = new String('○', len);
                    _data["聯絡電話"] =  mask;
                }
            }

            _data["戶籍電話"] = (phone != null) ? "" + phone.Permanent : "";
            if (_data["戶籍電話"] != null && _maskInfo.isMaskPhone)
            {
                int len = _data["戶籍電話"].Length;
                if (len > 0)
                {
                    string mask = new String('○', len);
                    _data["戶籍電話"] = mask;
                }
            }

            _data["戶籍地址"] = (address != null) ? address.Permanent.ToString() : "";
            if (_data["戶籍地址"] != null && _maskInfo.isMaskAddress)
            {
                int len = _data["戶籍地址"].Length;
                if (len > 0)
                {
                    string mask = new String('○', len);
                    _data["戶籍地址"] = mask;
                }
            }

            _data["通訊處"] = (address != null) ? address.Mailing.ToString() : "";
            if (_data["通訊處"] != null && _maskInfo.isMaskAddress)
            {
                int len = _data["通訊處"].Length;
                if (len > 0)
                {
                    string mask = new String('○', len);
                    _data["通訊處"] = mask;
                }
            }

            _data["行動電話"] = (phone != null) ? "" + phone.Cell : "";
            if (_data["行動電話"] != null && _maskInfo.isMaskPhone)
            {
                int len = _data["行動電話"].Length;
                if (len > 0)
                {
                    string mask = new String('○', len);
                    _data["行動電話"] = mask;
                }
            }

            _data["畢業證書字號"] = number1;
            _data["修業證明書字號"] = number2;
            _data["照片"] = base64;
            _data["簽呈"] = "承辦人員:" + Global.TransferName + tmpStr + "註冊組長:" + Global.RegManagerName + tmpStr + "教務主任:" + JHSchoolInfo.EduDirectorName + tmpStr + "校長:" + JHSchoolInfo.ChancellorChineseName;

            // 處理機敏資料



            // 處理服務學習時數
            if (Global._SLRDict.ContainsKey(student.ID))
            {
                foreach (SemesterHistoryItem shi in _semesterHistoryList)
                {
                    string key = shi.SchoolYear + "_" + shi.Semester;
                    if (Global._SLRDict[student.ID].ContainsKey(key))
                    {
                        string val = Global._SLRDict[student.ID][key];

                        if ((shi.GradeYear == 1 || shi.GradeYear == 7) && shi.Semester == 1)
                            _data["SLR1A"] = val;

                        if ((shi.GradeYear == 1 || shi.GradeYear == 7) && shi.Semester == 2)
                            _data["SLR1B"] = val;

                        if ((shi.GradeYear == 2 || shi.GradeYear == 8) && shi.Semester == 1)
                            _data["SLR2A"] = val;

                        if ((shi.GradeYear == 2 || shi.GradeYear == 8) && shi.Semester == 2)
                            _data["SLR2B"] = val;

                        if ((shi.GradeYear == 3 || shi.GradeYear == 9) && shi.Semester == 1)
                            _data["SLR3A"] = val;

                        if ((shi.GradeYear == 3 || shi.GradeYear == 9) && shi.Semester == 2)
                            _data["SLR3B"] = val;
                    }
                }
            }
            _builder.Document.MailMerge.Execute(GetFieldName(), GetFieldValue());
        }

        private void InitializeField()
        {
            _data = new Dictionary<string, string>();
            _data.Add("姓名", "");
            _data.Add("性別", "");
            _data.Add("身分證字號", "");
            _data.Add("學號", "");
            _data.Add("班級", "");
            _data.Add("座號", "");
            _data.Add("出生", "");
            _data.Add("出生地", "");
            _data.Add("家長或監護人", "");
            _data.Add("關係", "");
            _data.Add("聯絡電話", "");
            _data.Add("戶籍電話", "");
            _data.Add("戶籍地址", "");
            _data.Add("通訊處", "");
            _data.Add("行動電話", "");
            _data.Add("畢業證書字號", "");
            _data.Add("修業證明書字號", "");
            _data.Add("照片", "");
            _data.Add("簽呈", "");
            // 各學年度學期服務學習時數
            _data.Add("SLR1A", "0");
            _data.Add("SLR1B", "0");
            _data.Add("SLR2A", "0");
            _data.Add("SLR2B", "0");
            _data.Add("SLR3A", "0");
            _data.Add("SLR3B", "0");
            _names = new List<string>(_data.Keys);
        }

        private void ClearField()
        {
            foreach (string name in new List<string>(_data.Keys))
                _data[name] = "";
        }

        internal string[] GetFieldName()
        {
            return _names.ToArray();
        }

        internal string[] GetFieldValue()
        {
            List<string> values = new List<string>();
            foreach (string name in _names)
                values.Add(_data[name]);
            return values.ToArray();
        }
    }
}
