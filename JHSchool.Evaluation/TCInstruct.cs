using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Framework;
using FISCA.DSAUtil;
using JHSchool.Evaluation.Editor;
using FISCA.Presentation;
using JHSchool.Data;
using K12.Data;

namespace JHSchool.Evaluation
{
    /// <summary>
    /// 教師教授課程(課程與教師之間的多對多關連)。
    /// </summary>
    public class TCInstruct : CacheManager<TCInstructRecord>
    {
        private static TCInstruct _Instance = null;
        public static TCInstruct Instance { get { if (_Instance == null) _Instance = new TCInstruct(); return _Instance; } }

        private TCInstruct()
        {
            ItemLoaded += delegate
            {
                teacherField.Reload();
            };

            ItemUpdated += delegate(object sender, ItemUpdatedEventArgs e)
            {
                #region 更新(課程)及(教師)修課反查表
                // 2026/3/31 穎驊註記： 使用更精確的方式更新，只針對受影響的 ID 進行移除與重新加入
                foreach (var id in e.PrimaryKeys)
                {
                    // 移除舊的關聯
                    if (_id_course.ContainsKey(id))
                    {
                        string oldCourseID = _id_course[id];
                        if (_course_teachers.ContainsKey(oldCourseID))
                            _course_teachers[oldCourseID].Remove(id);
                    }

                    if (_id_teacher.ContainsKey(id))
                    {
                        string oldTeacherID = _id_teacher[id];
                        if (_teacher_courses.ContainsKey(oldTeacherID))
                            _teacher_courses[oldTeacherID].Remove(id);
                    }

                    // 加入新的關聯
                    var item = Items[id];
                    if (item != null)
                    {
                        if (!_course_teachers.ContainsKey(item.RefCourseID))
                            _course_teachers.Add(item.RefCourseID, new List<string>());
                        if (!_course_teachers[item.RefCourseID].Contains(id))
                            _course_teachers[item.RefCourseID].Add(id);

                        if (!_teacher_courses.ContainsKey(item.RefTeacherID))
                            _teacher_courses.Add(item.RefTeacherID, new List<string>());
                        if (!_teacher_courses[item.RefTeacherID].Contains(id))
                            _teacher_courses[item.RefTeacherID].Add(id);

                        _id_course[id] = item.RefCourseID;
                        _id_teacher[id] = item.RefTeacherID;
                    }
                    else
                    {
                        // 如果 item 為 null，表示是被刪除
                        _id_course.Remove(id);
                        _id_teacher.Remove(id);
                    }
                }

                teacherField.Reload();
                #endregion
            };
        }

        private bool _initialized = false;
        private ListPaneField teacherField;
        private RibbonBarButton assignTeacherButton;

        /// <summary>
        /// 設定使用者介面
        /// </summary>
        public void SetupPresentation()
        {
            if (_initialized) return;

            #region 課程加入授課教師
            Course.Instance.SelectedListChanged += delegate { assignTeacherButton.Enable = (Course.Instance.SelectedList.Count > 0 && Teacher.Instance.TemporaList.Count > 0); };
            Teacher.Instance.TemporaListChanged += delegate { assignTeacherButton.Enable = (Course.Instance.SelectedList.Count > 0 && Teacher.Instance.TemporaList.Count > 0); };

            assignTeacherButton = Course.Instance.RibbonBarItems["指定"]["評分教師"];
            assignTeacherButton.Enable = false;
            assignTeacherButton.Image = Properties.Resources.teacher_64;
            MenuButton loadingMenuButton = assignTeacherButton["載入中…"];
            assignTeacherButton.PopupOpen += delegate(object sender, PopupOpenEventArgs e)
            {
                if (Teacher.Instance.TemporaList.Count <= 0) return;
                if (Course.Instance.SelectedList.Count <= 0) return;

                loadingMenuButton.Visible = false;

                foreach (var item in Teacher.Instance.TemporaList)
                {
                    MenuButton mb = e.VirtualButtons[item.Name];
                    mb.Tag = item;
                    mb.Click += new EventHandler(MenuButton_Click);
                }
            };
            #endregion

            #region 授課教師 ListPanelField
            teacherField = new ListPaneField("授課教師");
            teacherField.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                string teacherName = "";
                if (TCInstruct.Instance.Loaded)
                {
                    TeacherRecord teacher1 = Course.Instance[e.Key].GetFirstTeacher();
                    if (teacher1 != null)
                        teacherName += teacher1.FullName;

                    TeacherRecord teacher2 = Course.Instance[e.Key].GetSecondTeacher();
                    if (teacher2 != null)
                    {
                        if (teacherName.Length > 0)
                            teacherName += "," + teacher2.FullName;
                        else
                            teacherName += teacher2.FullName;
                    }
                    TeacherRecord teacher3 = Course.Instance[e.Key].GetThirdTeacher();
                    if (teacher3 != null)
                    {
                        if (teacherName.Length > 0)
                            teacherName += "," + teacher3.FullName;
                        else
                            teacherName += teacher3.FullName;
                    }
                    if (teacherName.Length > 0)
                        e.Value = teacherName;
                    else
                        e.Value = string.Empty;
                }
                else
                    e.Value = "Loading...";
            };

            Course.Instance.AddListPaneField(teacherField);
            #endregion

            _initialized = true;
        }

        private void MenuButton_Click(object sender, EventArgs e)
        {
            TeacherRecord teacher = (sender as MenuButton).Tag as TeacherRecord;

            List<TCInstructRecordEditor> editors = new List<TCInstructRecordEditor>();
            foreach (var item in Course.Instance.SelectedList)
                editors.Add(item.SetFirstTeacher(teacher));

            if (editors.Count > 0)
            {
                MultiThreadBackgroundWorker<TCInstructRecordEditor> worker = new MultiThreadBackgroundWorker<TCInstructRecordEditor>();
                worker.PackageSize = 50;
                worker.Loading = MultiThreadLoading.Light;
                worker.DoWork += delegate(object worker_sender, PackageDoWorkEventArgs<TCInstructRecordEditor> worker_e)
                {
                    worker_e.Items.SaveAllEditors();
                };
                worker.RunWorkerCompleted += delegate
                {
                    MsgBox.Show("指定評分教師完成");
                };
                worker.RunWorkerAsync(editors);
            }
        }

        protected override Dictionary<string, TCInstructRecord> GetAllData()
        {
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("All");

            DSRequest dsreq = new DSRequest(helper);
            Dictionary<string, TCInstructRecord> result = new Dictionary<string, TCInstructRecord>();
            string srvname = "SmartSchool.Course.GetTCInstruct";

            var course_teachers = new Dictionary<string, List<string>>();
            var teacher_courses = new Dictionary<string, List<string>>();
            var id_course = new Dictionary<string, string>();
            var id_teacher = new Dictionary<string, string>();

            foreach (var item in FISCA.Authentication.DSAServices.CallService(srvname, dsreq).GetContent().GetElements("TCInstruct"))
            {
                var teacherid = item.SelectSingleNode("RefTeacherID")?.InnerText;
                var courseid = item.SelectSingleNode("RefCourseID")?.InnerText;
                var id = item.GetAttribute("ID");
                var sequence = item.SelectSingleNode("Sequence")?.InnerText;

                TCInstructRecord record = new TCInstructRecord(teacherid, courseid, id, sequence);
                result.Add(record.ID, record);

                if (!string.IsNullOrEmpty(courseid))
                {
                    if (!course_teachers.ContainsKey(courseid))
                        course_teachers.Add(courseid, new List<string>());
                    course_teachers[courseid].Add(id);
                    id_course[id] = courseid;
                }

                if (!string.IsNullOrEmpty(teacherid))
                {
                    if (!teacher_courses.ContainsKey(teacherid))
                        teacher_courses.Add(teacherid, new List<string>());
                    teacher_courses[teacherid].Add(id);
                    id_teacher[id] = teacherid;
                }
            }

            _course_teachers = course_teachers;
            _teacher_courses = teacher_courses;
            _id_course = id_course;
            _id_teacher = id_teacher;

            return result;
        }

        protected override Dictionary<string, TCInstructRecord> GetData(IEnumerable<string> primaryKeys)
        {
            // 指示是否需要呼叫 Service。
            bool execute_require = false;

            //建立 Request。
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("All");
            helper.AddElement("Condition");
            foreach (string id in primaryKeys)
            {
                helper.AddElement("Condition", "ID", id);
                execute_require = true;
            }

            //儲存最後結果的集合。
            Dictionary<string, TCInstructRecord> result = new Dictionary<string, TCInstructRecord>();

            if (execute_require)
            {
                string srvname = "SmartSchool.Course.GetTCInstruct";
                DSRequest dsreq = new DSRequest(helper);

                foreach (var item in FISCA.Authentication.DSAServices.CallService(srvname, dsreq).GetContent().GetElements("TCInstruct"))
                {
                    var teacherid = item.SelectSingleNode("RefTeacherID")?.InnerText;
                    var courseid = item.SelectSingleNode("RefCourseID")?.InnerText;
                    var id = item.GetAttribute("ID");
                    var sequence = item.SelectSingleNode("Sequence")?.InnerText;

                    TCInstructRecord record = new TCInstructRecord(teacherid, courseid, id, sequence);
                    result.Add(record.ID, record);
                }
            }

            return result;
        }

        /// <summary>
        /// 從(課程)查詢授課教師。
        /// </summary>
        private Dictionary<string, List<string>> _course_teachers = new Dictionary<string, List<string>>();

        /// <summary>
        /// 從(教師)查詢教授課程。
        /// </summary>
        private Dictionary<string, List<string>> _teacher_courses = new Dictionary<string, List<string>>();

        /// <summary>
        /// 記錄 ID 與課程 ID 的對照表。
        /// </summary>
        private Dictionary<string, string> _id_course = new Dictionary<string, string>();

        /// <summary>
        /// 記錄 ID 與教師 ID 的對照表。
        /// </summary>
        private Dictionary<string, string> _id_teacher = new Dictionary<string, string>();

        /// <summary>
        /// 取得課程的所有授課教師。
        /// </summary>
        /// <param name="courseID">課程編號。</param>
        /// <returns>授課教師清單。</returns>
        public List<TCInstructRecord> GetCourseTeachers(string courseID)
        {
            List<TCInstructRecord> result = new List<TCInstructRecord>();
            if (_course_teachers.ContainsKey(courseID))
            {
                foreach (var eachInstructID in _course_teachers[courseID])
                {
                    var objInstruct = Items[eachInstructID];
                    if (objInstruct.Course != null && objInstruct.Teacher != null)
                        result.Add(Items[eachInstructID]);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得教師所授教的課程。
        /// </summary>
        /// <param name="studentID">學生編號</param>
        /// <returns>修課記錄清單</returns>
        public List<TCInstructRecord> GetTeacherCourses(string studentID)
        {
            List<TCInstructRecord> result = new List<TCInstructRecord>();
            if (_teacher_courses.ContainsKey(studentID))
            {
                foreach (var eachInstructID in _teacher_courses[studentID])
                {
                    var objInstruct = Items[eachInstructID];
                    if (objInstruct.Course != null && objInstruct.Teacher != null)
                        result.Add(Items[eachInstructID]);
                }
            }
            return result;
        }
    }

    public static class TCInstruct_ExtendMethods
    {
        /// <summary>
        /// 取得課程的第一位授課教師。
        /// </summary>
        public static TeacherRecord GetFirstTeacher(this CourseRecord course)
        {
            if (course != null)
            {
                TCInstructRecord tc = course.GetFirstInstruct();
                if (tc != null) return tc.Teacher;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第二位授課教師。
        /// </summary>
        public static TeacherRecord GetSecondTeacher(this CourseRecord course)
        {
            if (course != null)
            {
                TCInstructRecord tc = course.GetSecondInstruct();
                if (tc != null) return tc.Teacher;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第三位授課教師。
        /// </summary>
        public static TeacherRecord GetThirdTeacher(this CourseRecord course)
        {
            if (course != null)
            {
                TCInstructRecord tc = course.GetThirdInstruct();
                if (tc != null) return tc.Teacher;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第一位授課教師。
        /// </summary>
        internal static TCInstructRecord GetFirstInstruct(this CourseRecord course)
        {
            if (course != null)
            {
                foreach (TCInstructRecord each in course.GetInstructs())
                    if (each.Sequence == "1") return each;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第二位授課教師。
        /// </summary>
        internal static TCInstructRecord GetSecondInstruct(this CourseRecord course)
        {
            if (course != null)
            {
                foreach (TCInstructRecord each in course.GetInstructs())
                    if (each.Sequence == "2") return each;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第三位授課教師。
        /// </summary>
        internal static TCInstructRecord GetThirdInstruct(this CourseRecord course)
        {
            if (course != null)
            {
                foreach (TCInstructRecord each in course.GetInstructs())
                    if (each.Sequence == "3") return each;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的所有上課教師關聯資料。
        /// </summary>
        public static List<TCInstructRecord> GetInstructs(this CourseRecord course)
        {
            if (course != null)
                return TCInstruct.Instance.GetCourseTeachers(course.ID);
            else
                return null;
        }

        /// <summary>
        /// 取得教師上的所有課程關聯資料。 
        /// </summary>
        public static List<TCInstructRecord> GetInstructs(this TeacherRecord teacher)
        {

            return TCInstruct.Instance.GetTeacherCourses(teacher.ID);
        }

        /// <summary>
        /// 取得課程的所有教師資料。
        /// </summary>
        public static List<TeacherRecord> GetInstructTeachers(this CourseRecord course)
        {
            if (course != null)
            {
                List<TeacherRecord> teachers = new List<TeacherRecord>();
                foreach (TCInstructRecord each in GetInstructs(course))
                    teachers.Add(each.Teacher);

                return teachers;
            }
            else
                return null;
        }

        /// <summary>
        /// 取得教師上的所有課程資料。
        /// </summary>
        public static List<CourseRecord> GetInstructCoruses(this TeacherRecord teacher)
        {
            List<CourseRecord> courses = new List<CourseRecord>();
            foreach (TCInstructRecord each in GetInstructs(teacher))
                courses.Add(each.Course);

            return courses;
        }
    }
}
