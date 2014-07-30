using Aspose.Cells;
using CourseGradeB;
using FISCA.Data;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClassExamSemester
{
    public partial class frm_printsetup : BaseForm
    {
        private List<string> _class_ids;
        private QueryHelper _Q;
        private int _schoolYear, _semester;

        public frm_printsetup(List<string> select_class_ids)
        {
            InitializeComponent();
            _class_ids = select_class_ids;
            _Q = new QueryHelper();
            //設定ComboBox清單數字
            _schoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _semester = int.Parse(K12.Data.School.DefaultSemester);

            //設定comBoxSchoolYear前二後二學年度
            for (int i = -2; i <= 2; i++)
            {
                comBoxSchoolYear.Items.Add(_schoolYear + i);
            }
            //設定comBoxSemester值
            comBoxSemester.Items.Add(1);
            comBoxSemester.Items.Add(2);
            comBoxSchoolYear.Text = _schoolYear + "";
            comBoxSemester.Text = _semester + "";
            //設定comBoxPeriod值
            comBoxPeriod.Items.Add("Midterm");
            comBoxPeriod.Items.Add("Final");
        }

        private string SelectTime() //取得Server的時間
        {
            QueryHelper Sql = new QueryHelper();
            DataTable dtable = Sql.Select("select now()"); //取得時間
            DateTime dt = DateTime.Now;
            DateTime.TryParse("" + dtable.Rows[0][0], out dt); //Parse資料
            string ComputerSendTime = dt.ToString("yyyy/MM/dd"); //最後時間

            return ComputerSendTime;
        }

        private void btn_yes_Click(object sender, EventArgs e)
        {
            //處理comboBox的例外狀況
            int SchoolYear;
            int Semester;
            if (!int.TryParse(comBoxSchoolYear.Text, out SchoolYear))
            {
                MsgBox.Show("學年度必須選擇為數字");
                return;
            }
            if (!int.TryParse(comBoxSemester.Text, out Semester))
            {
                MsgBox.Show("學期必須選擇為數字");
                return;
            }

            //取得班級物件
            List<ClassRecord> class_list = K12.Data.Class.SelectByIDs(_class_ids);
            Dictionary<string, List<string>> class_subject_list = new Dictionary<string, List<string>>();

            List<String> sID = new List<string>();
            //取得各班學生
            foreach (ClassRecord record in class_list)
            {
                //班級科目清單
                if (!class_subject_list.ContainsKey(record.ID))
                    class_subject_list.Add(record.ID, new List<string>());

                //該班級的所有學生存入students
                List<StudentRecord> students = record.Students;
                //整理各班學生資料型態變為List<String>
                foreach (StudentRecord sr in students)
                {
                    if (!sID.Contains(sr.ID))
                        sID.Add(sr.ID);
                }
            }
            //判斷選取的試別
            int period = 0;
            if (comBoxPeriod.Text == "Midterm")
            {
                period = 1;
            }
            else
            {
                period = 2;
            }

            string student_ID = string.Join(",", sID);
            //組SQL語法抓取學生ID、科目、成績
            string sql = "select student.name,$ischool.course.extend.grade_year as level,class.id as class_id,course.credit,sc_attend.ref_student_id,course.subject,$ischool.subject.list.group,xpath_string(sce_take.extension,'//Extension/Score') as score from sc_attend";
            sql += " join course on sc_attend.ref_course_id=course.id";
            sql += " join sce_take on sce_take.ref_sc_attend_id = sc_attend.id";
            sql += " join student on student.id=sc_attend.ref_student_id";
            sql += " join class on class.id=student.ref_class_id";
            sql += " left join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id";
            sql += " left join $ischool.subject.list on $ischool.subject.list.name=course.subject";
            sql += " where sc_attend.ref_student_id in (" + student_ID + ") and sce_take.ref_exam_id=" + period + " and course.school_year=" + _schoolYear + " and course.semester=" + _semester + "and student.status=1";
            DataTable dt = _Q.Select(sql);

            Dictionary<string, StudentObj> stuObjDic = new Dictionary<string, StudentObj>();
            foreach (DataRow row in dt.Rows)
            {
                string student_id = row["ref_student_id"] + "";
                string class_id = row["class_id"] + "";

                if (!stuObjDic.ContainsKey(student_id))
                    stuObjDic.Add(student_id, new StudentObj(row));

                stuObjDic[student_id].LoadData(row);
            }

            //Rank
            foreach (ClassRecord c in class_list)
            {
                List<StudentObj> score_list = new List<StudentObj>();
                foreach (StudentRecord s in c.Students)
                {
                    if (stuObjDic.ContainsKey(s.ID))
                        score_list.Add(stuObjDic[s.ID]);
                }

                score_list.Sort(delegate(StudentObj x, StudentObj y)
                {
                    return x.Avg.CompareTo(y.Avg);
                });

                score_list.Reverse();

                int rank = 0;
                int count = 0;
                decimal temp_score = decimal.MinValue;
                foreach (StudentObj obj in score_list)
                {
                    count++;

                    if (temp_score != obj.Avg)
                    {
                        rank = count;
                    }
                    obj.Rank = rank;
                    temp_score = obj.Avg;
                }
            }

            //各班科目清單
            foreach (string key in stuObjDic.Keys)
            {
                StudentObj obj = stuObjDic[key];
                foreach (string subj in obj.Subject_Score.Keys)
                {
                    if (class_subject_list.ContainsKey(obj.ClassID))
                        if (!class_subject_list[obj.ClassID].Contains(subj))
                            class_subject_list[obj.ClassID].Add(subj);
                }
            }

            //各班科目清單排序
            foreach (string key in class_subject_list.Keys)
                class_subject_list[key].Sort(Tool.GetSubjectCompare());

            //各班科目Columns Index
            Dictionary<string, int> class_columns_index = new Dictionary<string, int>();
            foreach (string key in class_subject_list.Keys)
            {
                int index = 2;
                foreach (string subj in class_subject_list[key])
                {
                    if (!class_columns_index.ContainsKey(key + "_" + subj))
                    {
                        class_columns_index.Add(key + "_" + subj, index);
                        index++;
                    }
                }
            }

            Workbook wb = new Workbook();
            wb.Open(new MemoryStream(Properties.Resources.Template));
            Worksheet template_sheet = wb.Worksheets["Template"];

            int sheet_index = 1;
            //每個班級列印一個Worksheet
            foreach (ClassRecord cls in class_list)
            {
                sheet_index = wb.Worksheets.AddCopy("Template");
                wb.Worksheets[sheet_index].Name = cls.Name;

                wb.Worksheets[sheet_index].Cells[0, 0].PutValue("雙語部  " + _schoolYear + "年度第" + _semester + "學期  評量成績");
                wb.Worksheets[sheet_index].Cells[1, 0].PutValue(" Class：" + cls.Name + "       Period：" + comBoxPeriod.Text);
                wb.Worksheets[sheet_index].Cells[2, 0].PutValue("SeatNo");
                wb.Worksheets[sheet_index].Cells[2, 1].PutValue("Name");

                //設定Style樣板：四邊框線 水平垂直字中 自動換行
                Style s = wb.Styles[wb.Styles.Add()];
                s.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                s.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                s.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                s.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                s.HorizontalAlignment = TextAlignmentType.Center;
                s.VerticalAlignment = TextAlignmentType.Center;
                s.IsTextWrapped = true;
                //設定Style2樣板：四邊框線
                Style s2 = wb.Styles[wb.Styles.Add()];
                s2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                s2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                s2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                s2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                wb.Worksheets[sheet_index].Cells[2, 0].Style = s;
                wb.Worksheets[sheet_index].Cells[2, 1].Style = s;

                //列印科目名稱
                int indexSub = 2;
                foreach (String sub in class_subject_list[cls.ID])
                {
                    wb.Worksheets[sheet_index].Cells[2, indexSub].Style = s;
                    wb.Worksheets[sheet_index].Cells[2, indexSub].PutValue(sub);
                    indexSub++;
                }
                wb.Worksheets[sheet_index].Cells[2, indexSub].Style = s;
                wb.Worksheets[sheet_index].Cells[2, indexSub].PutValue("Avg");
                wb.Worksheets[sheet_index].Cells[2, indexSub + 1].Style = s;
                wb.Worksheets[sheet_index].Cells[2, indexSub + 1].PutValue("Rank");
                wb.Worksheets[sheet_index].Cells[2, indexSub + 2].Style = s;
                wb.Worksheets[sheet_index].Cells[2, indexSub + 2].PutValue("Level");
                //合併儲存格：First Row合併 ；Second Row Column 前三後三合併
                wb.Worksheets[sheet_index].Cells.Merge(0, 0, 1, indexSub + 3); 
                wb.Worksheets[sheet_index].Cells.Merge(1, 0, 1, 3); 
                wb.Worksheets[sheet_index].Cells.Merge(1, indexSub, 1, 3);

                wb.Worksheets[sheet_index].Cells[1, indexSub].PutValue("列印日期：" + SelectTime());

                //列印學生
                int indexrow = 3;
                foreach (StudentRecord student in cls.Students)
                {
                    wb.Worksheets[sheet_index].Cells[indexrow, 0].Style = s;
                    wb.Worksheets[sheet_index].Cells[indexrow, 0].PutValue(student.SeatNo);
                    wb.Worksheets[sheet_index].Cells[indexrow, 1].Style = s2;
                    wb.Worksheets[sheet_index].Cells[indexrow, 1].PutValue(student.Name);
                    //列印學生各科成績
                    foreach (string subj in class_subject_list[cls.ID])
                    {
                        int column_index = class_columns_index[cls.ID + "_" + subj];
                        string score = "";

                        if (stuObjDic.ContainsKey(student.ID))
                            if (stuObjDic[student.ID].Subject_Score.ContainsKey(subj))
                                score = stuObjDic[student.ID].Subject_Score[subj] + "";
                        wb.Worksheets[sheet_index].Cells[indexrow, column_index].Style = s;
                        wb.Worksheets[sheet_index].Cells[indexrow, column_index].PutValue(score);
                    }

                    StudentObj obj = null;
                    if (stuObjDic.ContainsKey(student.ID))
                        obj = stuObjDic[student.ID];

                    wb.Worksheets[sheet_index].Cells[indexrow, indexSub].Style = s;
                    wb.Worksheets[sheet_index].Cells[indexrow, indexSub + 1].Style = s;
                    wb.Worksheets[sheet_index].Cells[indexrow, indexSub + 2].Style = s;


                    if (obj != null)
                    {
                        wb.Worksheets[sheet_index].Cells[indexrow, indexSub].PutValue(obj.Avg);
                        wb.Worksheets[sheet_index].Cells[indexrow, indexSub + 1].PutValue(obj.Rank);
                        wb.Worksheets[sheet_index].Cells[indexrow, indexSub + 2].PutValue(obj.Level);
                    }
                    indexrow++;
                }
                //sheet_index++;
            }

            wb.Worksheets.RemoveAt(0);
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "PeriodGrade_Semester.xls";
            save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(save.FileName, Aspose.Cells.FileFormatType.Excel2003);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }
        private void btn_no_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
