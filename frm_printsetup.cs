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
        private List<string> classIds;
        private QueryHelper queryHelper;
        private int schoolYear, semester;

        public frm_printsetup(List<string> selectClassIds)
        {
            InitializeComponent();
			classIds = selectClassIds;
            this.queryHelper = new QueryHelper();
            //設定ComboBox清單數字
            schoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            semester = int.Parse(K12.Data.School.DefaultSemester);

            //設定comBoxSchoolYear前二後二學年度
            for (int i = -2; i <= 2; i++){
                comBoxSchoolYear.Items.Add(schoolYear + i);
            }
            //設定comBoxSemester值
            comBoxSemester.Items.Add(1);
            comBoxSemester.Items.Add(2);
            comBoxSchoolYear.Text = schoolYear + "";
            comBoxSemester.Text = semester + "";
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
            int schoolYear;
            int semester;
            if (!int.TryParse(comBoxSchoolYear.Text, out schoolYear)){
                MsgBox.Show("學年度必須選擇為數字");
                return;
            }
            if (!int.TryParse(comBoxSemester.Text, out semester)){
                MsgBox.Show("學期必須選擇為數字");
                return;
            }

            //取得班級物件
			List<ClassRecord> classList = K12.Data.Class.SelectByIDs(classIds);
            Dictionary<string, List<string>> classSubjectList = new Dictionary<string, List<string>>();

            List<String> studentIds = new List<string>();
            //取得各班學生
            foreach (ClassRecord record in classList){
                //班級科目清單
				if (!classSubjectList.ContainsKey(record.ID)) {
					classSubjectList.Add(record.ID, new List<string>());
				}
                   

                //該班級的所有學生存入students
                List<StudentRecord> studentList = record.Students;
                //整理各班學生資料型態變為List<String>
                foreach (StudentRecord sr in studentList){
					if (!studentIds.Contains(sr.ID)) {
						studentIds.Add(sr.ID);
					}   
                }
            }
            //判斷選取的試別
            int period = 0;
            if (comBoxPeriod.Text == "Midterm"){
                period = 1;
            }
            else{
                period = 2;
            }

            string strStudentId = string.Join(",", studentIds);

           if (strStudentId == ""){
                MsgBox.Show("此班級無學生，請確認班級學生");
                return;
            }
            //組SQL語法抓取學生ID、科目、成績
            string sql = "select student.name,$ischool.course.extend.grade_year as level,class.id as class_id,course.credit,sc_attend.ref_student_id,course.subject,$ischool.subject.list.group,xpath_string(sce_take.extension,'//Extension/Score') as score from sc_attend";
            sql += " join course on sc_attend.ref_course_id=course.id";
            sql += " join sce_take on sce_take.ref_sc_attend_id = sc_attend.id";
            sql += " join student on student.id=sc_attend.ref_student_id";
            sql += " join class on class.id=student.ref_class_id";
            sql += " left join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id";
            sql += " left join $ischool.subject.list on $ischool.subject.list.name=course.subject";
            sql += " where sc_attend.ref_student_id in (" + strStudentId + ") and sce_take.ref_exam_id=" + period + " and course.school_year=" + schoolYear + " and course.semester=" + semester + "and student.status=1";
            DataTable dt = queryHelper.Select(sql);
            
            Dictionary<string, StudentObj> stuObjDic = new Dictionary<string, StudentObj>();
            foreach (DataRow row in dt.Rows){
                string studentId = row["ref_student_id"] + "";
                string classId = row["class_id"] + "";

				if (!stuObjDic.ContainsKey(studentId)) {
					stuObjDic.Add(studentId, new StudentObj(row));
				}
                stuObjDic[studentId].LoadData(row);
            }
            
            //Rank
            foreach (ClassRecord cr in classList){
                List<StudentObj> scoreList = new List<StudentObj>();
                foreach (StudentRecord sr in cr.Students) {
					if (stuObjDic.ContainsKey(sr.ID)) {
						scoreList.Add(stuObjDic[sr.ID]);
					}  
                }

                scoreList.Sort(delegate(StudentObj x, StudentObj y){
                    return x.Avg.CompareTo(y.Avg);
                });

                scoreList.Reverse();

                int rank = 0;
                int count = 0;
                decimal temp_score = decimal.MinValue;
                foreach (StudentObj obj in scoreList){
                    count++;

                    if (temp_score != obj.Avg){
                        rank = count;
                    }
                    obj.Rank = rank;
                    temp_score = obj.Avg;
                }
            }

            //各班科目清單
            foreach (string key in stuObjDic.Keys){
                StudentObj obj = stuObjDic[key];
                foreach (string subj in obj.Subject_Score.Keys){
					if (classSubjectList.ContainsKey(obj.ClassID)) {
						if (!classSubjectList[obj.ClassID].Contains(subj)) {
							classSubjectList[obj.ClassID].Add(subj);
						}
					}
                }
            }

            //各班科目清單排序
			foreach (string key in classSubjectList.Keys) {
				classSubjectList[key].Sort(Tool.GetSubjectCompare());
			}
            //各班科目Columns Index
            Dictionary<string, int> classColumnsIndex = new Dictionary<string, int>();
            foreach (string key in classSubjectList.Keys){
                int index = 2;
                foreach (string subj in classSubjectList[key]){
                    if (!classColumnsIndex.ContainsKey(key + "_" + subj)){
                        classColumnsIndex.Add(key + "_" + subj, index);
                        index++;
                    }
                }
            }

            Workbook wb = new Workbook();
            wb.Open(new MemoryStream(Properties.Resources.Template));
            Worksheet templateSheet = wb.Worksheets["Template"];

            int sheetIndex = 1;
            //每個班級列印一個Worksheet
            foreach (ClassRecord cls in classList){
                sheetIndex = wb.Worksheets.AddCopy("Template");
                wb.Worksheets[sheetIndex].Name = cls.Name;

				wb.Worksheets[sheetIndex].Cells[0, 0].PutValue("雙語部  " + (schoolYear + 1911) + "~" + (schoolYear + 1912) + "年度第" + semester + "學期  評量成績");
                wb.Worksheets[sheetIndex].Cells[1, 0].PutValue(" Class：" + cls.Name + "       Period：" + comBoxPeriod.Text);
                wb.Worksheets[sheetIndex].Cells[2, 0].PutValue("SeatNo");
                wb.Worksheets[sheetIndex].Cells[2, 1].PutValue("Name");

                //設定Style樣板：四邊框線 水平垂直字中 自動換行
                Style s = wb.Styles[wb.Styles.Add()];
                s.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                s.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                s.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                s.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                s.HorizontalAlignment = TextAlignmentType.Center;
                s.VerticalAlignment = TextAlignmentType.Center;
                s.IsTextWrapped = true;

                //設定Style2樣板：三邊細線 底線粗線 水平垂直字中 自動換行
                Style s2 = wb.Styles[wb.Styles.Add()];
                s2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                s2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                s2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                s2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                s2.HorizontalAlignment = TextAlignmentType.Center;
                s2.VerticalAlignment = TextAlignmentType.Center;
                s2.IsTextWrapped = true;

                //設定Style3樣板：三邊細線 右線粗線 水平字左 垂直字中 自動換行
                Style s3 = wb.Styles[wb.Styles.Add()];
                s3.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                s3.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                s3.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                s3.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
                s3.HorizontalAlignment = TextAlignmentType.Left;
                s3.VerticalAlignment = TextAlignmentType.Center;

                //設定Style4樣板：兩邊細線 右邊底線粗線 水平字左 垂直字中
                Style s4 = wb.Styles[wb.Styles.Add()];
                s4.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                s4.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                s4.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                s4.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
                s4.HorizontalAlignment = TextAlignmentType.Left;
                s4.VerticalAlignment = TextAlignmentType.Center;

                //列印科目名稱
                int indexSub = 2; //前面有NO & Name，所以從indexSub = 2
                foreach (String sub in classSubjectList[cls.ID]){
                    wb.Worksheets[sheetIndex].Cells[2, indexSub].PutValue(sub);
                    indexSub++;
                }
                wb.Worksheets[sheetIndex].Cells[2, indexSub].PutValue("Avg");
                wb.Worksheets[sheetIndex].Cells[2, indexSub + 1].PutValue("Rank");
                wb.Worksheets[sheetIndex].Cells[2, indexSub + 2].PutValue("Level");
                //合併儲存格：First Row合併 ；Second Row Column 前三後三合併
                wb.Worksheets[sheetIndex].Cells.Merge(0, 0, 1, indexSub + 3);
                wb.Worksheets[sheetIndex].Cells.Merge(1, 0, 1, 3);
                wb.Worksheets[sheetIndex].Cells.Merge(1, indexSub, 1, 3);

                wb.Worksheets[sheetIndex].Cells[1, indexSub].PutValue("列印日期：" + SelectTime());

                //列印學生
                int indexRow = 3;
                foreach (StudentRecord student in cls.Students){
                    wb.Worksheets[sheetIndex].Cells[indexRow, 0].Style = s;
                    wb.Worksheets[sheetIndex].Cells[indexRow, 0].PutValue(student.SeatNo);
                    wb.Worksheets[sheetIndex].Cells[indexRow, 1].Style = s;
                    wb.Worksheets[sheetIndex].Cells[indexRow, 1].PutValue(student.Name + "  " + student.EnglishName);
                    //列印學生各科成績
                    foreach (string subj in classSubjectList[cls.ID]){
                        int column_index = classColumnsIndex[cls.ID + "_" + subj];
                        string score = "";

						if (stuObjDic.ContainsKey(student.ID)){
							if (stuObjDic[student.ID].Subject_Score.ContainsKey(subj)) {
								score = stuObjDic[student.ID].Subject_Score[subj] + "";
							}	
						}
                        wb.Worksheets[sheetIndex].Cells[indexRow, column_index].Style = s;
                        wb.Worksheets[sheetIndex].Cells[indexRow, column_index].PutValue(score);
                    }

                    StudentObj obj = null;
					if (stuObjDic.ContainsKey(student.ID)){
						obj = stuObjDic[student.ID];
					}
                    wb.Worksheets[sheetIndex].Cells[indexRow, indexSub].Style = s;
                    wb.Worksheets[sheetIndex].Cells[indexRow, indexSub + 1].Style = s;
                    wb.Worksheets[sheetIndex].Cells[indexRow, indexSub + 2].Style = s;


                    if (obj != null){
                        wb.Worksheets[sheetIndex].Cells[indexRow, indexSub].PutValue(obj.Avg);
                        wb.Worksheets[sheetIndex].Cells[indexRow, indexSub + 1].PutValue(obj.Rank);
                        wb.Worksheets[sheetIndex].Cells[indexRow, indexSub + 2].PutValue(obj.Level);
                    }
                    indexRow++;
                }
                if (indexRow > 3){
                    Range test = wb.Worksheets[sheetIndex].Cells.CreateRange(3, 1, indexRow - 3, 1);
                    test.Style = s3;
                }
         
               //每5格劃分隔線
               for (int i = 2; i <= indexRow; i += 5) {
                   Range target = wb.Worksheets[sheetIndex].Cells.CreateRange(i, 0, 1, indexSub + 3);
                   target.Style = s2;
               }
			   for (int j = 2; j <= indexRow; j += 5) {
				   wb.Worksheets[sheetIndex].Cells[j, 1].Style = s4;
			   }
            }

            wb.Worksheets.RemoveAt(0);
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
			save.FileName = comBoxSchoolYear.Text + "." + comBoxSemester.Text + "班級評量成績單.xls";
            save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                try {
                    wb.Save(save.FileName, Aspose.Cells.FileFormatType.Excel2003);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch {
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
