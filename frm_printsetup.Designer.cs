namespace ClassExamSemester
{
    partial class frm_printsetup
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_yes = new DevComponents.DotNetBar.ButtonX();
            this.comBoxSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.comBoxSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btn_no = new DevComponents.DotNetBar.ButtonX();
            this.labSchoolYear = new DevComponents.DotNetBar.LabelX();
            this.labSemester = new DevComponents.DotNetBar.LabelX();
            this.labExam = new DevComponents.DotNetBar.LabelX();
            this.comBoxPeriod = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.SuspendLayout();
            // 
            // btn_yes
            // 
            this.btn_yes.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btn_yes.BackColor = System.Drawing.Color.Transparent;
            this.btn_yes.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btn_yes.Location = new System.Drawing.Point(236, 151);
            this.btn_yes.Name = "btn_yes";
            this.btn_yes.Size = new System.Drawing.Size(75, 23);
            this.btn_yes.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btn_yes.TabIndex = 0;
            this.btn_yes.Text = "確定";
            this.btn_yes.Click += new System.EventHandler(this.btn_yes_Click);
            // 
            // comBoxSchoolYear
            // 
            this.comBoxSchoolYear.DisplayMember = "Text";
            this.comBoxSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comBoxSchoolYear.FormattingEnabled = true;
            this.comBoxSchoolYear.ItemHeight = 19;
            this.comBoxSchoolYear.Location = new System.Drawing.Point(89, 38);
            this.comBoxSchoolYear.Name = "comBoxSchoolYear";
            this.comBoxSchoolYear.Size = new System.Drawing.Size(121, 25);
            this.comBoxSchoolYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comBoxSchoolYear.TabIndex = 1;
            // 
            // comBoxSemester
            // 
            this.comBoxSemester.DisplayMember = "Text";
            this.comBoxSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comBoxSemester.FormattingEnabled = true;
            this.comBoxSemester.ItemHeight = 19;
            this.comBoxSemester.Location = new System.Drawing.Point(294, 38);
            this.comBoxSemester.Name = "comBoxSemester";
            this.comBoxSemester.Size = new System.Drawing.Size(80, 25);
            this.comBoxSemester.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comBoxSemester.TabIndex = 2;
            // 
            // btn_no
            // 
            this.btn_no.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btn_no.BackColor = System.Drawing.Color.Transparent;
            this.btn_no.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btn_no.Location = new System.Drawing.Point(317, 151);
            this.btn_no.Name = "btn_no";
            this.btn_no.Size = new System.Drawing.Size(75, 23);
            this.btn_no.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btn_no.TabIndex = 3;
            this.btn_no.Text = "取消";
            this.btn_no.Click += new System.EventHandler(this.btn_no_Click);
            // 
            // labSchoolYear
            // 
            this.labSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labSchoolYear.BackgroundStyle.Class = "";
            this.labSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labSchoolYear.Location = new System.Drawing.Point(26, 38);
            this.labSchoolYear.Name = "labSchoolYear";
            this.labSchoolYear.Size = new System.Drawing.Size(57, 23);
            this.labSchoolYear.TabIndex = 4;
            this.labSchoolYear.Text = "學年度";
            // 
            // labSemester
            // 
            this.labSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labSemester.BackgroundStyle.Class = "";
            this.labSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labSemester.Location = new System.Drawing.Point(245, 38);
            this.labSemester.Name = "labSemester";
            this.labSemester.Size = new System.Drawing.Size(43, 23);
            this.labSemester.TabIndex = 5;
            this.labSemester.Text = "學期";
            // 
            // labExam
            // 
            this.labExam.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labExam.BackgroundStyle.Class = "";
            this.labExam.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labExam.Location = new System.Drawing.Point(26, 96);
            this.labExam.Name = "labExam";
            this.labExam.Size = new System.Drawing.Size(57, 29);
            this.labExam.TabIndex = 6;
            this.labExam.Text = "試別";
            // 
            // comBoxPeriod
            // 
            this.comBoxPeriod.DisplayMember = "Text";
            this.comBoxPeriod.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comBoxPeriod.FormattingEnabled = true;
            this.comBoxPeriod.ItemHeight = 19;
            this.comBoxPeriod.Location = new System.Drawing.Point(89, 96);
            this.comBoxPeriod.Name = "comBoxPeriod";
            this.comBoxPeriod.Size = new System.Drawing.Size(285, 25);
            this.comBoxPeriod.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comBoxPeriod.TabIndex = 7;
            this.comBoxPeriod.Text = "Midterm";
            // 
            // frm_printsetup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(404, 186);
            this.Controls.Add(this.comBoxPeriod);
            this.Controls.Add(this.labExam);
            this.Controls.Add(this.labSemester);
            this.Controls.Add(this.labSchoolYear);
            this.Controls.Add(this.btn_no);
            this.Controls.Add(this.comBoxSemester);
            this.Controls.Add(this.comBoxSchoolYear);
            this.Controls.Add(this.btn_yes);
            this.DoubleBuffered = true;
            this.Name = "frm_printsetup";
            this.Text = "班級學期成績單";
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btn_yes;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comBoxSchoolYear;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comBoxSemester;
        private DevComponents.DotNetBar.ButtonX btn_no;
        private DevComponents.DotNetBar.LabelX labSchoolYear;
        private DevComponents.DotNetBar.LabelX labSemester;
        private DevComponents.DotNetBar.LabelX labExam;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comBoxPeriod;
    }
}