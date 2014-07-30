using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassExamSemester
{
    class StudentObj
    {
        public string StudentID, ClassID,Level;
        public int Rank;
        public Dictionary<string, decimal> Subject_Score;
        public Dictionary<string, decimal> Subject_Credit;

        public StudentObj(DataRow row)
        {
            Subject_Score = new Dictionary<string,decimal>();
            Subject_Credit = new Dictionary<string,decimal>();

            StudentID = row["ref_student_id"] + "";
            ClassID = row["class_id"] + "";

            Level = "";
        }

        public void LoadData(DataRow row)
        {
            string subject = row["subject"] + "";

            decimal score_d = 0;
            decimal score = decimal.TryParse(row["score"] + "", out score_d) ? score_d : 0;
            decimal credit_d = 0;
            decimal credit = decimal.TryParse(row["credit"] + "", out credit_d) ? credit_d : 0;

            if (!Subject_Score.ContainsKey(subject))
                Subject_Score.Add(subject, score);

            if (!Subject_Credit.ContainsKey(subject))
                Subject_Credit.Add(subject, credit);

            string domain = row["group"] + "";
            string level = row["level"] + "";

            if (domain == "Chinese")
                Level = "Level " + level;
        }

        public decimal Avg
        {
            get
            {
                decimal count = 0;
                decimal total = 0;
                foreach(string subj in Subject_Score.Keys)
                {
                    decimal score = Subject_Score[subj];
                    decimal credit = Subject_Credit[subj];

                    count += credit;
                    total += score * credit;
                }

                //Subject_Credit.Values.Select(x => x * x).ToList().Sum();
                if (count > 0)
                    return Math.Round(total / count, 2, MidpointRounding.AwayFromZero);
                else
                    return 0;
            }
        }
    }
}
