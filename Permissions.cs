﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassExamSemester
{
    class Permissions
    {
        public static bool 班級評量成績單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級評量成績單].Executable;
            }
        }

        public static string 班級評量成績單 = "ClassExamSemester-{0183C5AB-BD58-4468-BBC6-D0AD48993859}";
    }
}
