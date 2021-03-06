﻿using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassExamSemester
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
			item1["報表"]["成績相關報表"]["班級評量成績單"].Enable = false;
            item1["報表"]["成績相關報表"]["班級評量成績單"].Click += delegate
            {
                frm_printsetup form = new frm_printsetup(K12.Presentation.NLDPanels.Class.SelectedSource);
                form.ShowDialog();
            };

			K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
			{
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.班級評量成績單權限)
				{
                    item1["報表"]["成績相關報表"]["班級評量成績單"].Enable = true;
				}
				else
                    item1["報表"]["成績相關報表"]["班級評量成績單"].Enable = false;
			};

            //權限設定
            Catalog permission = RoleAclSource.Instance["班級"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.班級評量成績單, "班級評量成績單"));

        }
    }
}
