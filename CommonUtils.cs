using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Diagnostics;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelTests {
    public class CommonUtils {
        public static void SwapRowCol() {            
            Excel.Range rg = Globals.ThisWorkbook.Application.Selection;

            HashSet<string> changed = new HashSet<string>();
            foreach (Excel.Range r in rg) {
                if (r.Row == r.Column || changed.Contains(r.Address)) continue;

                Excel.Range toChange = Globals.ThisWorkbook.Application.ActiveSheet.Cells[r.Column, r.Row];
                var val = toChange.Value;
                toChange.Value = r.Value;
                r.Value = val;

                changed.Add(toChange.Address);                
            }
        }
    }

}