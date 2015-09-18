using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTests {
    partial class HelloControl : UserControl {
        public HelloControl() {
            InitializeComponent();
        }

        private void button1_click(object sender, EventArgs e) {
            CommonUtils.SwapRowCol();
        }

        private void HelloControl_Load(object sender, EventArgs e) {

        }

        private void button2_Click(object sender, EventArgs e) {
            //MessageBox.Show("hello world 2");

            // 功能：从另外一个表中找数据，然后把它前一列的数据填到自己的前一列
            Excel.Range reference = Globals.Sheet2.Range["B2", "B113"];
            Excel.Range rg = Globals.ThisWorkbook.Application.Selection;
            

            HashSet<string> changed = new HashSet<string>();
            foreach (Excel.Range r in rg) {
                Excel.Range refer = reference.Find(r.Value);
                if (refer != null) {
                    Excel.Range leftCell = reference.Worksheet.Cells[refer.Row, refer.Column - 1];
                    if (leftCell.Value != null) {
                        int val = (int)leftCell.Value;
                        //Debug.WriteLine(val);
                        Excel.Range targetCell = r.Worksheet.Cells[r.Row, r.Column - 1];
                        targetCell.Value = val;
                    }
                }
            }
        }

        private void button3_click(object sender, EventArgs e) {
            MessageBox.Show("hello world 3");
        }
    }
}
