using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelTests {
    public partial class HelloRibbon {
        private void HelloRibbon_Load(object sender, RibbonUIEventArgs e) {

        }

        private void tb1_click(object sender, RibbonControlEventArgs e) {
            toggleButton1.Checked = !toggleButton1.Checked;
            //System.Windows.Forms.MessageBox.Show("clicked!!");
        }

        private void b1_click(object sender, RibbonControlEventArgs e) {
            toggleButton1.Checked = !toggleButton1.Checked;
        }
    }
}
