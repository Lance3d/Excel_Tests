using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

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
            MessageBox.Show("hello world 2");
        }

        private void button3_click(object sender, EventArgs e) {
            MessageBox.Show("hello world 3");
        }
    }
}
