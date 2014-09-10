using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelTests
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            //Microsoft.Office.Tools.Excel.NamedRange nr =
            //    this.Controls.AddNamedRange(this.Range["A2"], "NamedRange1");
            //nr.Value2 = "This text was added by using code";
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_click);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.Sheet1_SelectionChange);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void button1_click(object sender, EventArgs e) {
            this.webBrowser1.Url = new Uri("http://m.baidu.com");
            CommonUtils.SwapRowCol();
            Globals.ThisWorkbook.Save();
            MessageBox.Show("saved!");
        }

        private void Sheet1_SelectionChange(Excel.Range Target) {
            //foreach (Excel.Range c in Target.Cells) {
            //    var hehe = c.get_Address(false, false);
            //    var hehe2 = c.get_AddressLocal();
            //    var hehe3 = c.Address[false, true];
            //    var row = c.Row;
            //    var col = c.Column;
            //    System.Diagnostics.Debug.Print(hehe);
            //}
            //System.Diagnostics.Debug.Print("-------");

            //Excel.Range r1 = this.Cells[1, 2];
            //var nn = r1.Address;
            //var nn2 = r1.Text;
            //Debug.Print("haha2");            

        }
    }
}
