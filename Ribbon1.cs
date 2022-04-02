using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MyExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonSave_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show($"id:{this.RibbonId} Type:{this.RibbonType} Save clicked", this.Name, System.Windows.Forms.MessageBoxButtons.OK);
        }
    }
}
