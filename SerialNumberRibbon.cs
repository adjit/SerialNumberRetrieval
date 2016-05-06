using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace SerialNumberRetrieval
{
    public partial class SerialNumberRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void getSerialNumbersButton_Click(object sender, RibbonControlEventArgs e)
        {
            InvoiceNumberForm inf = new InvoiceNumberForm();
            inf.Show();
        }
    }
}
