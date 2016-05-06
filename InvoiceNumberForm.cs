using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SerialNumberRetrieval
{
    public partial class InvoiceNumberForm : Form
    {
        public InvoiceNumberForm()
        {
            InitializeComponent();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            String invoiceNum = invoiceNumber.Text;
            if(invoiceNum == null)
            {
                MessageBox.Show("Please enter valid Invoice Number.");
                return;
            }
            else
            {
                Globals.ThisAddIn.runDataRetreival(invoiceNum);
            }
        }
    }
}
